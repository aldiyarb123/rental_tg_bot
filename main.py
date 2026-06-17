import os
import csv
import time
import json
import math
import random
import logging
import argparse
from dataclasses import dataclass
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Dict, List, Tuple, Any, Optional

import requests
from requests.exceptions import RequestException
from dotenv import load_dotenv

import gspread
from google.oauth2.service_account import Credentials

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, ContextTypes, filters

# ------------------- CONFIG -------------------
from pathlib import Path
load_dotenv(dotenv_path=Path(__file__).with_name(".env"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger("fleet-reports")

DUBAI_TZ = ZoneInfo("Asia/Dubai")
FLEET_BASE = "https://fleet-api.taxi.yandex.net"

BOT_TOKEN = os.getenv("BOT_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")

YANDEX_CLIENT_ID = os.getenv("YANDEX_CLIENT_ID")   # X-Client-ID
YANDEX_API_KEY = os.getenv("YANDEX_API_KEY")       # X-API-Key
PARK_ID = os.getenv("PARK_ID")

GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_SA_JSON = os.getenv("GOOGLE_SA_JSON")

# процент дохода таксопарка (например 20 = парк получает 20% от суммы заказов)
PARK_COMMISSION_PERCENT = float(os.getenv("PARK_COMMISSION_PERCENT", "20"))

# Для продакшна: лимит Яндекса, судя по твоим логам
ORDERS_PAGE_LIMIT = 500

# Сколько максимум страниц качаем, чтобы не улететь в бесконечность
MAX_PAGES = 400  # 400 * 500 = 200k заказов. Если больше — значит фильтр не работает.
MAX_ORDERS_PER_DAY = 8000  # жёсткий предохранитель: если больше, значит критерий дня/статуса почти наверняка не тот

missing = [k for k, v in {
    "BOT_TOKEN": BOT_TOKEN,
    "YANDEX_CLIENT_ID": YANDEX_CLIENT_ID,
    "YANDEX_API_KEY": YANDEX_API_KEY,
    "PARK_ID": PARK_ID,
    "GOOGLE_SHEET_ID": GOOGLE_SHEET_ID,
    "GOOGLE_SA_JSON": GOOGLE_SA_JSON,
}.items() if not v]
if missing:
    raise SystemExit(f"Missing .env variables: {', '.join(missing)}")

# CHAT_ID можно не требовать для кнопок/команд, но оставлю если ты хочешь авто-отчёт в группу
CHAT_ID_INT: Optional[int] = None
if CHAT_ID and isinstance(CHAT_ID, str) and CHAT_ID.lstrip("-").isdigit():
    CHAT_ID_INT = int(CHAT_ID)

# ------------------- GOOGLE SHEETS -------------------
GS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def gs_client():
    # Если GOOGLE_SA_JSON содержит JSON-строку (Railway) — парсим напрямую
    # Если это путь к файлу (локально) — читаем файл
    sa_value = GOOGLE_SA_JSON.strip()
    if sa_value.startswith("{"):
        info = json.loads(sa_value)
        creds = Credentials.from_service_account_info(info, scopes=GS_SCOPES)
    else:
        creds = Credentials.from_service_account_file(sa_value, scopes=GS_SCOPES)
    return gspread.authorize(creds)

def gs_get_or_create_ws(sh, title: str, rows: int = 2000, cols: int = 20):
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
    return ws

def ws_clear_and_write(ws, values: list[list]):
    ws.clear()
    if values:
        ws.update(values=values, range_name="A1", value_input_option="USER_ENTERED")

# ------------------- YANDEX -------------------
def fleet_headers():
    return {
        "X-Client-ID": YANDEX_CLIENT_ID,
        "X-API-Key": YANDEX_API_KEY,
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Accept-Language": "ru",
    }

def post_with_retry(session: requests.Session, url: str, payload: dict, max_retries: int = 8, timeout: int = 60):
    base_delay = 1.0
    last = None

    for attempt in range(1, max_retries + 1):
        try:
            r = session.post(url, headers=fleet_headers(), json=payload, timeout=timeout)
            last = r
        except RequestException as e:
            if attempt == max_retries:
                raise
            sleep_s = min(base_delay * (2 ** (attempt - 1)), 30.0) + random.uniform(0.2, 0.8)
            log.warning("YANDEX RETRY (network): %s | sleeping %.1fs", e, sleep_s)
            time.sleep(sleep_s)
            continue

        if r.status_code < 400:
            return r

        # 4xx кроме 429 — это ошибка запроса/прав, ретраями не лечится
        if 400 <= r.status_code < 500 and r.status_code != 429:
            return r

        retry_after = r.headers.get("Retry-After")
        if retry_after:
            try:
                base_delay = max(base_delay, float(retry_after))
            except ValueError:
                pass

        if attempt < max_retries:
            sleep_s = min(base_delay * (2 ** (attempt - 1)), 30.0) + random.uniform(0.2, 0.9)
            log.warning("YANDEX RETRY: HTTP %s (attempt %s/%s) sleeping %.1fs",
                        r.status_code, attempt, max_retries, sleep_s)
            time.sleep(sleep_s)
            continue

        return r

    return last

def iso_dubai(dt: datetime) -> str:
    """RFC3339 с таймзоной +04:00 (не Z)"""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=DUBAI_TZ)
    return dt.astimezone(DUBAI_TZ).isoformat()

def dubai_day_range(day_date) -> Tuple[str, str, datetime, datetime]:
    start_dt = datetime(day_date.year, day_date.month, day_date.day, 0, 0, 0, tzinfo=DUBAI_TZ)
    end_dt   = datetime(day_date.year, day_date.month, day_date.day, 23, 59, 59, tzinfo=DUBAI_TZ)
    return iso_dubai(start_dt), iso_dubai(end_dt), start_dt, end_dt

def extract_booked_at(order: dict) -> Optional[str]:
    # В разных ответах бывает order.booked_at или прямо booked_at
    if isinstance(order.get("order"), dict) and order["order"].get("booked_at"):
        return order["order"]["booked_at"]
    if order.get("booked_at"):
        return order["booked_at"]
    return None

def parse_dt(s: str) -> Optional[datetime]:
    if not s:
        return None
    try:
        # поддержим Z
        if s.endswith("Z"):
            s = s.replace("Z", "+00:00")
        return datetime.fromisoformat(s)
    except Exception:
        return None

def normalize_status(o: dict) -> str:
    """Нормализуем статус заказа из верхнего уровня или из вложенного order."""
    st = (o.get("status") or "").lower().strip()
    if not st and isinstance(o.get("order"), dict):
        st = str(o["order"].get("status") or "").lower().strip()
    return st


def extract_effective_dt(order: dict) -> Optional[datetime]:
    """Для завершённых заказов используем ended_at как более точный критерий попадания в день.
    Если ended_at нет — падаем обратно на booked_at.
    """
    ended_at = order.get("ended_at")
    if isinstance(order.get("order"), dict):
        ended_at = ended_at or order["order"].get("ended_at")
    if isinstance(ended_at, str) and ended_at:
        dt = parse_dt(ended_at)
        if dt is not None:
            return dt

    booked_at = extract_booked_at(order)
    if booked_at:
        return parse_dt(booked_at)
    return None

def orders_list(
    session: requests.Session,
    time_from_iso: str,
    time_to_iso: str,
    limit: int,
    offset: int = 0,
    cursor: Optional[str] = None,
) -> dict:
    """
    Fleet API на практике часто работает стабильнее с cursor-пагинацией (в ответе приходит поле `cursor`).

    ВАЖНО:
    - `order` должен быть ВНУТРИ `park` (query.park.order)
    - фильтр `status` у некоторых парков/ключей может игнорироваться API (мы дополнительно фильтруем локально)
    """
    url = f"{FLEET_BASE}/v1/parks/orders/list"
    limit = min(int(limit), ORDERS_PAGE_LIMIT)

    payload: Dict[str, Any] = {
        "limit": limit,
        "query": {
            "park": {
                "id": str(PARK_ID),
                "order": {
                    "booked_at": {"from": time_from_iso, "to": time_to_iso},
                },
            }
        },
        "fields": {
            # поля верхнего уровня (в ответах реально приходят именно так)
            "order": [
            "id",
            "status",
            "booked_at",
            "accepted_at",
            "started_at",
            "ended_at"
            ],
            "driver_profile": ["id", "first_name", "last_name", "callsign", "name","balance"],
            "car": ["callsign", "brand_model", "license"],
            "payment": ["amount", "type"],
            "fees": [],
            "cost": [],
            "price": [],
        },
    }

    # пагинация: предпочтительно cursor, иначе offset
    if cursor:
        payload["cursor"] = cursor
    else:
        payload["offset"] = int(offset)

    if os.getenv("DEBUG_PAYLOAD") == "1":
        log.warning(
            "ORDERS PAYLOAD (cursor=%s, offset=%s, limit=%s): %s",
            (cursor[:24] + "...") if cursor else None,
            offset,
            limit,
            payload,
        )

    r = post_with_retry(session, url, payload)
    if r.status_code >= 400:
        log.error("YANDEX ORDERS ERROR: %s", r.status_code)
        log.error(r.text)
        if r.status_code == 400:
            log.error("ORDERS PAYLOAD CAUSED 400: %s", payload)
    r.raise_for_status()
    return r.json()

def orders_list_all(
    session: requests.Session,
    time_from_iso: str,
    time_to_iso: str,
    from_dt: datetime,
    to_dt: datetime,
) -> List[dict]:
    """Качаем заказы за окно дня.

    На практике для корректного отчёта по "вчера" нам нужно:
    - брать только уникальные заказы
    - фильтровать локально только завершённые статусы
    - для попадания в день использовать ended_at, а не booked_at (если ended_at есть)
    - ходить по cursor, а не по offset
    """
    kept: List[dict] = []
    seen_ids = set()
    pages = 0
    cursor: Optional[str] = None

    while True:
        pages += 1
        if pages > MAX_PAGES:
            raise RuntimeError(
                f"Слишком много страниц (> {MAX_PAGES}). Похоже, фильтр/пагинация не работает."
            )

        data = orders_list(
            session,
            time_from_iso,
            time_to_iso,
            limit=ORDERS_PAGE_LIMIT,
            offset=(pages - 1) * ORDERS_PAGE_LIMIT,
            cursor=cursor,
        )

        items = data.get("orders") or data.get("result", {}).get("orders") or []
        if not isinstance(items, list):
            items = []

        next_cursor = data.get("cursor")
        new_unique = 0
        kept_now = 0

        for o in items:
            oid = o.get("id")
            if oid is None and isinstance(o.get("order"), dict):
                oid = o["order"].get("id")

            if oid is not None and oid in seen_ids:
                continue
            if oid is not None:
                seen_ids.add(oid)
            new_unique += 1

            st = normalize_status(o)
            if st not in ("complete", "completed"):
                continue

            dt = extract_effective_dt(o)
            if dt is None:
                continue

            dt_local = dt.astimezone(DUBAI_TZ)
            if from_dt <= dt_local <= to_dt:
                kept.append(o)
                kept_now += 1

        log.info(
            "Fetched %s orders (page=%s) new_unique=%s kept_in_window=%s total_kept=%s",
            len(items),
            pages,
            new_unique,
            kept_now,
            len(kept),
        )

        if len(kept) >= MAX_ORDERS_PER_DAY:
            raise RuntimeError(
                f"Слишком много завершённых заказов за сутки (>= {MAX_ORDERS_PER_DAY}). "
                f"Проверь критерий отчёта. total_kept={len(kept)}"
            )

        if len(items) == 0:
            break

        if new_unique == 0:
            log.warning("Pagination seems stuck (repeats). Stopping to avoid infinite fetch.")
            break

        if not next_cursor:
            break

        cursor = next_cursor
        time.sleep(0.35)

    return kept


def fetch_driver_balances(session: requests.Session) -> Dict[str, float]:
    """
    Загружает текущий баланс всех водителей через v1/parks/driver-profiles/list.
    Возвращает dict: "Фамилия Имя Отчество" -> balance (float)
    """
    url = f"{FLEET_BASE}/v1/parks/driver-profiles/list"
    balances: Dict[str, float] = {}
    limit = 200
    offset = 0

    while True:
        payload = {
            "fields": {
                "account": ["balance"],
                "driver_profile": ["first_name", "last_name", "middle_name", "id"],
            },
            "limit": limit,
            "offset": offset,
            "query": {
                "park": {"id": str(PARK_ID)}
            },
        }

        r = post_with_retry(session, url, payload)
        if r is None or r.status_code >= 400:
            log.warning("fetch_driver_balances error: %s", getattr(r, "text", "no response"))
            break

        data = r.json()
        items = data.get("driver_profiles") or []
        if not items:
            break

        for item in items:
            dp = item.get("driver_profile") or {}
            account = item.get("accounts") or item.get("account") or []
            balance_val = None

            # accounts может быть списком или одним объектом
            if isinstance(account, list):
                for acc in account:
                    if "balance" in acc:
                        try:
                            balance_val = float(acc["balance"])
                            break
                        except Exception:
                            pass
            elif isinstance(account, dict):
                if "balance" in account:
                    try:
                        balance_val = float(account["balance"])
                    except Exception:
                        pass

            if balance_val is None:
                continue

            first = (dp.get("first_name") or "").strip()
            last = (dp.get("last_name") or "").strip()
            middle = (dp.get("middle_name") or "").strip()
            fio = " ".join([x for x in [last, first, middle] if x]).strip()
            if fio:
                balances[fio] = balance_val

        if len(items) < limit:
            break
        offset += limit
        time.sleep(0.2)

    return balances

# ------------------- TARIFFS -------------------
BULK_TARIFF_IMPORT = [
    ("Абдукадыров Дониёр Абдукаримович", "Штатный", 5.0, None),
    ("Алиев Сулейман Рутамович", "Штатный", 5.0, None),
    ("Амарбеков Сухраб Султанмуратович", "Штатный", 5.0, None),
    ("Аяпов Динмухамед Абуталипулы", "Штатный", 5.0, None),
    ("Бегимбетов Арнур Абдикалиевич", "Штатный", 5.0, None),
    ("Валиев Шахмурат Аркинжанович", "Штатный", 5.0, None),
    ("Диханбай Кайрат Қуанышұлы", "Штатный", 5.0, None),
    ("Дөңес Ерсұлтан Дөңесұлы", "Штатный", 5.0, None),
    ("Әли Әділет Бактиярұлы", "Штатный", 5.0, None),
    ("Еркетайұлы Ерсін", "Штатный", 5.0, None),
    ("Ибодов Рахматулла", "Штатный", 5.0, None),
    ("Исламов Хамиджан Нуриевич", "Штатный", 5.0, None),
    ("Кадыров Ернұр Ерланұлы", "Штатный", 5.0, None),
    ("Касымов Эдуард Маратович", "Штатный", 5.0, None),
    ("Ким Игорь Юрьевич", "Штатный", 5.0, None),
    ("Ким Денис Витальевич", "Штатный", 5.0, None),
    ("Косанов Марат Гусманович", "Штатный", 5.0, None),
    ("Мақсұт Темірлан Есенгелдіұлы", "Штатный", 5.0, None),
    ("Мансуров Аслан Келисович", "Штатный", 5.0, None),
    ("Махамбетов Мұқасан Нұрқас Уғли", "Штатный", 5.0, None),
    ("Мухамеджан Нурбол Ахметиллаұлы", "Штатный", 5.0, None),
    ("Нұрғожа Алмас Алмас", "Штатный", 5.0, None),
    ("Пруглов Олег Витальевич", "Штатный", 5.0, None),
    ("Рулёв Даниэль Владмирович", "Штатный", 5.0, None),
    ("Садвахасов Қанат Шамбилұлы", "Штатный", 5.0, None),
    ("Спанкулов Алдияр Замирович", "Штатный", 5.0, None),
    ("Спатаев Талгат Каримович", "Штатный", 5.0, None),
    ("Султанбеков Тимур Бахытұлы", "Штатный", 5.0, None),
    ("Сунакбаева Жадыра Рахимжановна", "Штатный", 5.0, None),
    ("Тілен Ержігіт Аманұлы", "Штатный", 5.0, None),
    ("Тутешов Меңдібай Ғалымжанұлы", "Штатный", 5.0, None),
    ("Тюлеев Рустам Оспанович", "Штатный", 5.0, None),
    ("Феллер Вячеслав Анатольевич", "Штатный", 5.0, None),
    ("Чокпаров Руслан Назымович", "Штатный", 5.0, None),
    ("Шавкатов Агабей Юнусович", "Штатный", 5.0, None),
    ("Эрканбаев Сардор Артыкалиевич", "Штатный", 5.0, None),
    ("Айдаров Дильшат Равхатович", "Тариф1", 2.0, None),
    ("Алмасұлы Медеу Медеу", "Тариф1", 2.0, None),
    ("Джаксыбаев Ерлан Курмангазыевич", "Тариф1", 2.0, None),
    ("Джапарханов Таш Дылшатович", "Тариф1", 2.0, None),
    ("Досмамбетов Азат Кайратович", "Тариф1", 2.0, None),
    ("Мауыт Даниал Ержанұлы", "Тариф1", 2.0, None),
    ("Рахимгалиев Салимгерей Галимжанович", "Тариф1", 2.0, None),
    ("Айтжанов Мирас Бахтиярұлы", "АРА", None, '2026-01-05'),
    ("Алпысбаева Динара Агибаевна", "АРА", None, '2026-01-13'),
    ("Ганиев Ринат Шамильевич", "АРА", None, '2026-01-23'),
    ("Егоров Фархат Александрович", "АРА", None, '2026-02-03'),
    ("Жалгасбаев Алтынбек Жетилдикулы", "АРА", None, '2026-04-05'),
    ("Жұмағұл Мәдияр Жақсыбайұлы", "АРА", None, '2026-05-26'),
    ("Жумахан Сержан Бақтығалиұлы", "АРА", None, '2026-02-20'),
    ("Исаба Дамир Серiкулы", "АРА", None, '2026-05-17'),
    ("Қадырұлы Арман Арман", "АРА", None, '2025-11-23'),
    ("Капар Ернур Ерболулы", "АРА", None, '2026-04-11'),
    ("Қасымбеков Серік Мұхаметұлы", "АРА", None, '2026-03-07'),
    ("Колыхалова Кристина Владимировна", "АРА", None, '2026-03-18'),
    ("Қыдырәлі Дінмұхамбет Манасулы", "АРА", None, '2026-04-23'),
    ("Мухамметхали Әли Жексенбайұлы", "АРА", None, '2026-03-27'),
    ("Мырзахметов Медет Прiмбетулы", "АРА", None, '2026-04-28'),
    ("Өзгелді Нұрсәулет Ғаниұлы", "АРА", None, '2026-03-18'),
    ("Пак Андрей Афанасьевич", "АРА", None, '2026-04-28'),
    ("Рауфов Даниял", "АРА", None, None),
    ("Серікбай Айдын Серікбайұлы", "АРА", None, '2025-11-21'),
    ("Сулейманов Рамиль Абдухахарович", "АРА", None, '2026-04-29'),
    ("Хамитов Ерасыл Серiкулы", "АРА", None, None),
    ("Шатаев Шохан Аубакирович", "АРА", None, '2026-05-26'),
]

TARIFF_SHEET_NAME = "Тарифы"
TARIFF_HEADER = ["ФИО водителя", "Тариф", "Процент", "Дата начала АРА"]

def load_tariffs() -> Dict[str, dict]:
    """
    Загружает справочник тарифов из Google Sheets (лист "Тарифы").
    Возвращает dict: fio -> {"tariff": str, "percent": float, "ara_start": date|None}
    """
    import datetime as dt_mod
    result = {}
    try:
        gc = gs_client()
        sh = gc.open_by_key(GOOGLE_SHEET_ID)
        try:
            ws = sh.worksheet(TARIFF_SHEET_NAME)
        except gspread.WorksheetNotFound:
            return result
        rows = ws.get_all_records()
        for r in rows:
            fio = str(r.get("ФИО водителя") or "").strip()
            if not fio:
                continue
            tariff = str(r.get("Тариф") or "").strip()
            percent_raw = r.get("Процент")
            try:
                percent = float(percent_raw) if percent_raw not in (None, "", "авто") else None
            except Exception:
                percent = None
            ara_start_raw = str(r.get("Дата начала АРА") or "").strip()
            ara_start = None
            if ara_start_raw:
                try:
                    ara_start = dt_mod.date.fromisoformat(ara_start_raw)
                except Exception:
                    ara_start = None
            result[fio] = {"tariff": tariff, "percent": percent, "ara_start": ara_start}
    except Exception as e:
        log.warning("load_tariffs error: %s", e)
    return result


def get_driver_percent(fio: str, tariffs: Dict[str, dict], for_date=None) -> float:
    """
    Возвращает процент комиссии парка для конкретного водителя.
    Если тариф не задан — используется PARK_COMMISSION_PERCENT (дефолт).
    Для АРА: 1% первые 14 дней с ara_start, потом 2%.
    """
    import datetime as dt_mod
    info = tariffs.get(fio)
    if not info:
        return PARK_COMMISSION_PERCENT

    tariff = (info.get("tariff") or "").strip().upper()
    if tariff == "АРА":
        ara_start = info.get("ara_start")
        if ara_start:
            check_date = for_date or dt_mod.date.today()
            days_passed = (check_date - ara_start).days
            return 1.0 if days_passed < 14 else 2.0
        return 1.0

    if info.get("percent") is not None:
        return info["percent"]

    return PARK_COMMISSION_PERCENT


def set_tariff(fio: str, tariff: str, percent: Optional[float] = None, ara_start=None) -> None:
    """Создаёт/обновляет запись в листе 'Тарифы' для водителя."""
    gc = gs_client()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    ws = gs_get_or_create_ws(sh, TARIFF_SHEET_NAME, rows=500, cols=10)

    values = ws.get_all_values()
    # Проверяем, есть ли заголовок (первая строка должна совпадать с TARIFF_HEADER)
    has_header = bool(values) and len(values[0]) >= 1 and values[0][0].strip() == TARIFF_HEADER[0]
    if not has_header:
        # Вставляем заголовок первой строкой, сдвигая всё остальное вниз
        ws.insert_row(TARIFF_HEADER, index=1, value_input_option="USER_ENTERED")
        values = ws.get_all_values()

    header = values[0]
    fio_col = 0
    row_idx = None
    for i, row in enumerate(values[1:], start=2):
        if len(row) > fio_col and row[fio_col].strip() == fio.strip():
            row_idx = i
            break

    percent_str = "" if percent is None else str(percent)
    ara_str = "" if not ara_start else str(ara_start)
    new_row = [fio, tariff, percent_str, ara_str]

    if row_idx:
        ws.update(values=[new_row], range_name=f"A{row_idx}:D{row_idx}", value_input_option="USER_ENTERED")
    else:
        ws.append_row(new_row, value_input_option="USER_ENTERED")


def calc_park_income_individual(agg: Dict[str, "DriverAgg"], tariffs: Dict[str, dict], for_date=None) -> float:
    """Считает доход парка с учётом индивидуальных процентов каждого водителя."""
    total = 0.0
    for fio, a in agg.items():
        pct = get_driver_percent(fio, tariffs, for_date)
        total += a.net * (pct / 100)
    return total


# ------------------- REPORT (без VLOOKUP) -------------------
def load_driver_directory() -> Dict[str, dict]:
    """
    Загружает справочник водителей из Google Sheets.
    Возвращает dict: callsign(lower) -> {fio, vu, balance}
    """
    gc = gs_client()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)

    # ⚠️ ВАЖНО: поставь точное название листа где у тебя ФИО
    ws = sh.worksheet("Trips Per Driver")  # если другое название — поменяй

    rows = ws.get_all_records()
    directory = {}

    for r in rows:
        callsign = str(
            r.get("Позывной")
            or r.get("callsign")
            or r.get("Номер")
            or r.get("Номер авто")
            or ""
        ).strip()

        fio = str(
            r.get("ФИО")
            or r.get("ФИО водителя")
            or ""
        ).strip()

        vu = str(
            r.get("Номер ВУ")
            or r.get("ВУ")
            or ""
        ).strip()

        balance = r.get("Баланс")

        if callsign:
            directory[callsign.lower()] = {
                "fio": fio,
                "vu": vu,
                "balance": balance,
            }

    return directory
@dataclass
class DriverAgg:
    fio: str
    callsign: str
    plate: str
    done: int
    net: float
    cash: float
    cashless: float
    line_seconds: float
    balance: float

def _get_payment_obj(o: dict) -> dict:
    if isinstance(o.get("payment"), dict):
        return o["payment"]
    if isinstance(o.get("order"), dict) and isinstance(o["order"].get("payment"), dict):
        return o["order"]["payment"]
    return {}

def safe_amount(o: dict) -> float:
    p = _get_payment_obj(o)

    # возможные места
    candidates = [
        p.get("amount"),
        p.get("total"),
        p.get("value"),
        (o.get("order") or {}).get("cost"),
        (o.get("order") or {}).get("price"),
        o.get("cost"),
        o.get("price"),
        o.get("amount"),
    ]

    for v in candidates:
        if v is None:
            continue
        if isinstance(v, dict):
            for k in ("amount", "total", "value", "price"):
                if v.get(k) is not None:
                    try:
                        return float(v[k])
                    except Exception:
                        pass
        else:
            try:
                return float(v)
            except Exception:
                pass
    return 0.0

def pay_type(o: dict) -> str:
    # Fleet API часто отдаёт payment_method на верхнем уровне
    top_level = str(o.get("payment_method") or "").lower().strip()
    if top_level:
        return top_level

    p = _get_payment_obj(o)
    t = p.get("type")
    if isinstance(t, dict):
        t = t.get("id") or t.get("name")
    return str(t).lower().strip() if t else ""


def fio_from_order(o: dict) -> Tuple[str, str]:
    # защита от некорректных данных API
    if not isinstance(o, dict):
        return "", ""

    dp = o.get("driver_profile") or {}
    car = o.get("car") or {}

    # В ответах Fleet API часто есть готовое поле name
    fio = str(dp.get("name") or "").strip()

    if not fio:
        fio = " ".join([x for x in [dp.get("first_name"), dp.get("last_name")] if x]).strip()

    callsign = str(dp.get("callsign") or car.get("callsign") or "").strip()

    # гарантируем что функция ВСЕГДА возвращает tuple
    return fio or "", callsign or ""


def plate_from_order(o: dict) -> str:
    car = o.get("car") or {}
    license_obj = car.get("license") or {}

    if isinstance(license_obj, dict):
        return str(license_obj.get("number") or license_obj.get("plate") or "").strip()

    if isinstance(license_obj, str):
        return license_obj.strip()

    return ""

def build_report_rows(day_str: str, orders: List[dict], directory: Dict[str, dict], balances: Optional[Dict[str, float]] = None) -> Tuple[List[List[Any]], Dict[str, DriverAgg]]:
    agg: Dict[str, DriverAgg] = {}
    balances = balances or {}

    for o in orders:
        result = fio_from_order(o)
        fio_api, callsign = result if isinstance(result, tuple) else ("", "")
        plate = plate_from_order(o)
        key_callsign = (callsign or "").lower()

        fio_final = fio_api
        vu_final = ""
        balance_final = ""

        if (not fio_final) and key_callsign in directory:
            fio_final = directory[key_callsign].get("fio", "") or ""
            vu_final = directory[key_callsign].get("vu", "") or ""
            balance_final = directory[key_callsign].get("balance", "") or ""

        # если даже справочник не помог — оставим callsign
        if not fio_final:
            fio_final = callsign or "неизвестно"

        key = fio_final  # группируем по ФИО (или callsign если ФИО нет)

        if key not in agg:
            agg[key] = DriverAgg(
                fio=fio_final,
                callsign=callsign,
                plate=plate,
                done=0,
                net=0.0,
                cash=0.0,
                cashless=0.0,
                line_seconds=0.0,
                balance=0.0,
            )

        amount = safe_amount(o)

        # -------- баланс водителя из API если есть --------
        driver_balance = None
        dp = o.get("driver_profile") or {}
        bal_val = dp.get("balance")
        if bal_val is not None:
            try:
                driver_balance = float(bal_val)
            except Exception:
                pass

        # -------- время на линии (приблизительно) --------
        start_time = o.get("accepted_at") or o.get("started_at")
        end_time = o.get("ended_at")

        if isinstance(o.get("order"), dict):
            start_time = start_time or o["order"].get("accepted_at") or o["order"].get("started_at")
            end_time = end_time or o["order"].get("ended_at")

        # fallback: используем events если API не отдаёт accepted_at
        if not start_time:
            events = o.get("events") or []
            for ev in events:
                if ev.get("order_status") in ("driving", "transporting"):
                    start_time = ev.get("event_at")
                    break

        duration_sec = 0
        if start_time and end_time:
            dt1 = parse_dt(start_time)
            dt2 = parse_dt(end_time)
            if dt1 and dt2:
                duration_sec = max(0, (dt2 - dt1).total_seconds())

        t = pay_type(o)

        agg[key].done += 1
        agg[key].net += amount
        agg[key].line_seconds += duration_sec
        if driver_balance is not None:
            agg[key].balance = driver_balance
        if "cash" in t or "нал" in t:
            agg[key].cash += amount
        else:
            agg[key].cashless += amount

    rows_sorted = sorted(agg.values(), key=lambda x: x.net, reverse=True)

    out: List[List[Any]] = []
    for r in rows_sorted:
        hours = round(r.line_seconds / 3600, 2)

        bal = balances.get(r.fio)
        if bal is None:
            bal = r.balance
        if not bal:
            drow = directory.get((r.callsign or "").lower(), {})
            bal = drow.get("balance", "")

        out.append([
            r.fio,
            r.plate,
            r.done,
            hours,
            round(r.net, 2),
            17000,
            round(r.net - 17000, 2),
            bal,
        ])

    return out, agg
def write_to_google_sheets(day_str: str, report_rows: List[List[Any]]):
    gc = gs_client()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)

    ws = gs_get_or_create_ws(sh, day_str, rows=max(2000, len(report_rows) + 10), cols=20)

    header = [
    "ФИО водителя",
    "Госномер",
    "Завершено заказов",
    "Время на линии",
    "Чистый заработок",
    "Цель",
    "Отклонение плана",
    "Баланс",
]

    values = [header] + report_rows
    ws_clear_and_write(ws, values)

def write_csv(day_str: str, report_rows: List[List[Any]]) -> str:
    filename = f"report_{day_str}.csv"
    with open(filename, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow([
        "ФИО водителя",
        "Номер ВУ",
        "Завершено заказов",
        "Время на линии",
        "Чистый заработок",
        "Цель",
        "Отклонение плана",
        "Баланс",
    ])
        w.writerows(report_rows)
    return filename

def get_month_totals_sync(day_date) -> tuple:
    """Считает накопленные поездки и выручку с начала текущего месяца по day_date включительно."""
    month_start = day_date.replace(day=1)
    total_orders = 0
    total_net = 0.0
    try:
        directory = load_driver_directory()
    except Exception:
        directory = {}
    try:
        with requests.Session() as session:
            d = month_start
            while d <= day_date:
                time_from, time_to, from_dt, to_dt = dubai_day_range(d)
                orders = orders_list_all(session, time_from, time_to, from_dt, to_dt)
                _, agg = build_report_rows(str(d), orders, directory)
                total_orders += sum(a.done for a in agg.values())
                total_net += sum(a.net for a in agg.values())
                d += timedelta(days=1)
    except Exception as e:
        log.warning("Month totals error: %s", e)
    return total_orders, total_net


def get_month_totals_from_sheets(day_date) -> tuple:
    """Суммирует поездки и выручку за текущий месяц из Google Sheets (листы по датам)."""
    import datetime as dt_mod
    try:
        gc = gs_client()
        sh = gc.open_by_key(GOOGLE_SHEET_ID)
        month_prefix = day_date.strftime("%Y-%m")
        total_orders = 0
        total_net = 0.0
        for ws in sh.worksheets():
            title = ws.title.strip()
            if not title.startswith(month_prefix):
                continue
            try:
                sheet_date = dt_mod.date.fromisoformat(title)
                if sheet_date > day_date:
                    continue
            except Exception:
                continue
            values = ws.get_all_values()
            for row in values[1:]:
                if len(row) >= 5:
                    try:
                        orders_val = int(row[2]) if str(row[2]).strip() else 0
                        net_str = str(row[4]).replace(",", "").replace(" ", "").strip()
                        net_val = float(net_str) if net_str else 0.0
                        total_orders += orders_val
                        total_net += net_val
                    except Exception:
                        pass
        return total_orders, total_net
    except Exception as e:
        log.warning("get_month_totals_from_sheets error: %s", e)
        return 0, 0.0


def format_quick_summary(day_str: str, agg: Dict[str, DriverAgg], month_orders: int = 0, month_net: float = 0.0, tariffs: Optional[Dict[str, dict]] = None) -> str:
    import datetime as dt_mod
    total_orders = sum(a.done for a in agg.values())
    total_net = sum(a.net for a in agg.values())
    if tariffs:
        try:
            report_date = dt_mod.date.fromisoformat(day_str)
        except Exception:
            report_date = dt_mod.date.today()
        park_income = calc_park_income_individual(agg, tariffs, report_date)
    else:
        park_income = total_net * (PARK_COMMISSION_PERCENT / 100)
    avg_check = round(total_net / total_orders, 2) if total_orders else 0.0

    # топы
    drivers = list(agg.values())
    top_best = sorted(drivers, key=lambda x: x.net, reverse=True)[:5]
    top_worst = sorted(drivers, key=lambda x: x.net)[:5]

    text = (
        f"📅 <b>Отчёт за {day_str}</b>\n"
        f"✅ Заказов: <b>{total_orders}</b>\n"
        f"💰 Выручка: <b>{total_net:,.0f} ₸</b>\n"
        f"🏦 Доход таксопарка: <b>{park_income:,.0f} ₸</b>\n"
        f"📊 Средний чек: <b>{avg_check:,.0f} ₸</b>\n"
        f"👤 Водителей: <b>{len(agg)}</b>\n"
    )

    # Разбивка по тарифным группам
    groups: Dict[str, Dict[str, float]] = {
        "Штатные": {"orders": 0, "net": 0.0, "count": 0},
        "АРА": {"orders": 0, "net": 0.0, "count": 0},
        "Тариф1": {"orders": 0, "net": 0.0, "count": 0},
        "Другие": {"orders": 0, "net": 0.0, "count": 0},
    }
    tariffs = tariffs or {}
    for fio, a in agg.items():
        info = tariffs.get(fio)
        if info is None:
            group = "Штатные"
        else:
            t = (info.get("tariff") or "").strip().lower()
            if t == "штатный":
                group = "Штатные"
            elif t == "ара":
                group = "АРА"
            elif t == "тариф1":
                group = "Тариф1"
            else:
                group = "Другие"
        groups[group]["orders"] += a.done
        groups[group]["net"] += a.net
        groups[group]["count"] += 1

    group_emoji = {"Штатные": "👔", "АРА": "🆕", "Тариф1": "2️⃣", "Другие": "📋"}
    text += "\n📂 <b>По тарифам:</b>\n"
    for gname, gdata in groups.items():
        if gdata["count"] == 0:
            continue
        emoji = group_emoji.get(gname, "•")
        text += f"{emoji} {gname} ({gdata['count']}): {gdata['net']:,.0f} ₸\n"
    text += "\n"

    if month_orders > 0:
        month_park = month_net * (PARK_COMMISSION_PERCENT / 100)
        text += (
            f"\n📆 <b>Итого с начала месяца:</b>\n"
            f"   Поездок: <b>{month_orders:,}</b>\n"
            f"   Выручка: <b>{month_net:,.0f} ₸</b>\n"
            f"   Доход парка: <b>{month_park:,.0f} ₸</b>\n"
        )

    text += "\n🏆 <b>ТОП‑5 водителей</b>\n"
    for d in top_best:
        text += f"• {d.fio} — {d.net:,.0f}\n"

    text += "\n📉 <b>Худшие 5</b>\n"
    for d in top_worst:
        text += f"• {d.fio} — {d.net:,.0f}\n"

    return text

# ------------------- BOT ACTIONS -------------------
CACHE: Dict[str, Dict[str, Any]] = {}  # key: "YYYY-MM-DD" -> {"rows":..., "agg":...}

async def _send_report_for_date(day_date, send_to_chat_id: int, bot):
    """Core report runner: fetch orders for day_date, write sheets, send summary + CSV to send_to_chat_id using given bot."""
    day_str = str(day_date)
    time_from, time_to, from_dt, to_dt = dubai_day_range(day_date)

    log.info("Run report for %s Dubai window=%s..%s", day_str, time_from, time_to)

    with requests.Session() as session:
        orders = orders_list_all(session, time_from, time_to, from_dt, to_dt)
        try:
            balances = fetch_driver_balances(session)
        except Exception as e:
            log.warning("fetch_driver_balances failed: %s", e)
            balances = {}

    try:
        directory = load_driver_directory()
    except Exception as e:
        log.warning("Driver directory unavailable (Google Sheets). Continue without it. Error: %s", e)
        directory = {}
    report_rows, agg = build_report_rows(day_str, orders, directory, balances)
    if not orders:
        await bot.send_message(
            chat_id=send_to_chat_id,
            text=f"⚠ За выбранный день ({day_str}) не найдено завершённых заказов по текущим критериям.",
        )
        return

    # если ФИО не пришло — предупреждаем сразу
    no_fio = any((a.fio == a.callsign or a.fio.strip() == "" or a.fio.lower() in ["неизвестно"]) for a in agg.values())
    warn = ""
    if no_fio:
        warn = "\n⚠️ <b>Внимание:</b> Яндекс не отдает ФИО по некоторым водителям (права/поля)."

    write_to_google_sheets(day_str, report_rows)
    csv_path = write_csv(day_str, report_rows)

    CACHE[day_str] = {"rows": report_rows, "agg": agg}

    # Считаем итог с начала месяца из Google Sheets
    import asyncio
    month_orders, month_net = await asyncio.to_thread(get_month_totals_from_sheets, day_date)
    tariffs = await asyncio.to_thread(load_tariffs)

    await bot.send_message(chat_id=send_to_chat_id, text=format_quick_summary(day_str, agg, month_orders, month_net, tariffs) + warn, parse_mode=ParseMode.HTML)
    with open(csv_path, "rb") as f:
        await bot.send_document(chat_id=send_to_chat_id, document=f, caption=f"CSV за {day_str}")


async def run_report_for_date(day_date, send_to_chat_id: int, context: ContextTypes.DEFAULT_TYPE):
    """Bot handler wrapper (kept for button/job usage)."""
    await _send_report_for_date(day_date, send_to_chat_id, context.bot)

async def on_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    chat_id = q.message.chat_id
    now_dubai = datetime.now(DUBAI_TZ).date()
    yesterday = now_dubai - timedelta(days=1)

    try:
        if q.data == "rep:yesterday":
            await q.edit_message_text("⏳ Загружаю вчерашний отчёт...")
            await run_report_for_date(yesterday, chat_id, context)

        elif q.data == "rep:today":
            await q.edit_message_text("⏳ Загружаю сегодняшний отчёт...")
            await run_report_for_date(now_dubai, chat_id, context)

        elif q.data == "rep:10days":
            await q.edit_message_text("⏳ Считаю отчёт за 10 дней...")

            end_date = now_dubai
            start_date = now_dubai - timedelta(days=9)

            total_orders = 0
            total_net = 0.0
            driver_net = {}
            driver_done = {}

            try:
                directory = load_driver_directory()
            except Exception:
                directory = {}

            with requests.Session() as session:
                for i in range(10):
                    d = start_date + timedelta(days=i)
                    time_from, time_to, from_dt, to_dt = dubai_day_range(d)
                    orders = orders_list_all(session, time_from, time_to, from_dt, to_dt)
                    rows, agg = build_report_rows(str(d), orders, directory)
                    for a in agg.values():
                        total_orders += a.done
                        total_net += a.net
                        driver_net[a.fio] = driver_net.get(a.fio, 0.0) + a.net
                        driver_done[a.fio] = driver_done.get(a.fio, 0) + a.done

            park_income = total_net * (PARK_COMMISSION_PERCENT / 100)
            avg_check = total_net / total_orders if total_orders else 0

            sorted_drivers = sorted(driver_net.items(), key=lambda x: x[1], reverse=True)
            top5 = sorted_drivers[:5]
            worst5 = sorted_drivers[-5:][::-1]
            top_lines = "\n".join(f"• {fio} — {int(net):,}".replace(",", " ") for fio, net in top5)
            worst_lines = "\n".join(f"• {fio} — {int(net):,}".replace(",", " ") for fio, net in worst5)

            date_range = f"{start_date} – {end_date}"
            text = (
                f"📅 <b>Отчёт за 10 дней ({date_range})</b>\n\n"
                f"✅ Заказов: <b>{total_orders}</b>\n"
                f"💰 Выручка: <b>{total_net:,.0f}</b>\n"
                f"🏦 Доход таксопарка: <b>{park_income:,.0f}</b>\n"
                f"📊 Средний чек: <b>{avg_check:,.0f}</b>\n"
                f"👤 Водителей: <b>{len(driver_net)}</b>\n\n"
                f"🏆 <b>ТОП-5 водителей</b>\n{top_lines}\n\n"
                f"📉 <b>Худшие 5</b>\n{worst_lines}"
            )

            await context.bot.send_message(
                chat_id=chat_id,
                text=text,
                parse_mode=ParseMode.HTML
            )

        elif q.data == "rep:compare":
            y = str(yesterday)
            t = str(now_dubai)

            # если нет в кэше — догружаем
            await q.edit_message_text("⏳ Сравниваю... (если данных нет — сначала загружу)")

            if y not in CACHE:
                await run_report_for_date(yesterday, chat_id, context)
            if t not in CACHE:
                await run_report_for_date(now_dubai, chat_id, context)

            agg_y: Dict[str, DriverAgg] = CACHE[y]["agg"]
            agg_t: Dict[str, DriverAgg] = CACHE[t]["agg"]

            orders_y = sum(a.done for a in agg_y.values())
            orders_t = sum(a.done for a in agg_t.values())
            net_y = sum(a.net for a in agg_y.values())
            net_t = sum(a.net for a in agg_t.values())

            text = (
                f"🆚 <b>Сравнение</b>\n\n"
                f"📅 Вчера ({y}): {orders_y} заказов, {net_y:.2f}\n"
                f"📅 Сегодня ({t}): {orders_t} заказов, {net_t:.2f}\n\n"
                f"Δ Заказы: <b>{orders_t - orders_y}</b>\n"
                f"Δ Сумма: <b>{(net_t - net_y):.2f}</b>\n"
            )
            await context.bot.send_message(chat_id=chat_id, text=text, parse_mode=ParseMode.HTML)
        elif q.data == "rep:month":
            # Показываем кнопки выбора месяца
            now = datetime.now(DUBAI_TZ)
            buttons = []
            row = []
            for i in range(5, -1, -1):  # последние 6 месяцев
                dt = (now.replace(day=1) - timedelta(days=1)) if i > 0 else now
                # вычисляем месяц назад
                month_dt = now.date().replace(day=1)
                for _ in range(i):
                    month_dt = (month_dt - timedelta(days=1)).replace(day=1)
                label = month_dt.strftime("%B %Y")
                cb = f"month:{month_dt.strftime('%Y-%m')}"
                row.append(InlineKeyboardButton(label, callback_data=cb))
                if len(row) == 2:
                    buttons.append(row)
                    row = []
            if row:
                buttons.append(row)
            kb = InlineKeyboardMarkup(buttons)
            await q.edit_message_text("Выбери месяц:", reply_markup=kb)

        elif q.data.startswith("month:"):
            year_month = q.data.split(":")[1]
            year, month = int(year_month.split("-")[0]), int(year_month.split("-")[1])
            import calendar
            month_start = now_dubai.replace(year=year, month=month, day=1)
            last_day = calendar.monthrange(year, month)[1]
            month_end = month_start.replace(day=last_day)
            # не уходим в будущее
            if month_end > now_dubai:
                month_end = now_dubai
            total_days = (month_end - month_start).days + 1
            month_name = month_start.strftime("%B %Y")
            await q.edit_message_text(f"⏳ Считаю отчёт за {month_name}...")
            total_orders = 0
            total_net = 0.0
            driver_net = {}
            try:
                directory = load_driver_directory()
            except Exception:
                directory = {}
            total_orders, total_net, driver_net = await fetch_range_data(month_start, month_end, directory)
            park_income = total_net * (PARK_COMMISSION_PERCENT / 100)
            avg_check = total_net / total_orders if total_orders else 0
            summary = (
                f"📆 <b>Отчёт за {month_name}</b>\n\n"
                f"✅ Заказов: <b>{total_orders}</b>\n"
                f"💰 Выручка: <b>{total_net:,.0f}</b>\n"
                f"🏦 Доход таксопарка: <b>{park_income:,.0f}</b>\n"
                f"📊 Средний чек: <b>{avg_check:,.0f}</b>\n"
                f"👤 Водителей: <b>{len(driver_net)}</b>"
            )
            await context.bot.send_message(chat_id=chat_id, text=summary, parse_mode=ParseMode.HTML)
            await send_all_drivers(context.bot, chat_id, driver_net, f"Все водители за {month_name}", None)

        elif q.data == "rep:top":
            await q.edit_message_text("⏳ Считаю ТОП водителей...")

            today = str(now_dubai)

            if today not in CACHE:
                await run_report_for_date(now_dubai, chat_id, context)

            agg = CACHE[today]["agg"]

            drivers = list(agg.values())

            top_best = sorted(drivers, key=lambda x: x.net, reverse=True)[:5]
            top_worst = sorted(drivers, key=lambda x: x.net)[:5]

            text = "🏆 <b>ТОП-5 водителей</b>\n\n"

            for d in top_best:
                hours = round(d.line_seconds / 3600, 2)
                text += f"• {d.fio} — {d.net:.0f} | {hours}ч\n"

            text += "\n📉 <b>Худшие 5</b>\n\n"

            for d in top_worst:
                hours = round(d.line_seconds / 3600, 2)
                text += f"• {d.fio} — {d.net:.0f} | {hours}ч\n"

            await context.bot.send_message(
                chat_id=chat_id,
                text=text,
                parse_mode=ParseMode.HTML
            )        
    except Exception as e:
        log.exception("button error")
        await context.bot.send_message(chat_id=chat_id, text=f"⚠ Ошибка: {e}")

# ------------------- DAILY AUTO SEND -------------------
async def job_daily(context: ContextTypes.DEFAULT_TYPE):
    """Каждый день в 09:00 Dubai шлём вчерашнее в личку CHAT_ID."""
    target = CHAT_ID_INT
    if target is None:
        log.warning("CHAT_ID not set -> skipping daily auto report")
        return
    day = datetime.now(DUBAI_TZ).date() - timedelta(days=1)
    try:
        await run_report_for_date(day, target, context)
    except Exception as e:
        await context.bot.send_message(chat_id=target, text=f"⚠ Ошибка авто-отчёта: {e}")

MAIN_KEYBOARD = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton("📊 Отчёт за вчера"), KeyboardButton("📈 Отчёт за сегодня")],
        [KeyboardButton("📅 Отчёт за 10 дней"), KeyboardButton("📆 Отчёт за месяц")],
        [KeyboardButton("🏆 ТОП водителей"), KeyboardButton("🆚 Сравнить вчера и сегодня")],
        [KeyboardButton("👔 Штатные водители")],
    ],
    resize_keyboard=True,

)

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Выбери действие:", reply_markup=MAIN_KEYBOARD)


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "📋 <b>Список команд</b>\n\n"
        "/start — Главное меню с кнопками\n\n"
        "<b>Тарифы водителей:</b>\n"
        "/settariff ФИО Процент — установить тариф (число, 'штатный'=5%, 'тариф1'=2%)\n"
        "<code>/settariff Феллер Вячеслав Анатольевич штатный</code>\n\n"
        "/setara ФИО [дата] — установить тариф АРА (1% первые 14 дней, потом 2%)\n"
        "<code>/setara Феллер Вячеслав Анатольевич 2026-06-11</code>\n\n"
        "/tariffs — показать список всех установленных тарифов\n\n"
        "/staff [вчера|сегодня] — отчёт только по штатным водителям\n"
        "<code>/staff вчера</code>\n\n"
        "/help — это сообщение"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=MAIN_KEYBOARD)


async def cmd_settariff(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    /settariff Фамилия Имя Отчество 5
    /settariff Фамилия Имя Отчество тариф1   -> 2%
    /settariff Фамилия Имя Отчество штатный  -> 5%
    """
    import asyncio
    args = context.args
    if len(args) < 2:
        await update.message.reply_text(
            "Использование:\n"
            "<code>/settariff ФИО Процент</code>\n\n"
            "Примеры:\n"
            "<code>/settariff Феллер Вячеслав Анатольевич 5</code>\n"
            "<code>/settariff Феллер Вячеслав Анатольевич штатный</code>\n"
            "<code>/settариff Феллер Вячеслав Анатольевич тариф1</code>",
            parse_mode=ParseMode.HTML,
        )
        return

    last_arg = args[-1].lower()
    fio = " ".join(args[:-1])

    if last_arg in ("штатный", "штат", "5"):
        tariff_name, percent = "Штатный", 5.0
    elif last_arg in ("тариф1", "тариф 1", "2"):
        tariff_name, percent = "Тариф1", 2.0
    else:
        try:
            percent = float(last_arg.replace(",", "."))
            tariff_name = "Кастом"
        except ValueError:
            await update.message.reply_text("⚠ Не понял процент. Укажи число, 'штатный' или 'тариф1'.")
            return

    try:
        await asyncio.to_thread(set_tariff, fio, tariff_name, percent, None)
        await update.message.reply_text(
            f"✅ Тариф для <b>{fio}</b> установлен: <b>{tariff_name} ({percent}%)</b>",
            parse_mode=ParseMode.HTML,
        )
    except Exception as e:
        log.exception("settariff error")
        await update.message.reply_text(f"⚠ Ошибка: {e}")


async def cmd_setara(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    /setara Фамилия Имя Отчество [YYYY-MM-DD]
    Если дата не указана — используется сегодняшняя (Dubai).
    Первые 14 дней с этой даты — 1%, далее автоматически 2%.
    """
    import asyncio
    import datetime as dt_mod
    args = context.args
    if len(args) < 1:
        await update.message.reply_text(
            "Использование:\n"
            "<code>/setara ФИО [YYYY-MM-DD]</code>\n\n"
            "Пример:\n"
            "<code>/setara Феллер Вячеслав Анатольевич 2026-06-11</code>",
            parse_mode=ParseMode.HTML,
        )
        return

    last_arg = args[-1]
    ara_date = None
    try:
        ara_date = dt_mod.date.fromisoformat(last_arg)
        fio = " ".join(args[:-1])
    except ValueError:
        fio = " ".join(args)
        ara_date = datetime.now(DUBAI_TZ).date()

    try:
        await asyncio.to_thread(set_tariff, fio, "АРА", None, ara_date)
        await update.message.reply_text(
            f"✅ Тариф АРА для <b>{fio}</b> установлен с {ara_date}.\n"
            f"Первые 14 дней — 1%, далее автоматически 2%.",
            parse_mode=ParseMode.HTML,
        )
    except Exception as e:
        log.exception("setara error")
        await update.message.reply_text(f"⚠ Ошибка: {e}")


async def cmd_importtariffs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Одноразовая массовая загрузка тарифов из встроенного списка BULK_TARIFF_IMPORT."""
    import asyncio
    import datetime as dt_mod

    await update.message.reply_text(f"⏳ Импортирую {len(BULK_TARIFF_IMPORT)} тарифов...")

    ok_count = 0
    fail_count = 0
    for fio, tariff, percent, date_str in BULK_TARIFF_IMPORT:
        try:
            ara_date = None
            if date_str:
                ara_date = dt_mod.date.fromisoformat(date_str)
            await asyncio.to_thread(set_tariff, fio, tariff, percent, ara_date)
            ok_count += 1
        except Exception as e:
            log.warning("importtariffs failed for %s: %s", fio, e)
            fail_count += 1
        await asyncio.sleep(0.3)

    await update.message.reply_text(
        f"✅ Импорт завершён.\nУспешно: {ok_count}\nОшибок: {fail_count}",
        reply_markup=MAIN_KEYBOARD,
    )



async def cmd_tariffs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показывает список всех установленных тарифов."""
    if update.message is None:
        return
    import asyncio
    import datetime as dt_mod
    try:
        tariffs = await asyncio.to_thread(load_tariffs)
    except Exception as e:
        await update.message.reply_text(f"⚠ Ошибка: {e}")
        return

    if not tariffs:
        await update.message.reply_text("Список тарифов пуст. Используй /settariff или /setara чтобы добавить.")
        return

    today = datetime.now(DUBAI_TZ).date()
    lines = ["📋 <b>Тарифы водителей</b>\n"]
    for fio, info in tariffs.items():
        tariff = info.get("tariff") or "—"
        if tariff.upper() == "АРА":
            ara_start = info.get("ara_start")
            if ara_start:
                pct = get_driver_percent(fio, tariffs, today)
                days = (today - ara_start).days
                lines.append(f"• {fio} — АРА (с {ara_start}, день {days+1}) → <b>{pct}%</b>")
            else:
                lines.append(f"• {fio} — АРА → <b>1%</b>")
        else:
            pct = info.get("percent")
            pct_str = f"{pct}%" if pct is not None else f"{PARK_COMMISSION_PERCENT}% (дефолт)"
            lines.append(f"• {fio} — {tariff} → <b>{pct_str}</b>")

    await update.message.reply_text("\n".join(lines), parse_mode=ParseMode.HTML)


async def _send_staff_report(day_date, chat_id: int, bot):
    """Core: отчёт только по штатным водителям (тариф 'Штатный' или без тарифа) за day_date."""
    import asyncio
    day_str = str(day_date)
    await bot.send_message(chat_id=chat_id, text=f"⏳ Считаю отчёт по штатным водителям за {day_str}...")

    time_from, time_to, from_dt, to_dt = dubai_day_range(day_date)
    try:
        with requests.Session() as session:
            orders = await asyncio.to_thread(orders_list_all, session, time_from, time_to, from_dt, to_dt)
            balances = await asyncio.to_thread(fetch_driver_balances, session)
    except Exception as e:
        await bot.send_message(chat_id=chat_id, text=f"⚠ Ошибка загрузки заказов: {e}")
        return

    try:
        directory = await asyncio.to_thread(load_driver_directory)
    except Exception:
        directory = {}

    _, agg = build_report_rows(day_str, orders, directory, balances)
    tariffs = await asyncio.to_thread(load_tariffs)

    # Штатные = явно "Штатный" в тарифах, ИЛИ вообще без записи в тарифах (дефолт)
    staff_agg = {}
    for fio, a in agg.items():
        info = tariffs.get(fio)
        is_staff = (info is None) or ((info.get("tariff") or "").strip().lower() == "штатный")
        if is_staff:
            staff_agg[fio] = a

    if not staff_agg:
        await bot.send_message(chat_id=chat_id, text="Штатных водителей за этот день не найдено.", reply_markup=MAIN_KEYBOARD)
        return

    total_orders = sum(a.done for a in staff_agg.values())
    total_net = sum(a.net for a in staff_agg.values())
    park_income = calc_park_income_individual(staff_agg, tariffs, day_date)
    avg_check = round(total_net / total_orders, 2) if total_orders else 0.0

    text = (
        f"👔 <b>Штатные водители за {day_str}</b>\n\n"
        f"✅ Заказов: <b>{total_orders}</b>\n"
        f"💰 Выручка: <b>{total_net:,.0f} ₸</b>\n"
        f"🏦 Доход таксопарка: <b>{park_income:,.0f} ₸</b>\n"
        f"📊 Средний чек: <b>{avg_check:,.0f} ₸</b>\n"
        f"👤 Водителей: <b>{len(staff_agg)}</b>"
    )
    await bot.send_message(chat_id=chat_id, text=text, parse_mode=ParseMode.HTML)

    try:
        await send_staff_drivers_with_balance(bot, chat_id, staff_agg, balances, f"Штатные водители за {day_str}", MAIN_KEYBOARD)
    except Exception as e:
        log.exception("_send_staff_report send list error")
        await bot.send_message(chat_id=chat_id, text=f"⚠ Ошибка при отправке списка водителей: {e}")


async def cmd_staff(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    /staff [вчера|сегодня] — отчёт только по штатным водителям (тариф "Штатный" или без тарифа).
    """
    chat_id = update.effective_chat.id
    args = context.args
    period = (args[0].lower() if args else "вчера")

    now_dubai = datetime.now(DUBAI_TZ).date()
    if "сегодня" in period:
        day_date = now_dubai
    else:
        day_date = now_dubai - timedelta(days=1)

    await _send_staff_report(day_date, chat_id, context.bot)



async def send_all_drivers(bot, chat_id, driver_net: dict, title: str, keyboard):
    """Отправляет всех водителей чанками по 20, отсортированных по выручке."""
    sorted_drivers = sorted(driver_net.items(), key=lambda x: x[1], reverse=True)
    chunk_size = 20
    for i in range(0, len(sorted_drivers), chunk_size):
        chunk = sorted_drivers[i:i+chunk_size]
        num_start = i + 1
        lines = "\n".join(
            f"{num_start + j}. {fio} — {int(net):,}".replace(",", " ")
            for j, (fio, net) in enumerate(chunk)
        )
        part_num = i // chunk_size + 1
        total_parts = (len(sorted_drivers) + chunk_size - 1) // chunk_size
        header = f"👥 <b>{title}</b> (часть {part_num}/{total_parts})\n\n" if total_parts > 1 else f"👥 <b>{title}</b>\n\n"
        await bot.send_message(
            chat_id=chat_id,
            text=header + lines,
            parse_mode="HTML",
            reply_markup=keyboard if i + chunk_size >= len(sorted_drivers) else None,
        )

async def send_staff_drivers_with_balance(bot, chat_id, staff_agg: Dict[str, "DriverAgg"], balances: Dict[str, float], title: str, keyboard):
    """Отправляет штатных водителей чанками по 20: ФИО, выручка и текущий баланс из Яндекса."""
    sorted_drivers = sorted(staff_agg.items(), key=lambda x: x[1].net, reverse=True)
    chunk_size = 20
    for i in range(0, len(sorted_drivers), chunk_size):
        chunk = sorted_drivers[i:i+chunk_size]
        num_start = i + 1
        lines_list = []
        for j, (fio, a) in enumerate(chunk):
            bal = balances.get(fio)
            if bal is None:
                bal = a.balance
            net_str = f"{int(a.net):,}".replace(",", " ")
            bal_str = f"{bal:,.0f}".replace(",", " ") if bal else "—"
            lines_list.append(f"{num_start + j}. {fio} — {net_str} ₸ | Баланс: {bal_str} ₸")
        lines = "\n".join(lines_list)
        part_num = i // chunk_size + 1
        total_parts = (len(sorted_drivers) + chunk_size - 1) // chunk_size
        header = f"👥 <b>{title}</b> (часть {part_num}/{total_parts})\n\n" if total_parts > 1 else f"👥 <b>{title}</b>\n\n"
        await bot.send_message(
            chat_id=chat_id,
            text=header + lines,
            parse_mode="HTML",
            reply_markup=keyboard if i + chunk_size >= len(sorted_drivers) else None,
        )


async def fetch_day_data(d, directory):
    """Загружает данные за один день в отдельном потоке (для параллельности)."""
    import asyncio
    def _sync():
        with requests.Session() as session:
            time_from, time_to, from_dt, to_dt = dubai_day_range(d)
            orders = orders_list_all(session, time_from, time_to, from_dt, to_dt)
            rows, agg = build_report_rows(str(d), orders, directory)
            return agg
    return await asyncio.to_thread(_sync)


async def fetch_range_data(start_date, end_date, directory):
    """Параллельно загружает данные за все дни диапазона."""
    import asyncio
    days = []
    d = start_date
    while d <= end_date:
        days.append(d)
        d += timedelta(days=1)
    results = await asyncio.gather(*[fetch_day_data(d, directory) for d in days], return_exceptions=True)
    total_orders = 0
    total_net = 0.0
    driver_net = {}
    for agg in results:
        if isinstance(agg, Exception):
            log.warning(f"Day fetch error: {agg}")
            continue
        for a in agg.values():
            total_orders += a.done
            total_net += a.net
            driver_net[a.fio] = driver_net.get(a.fio, 0.0) + a.net
    return total_orders, total_net, driver_net


async def on_text_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обрабатываем нажатия Reply Keyboard кнопок."""
    text = update.message.text.strip()
    chat_id = update.effective_chat.id
    bot = context.bot
    now_dubai = datetime.now(DUBAI_TZ).date()
    yesterday = now_dubai - timedelta(days=1)

    try:
        if "вчера" in text.lower() and "сравн" not in text.lower():
            await update.message.reply_text("⏳ Загружаю вчерашний отчёт...")
            await _send_report_for_date(yesterday, chat_id, bot)

        elif "сегодня" in text.lower() and "сравн" not in text.lower():
            await update.message.reply_text("⏳ Загружаю сегодняшний отчёт...")
            await _send_report_for_date(now_dubai, chat_id, bot)

        elif "10 дней" in text.lower():
            await update.message.reply_text("⏳ Считаю отчёт за 10 дней...")
            total_orders = 0
            total_net = 0.0
            driver_net = {}
            start_date = now_dubai - timedelta(days=9)
            end_date = now_dubai
            try:
                directory = load_driver_directory()
            except Exception:
                directory = {}
            total_orders, total_net, driver_net = await fetch_range_data(start_date, end_date, directory)
            park_income = total_net * (PARK_COMMISSION_PERCENT / 100)
            avg_check = total_net / total_orders if total_orders else 0
            date_range = f"{start_date} – {end_date}"
            summary = (
                f"📅 <b>Отчёт за 10 дней ({date_range})</b>\n\n"
                f"✅ Заказов: <b>{total_orders}</b>\n"
                f"💰 Выручка: <b>{total_net:,.0f}</b>\n"
                f"🏦 Доход таксопарка: <b>{park_income:,.0f}</b>\n"
                f"📊 Средний чек: <b>{avg_check:,.0f}</b>\n"
                f"👤 Водителей: <b>{len(driver_net)}</b>"
            )
            await update.message.reply_text(summary, parse_mode=ParseMode.HTML)
            await send_all_drivers(bot, chat_id, driver_net, "Все водители за 10 дней", MAIN_KEYBOARD)

        elif "месяц" in text.lower():
            # Показываем inline кнопки выбора месяца
            buttons = []
            row = []
            for i in range(5, -1, -1):
                month_dt = now_dubai.replace(day=1)
                for _ in range(i):
                    month_dt = (month_dt - timedelta(days=1)).replace(day=1)
                label = month_dt.strftime("%B %Y")
                cb = f"month:{month_dt.strftime('%Y-%m')}"
                row.append(InlineKeyboardButton(label, callback_data=cb))
                if len(row) == 2:
                    buttons.append(row)
                    row = []
            if row:
                buttons.append(row)
            kb = InlineKeyboardMarkup(buttons)
            await update.message.reply_text("Выбери месяц:", reply_markup=kb)

        elif "штатные" in text.lower():
            await _send_staff_report(yesterday, chat_id, bot)

        elif "топ" in text.lower():
            await update.message.reply_text("⏳ Считаю ТОП водителей...")
            today = str(now_dubai)
            if today not in CACHE:
                await _send_report_for_date(now_dubai, chat_id, bot)
            agg = CACHE.get(today, {}).get("agg", {})
            if not agg:
                await update.message.reply_text("Нет данных за сегодня.", reply_markup=MAIN_KEYBOARD)
                return
            drivers = list(agg.values())
            top_best = sorted(drivers, key=lambda x: x.net, reverse=True)[:5]
            top_worst = sorted(drivers, key=lambda x: x.net)[:5]
            text_out = "🏆 <b>ТОП-5 водителей</b>\n\n"
            for d in top_best:
                text_out += f"• {d.fio} — {d.net:.0f} | {round(d.line_seconds/3600,2)}ч\n"
            text_out += "\n📉 <b>Худшие 5</b>\n\n"
            for d in top_worst:
                text_out += f"• {d.fio} — {d.net:.0f} | {round(d.line_seconds/3600,2)}ч\n"
            await update.message.reply_text(text_out, parse_mode=ParseMode.HTML, reply_markup=MAIN_KEYBOARD)

        elif "сравн" in text.lower():
            await update.message.reply_text("⏳ Сравниваю...")
            y = str(yesterday)
            t = str(now_dubai)
            if y not in CACHE:
                await _send_report_for_date(yesterday, chat_id, bot)
            if t not in CACHE:
                await _send_report_for_date(now_dubai, chat_id, bot)
            agg_y = CACHE.get(y, {}).get("agg", {})
            agg_t = CACHE.get(t, {}).get("agg", {})
            orders_y = sum(a.done for a in agg_y.values())
            orders_t = sum(a.done for a in agg_t.values())
            net_y = sum(a.net for a in agg_y.values())
            net_t = sum(a.net for a in agg_t.values())
            await update.message.reply_text(
                f"🆚 <b>Сравнение</b>\n\n"
                f"📅 Вчера ({y}): {orders_y} заказов, {net_y:.2f}\n"
                f"📅 Сегодня ({t}): {orders_t} заказов, {net_t:.2f}\n\n"
                f"Δ Заказы: <b>{orders_t - orders_y}</b>\n"
                f"Δ Сумма: <b>{(net_t - net_y):.2f}</b>",
                parse_mode=ParseMode.HTML,
                reply_markup=MAIN_KEYBOARD,
            )

        else:
            await update.message.reply_text("Выбери действие:", reply_markup=MAIN_KEYBOARD)

    except Exception as e:
        log.exception("text button error")
        await update.message.reply_text(f"⚠ Ошибка: {e}", reply_markup=MAIN_KEYBOARD)

async def setup_bot_commands(app: Application):
    """Регистрирует команды бота для подсказок при вводе '/'."""
    from telegram import BotCommand
    commands = [
        BotCommand("start", "Главное меню с кнопками"),
        BotCommand("help", "Список всех команд"),
        BotCommand("settariff", "Установить тариф водителю (штатный/тариф1/%)"),
        BotCommand("setara", "Установить тариф АРА с датой начала"),
        BotCommand("tariffs", "Показать список всех тарифов"),
        BotCommand("staff", "Отчёт только по штатным водителям"),
    ]
    await app.bot.set_my_commands(commands)


def build_app() -> Application:
    app = Application.builder().token(BOT_TOKEN).post_init(setup_bot_commands).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("settariff", cmd_settariff))
    app.add_handler(CommandHandler("setara", cmd_setara))
    app.add_handler(CommandHandler("tariffs", cmd_tariffs))
    app.add_handler(CommandHandler("staff", cmd_staff))
    app.add_handler(CommandHandler("importtariffs", cmd_importtariffs))
    app.add_handler(CallbackQueryHandler(on_button))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text_button))

    # job queue: 09:00 Dubai = 05:00 UTC (Dubai = UTC+4)
    from datetime import timezone, time as dtime
    report_time_utc = dtime(hour=5, minute=0, tzinfo=timezone.utc)
    app.job_queue.run_daily(job_daily, time=report_time_utc, days=(0,1,2,3,4,5,6), name="daily_report")
    return app

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-now", action="store_true", help="Run yesterday report (Dubai) to CHAT_ID and exit")
    parser.add_argument("--run-date", type=str, default=None, help="Run report for YYYY-MM-DD to CHAT_ID and exit")
    parser.add_argument("--import-tariffs", type=str, default=None,
                         help="Bulk import tariffs from a commands file (one /settariff or /setara per line) and exit")
    args = parser.parse_args()

    if args.import_tariffs:
        count = 0
        with open(args.import_tariffs, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = line.split()
                if parts[0] == "/settariff":
                    last = parts[-1].lower()
                    fio = " ".join(parts[1:-1])
                    if last in ("штатный", "штат", "5"):
                        tariff_name, percent = "Штатный", 5.0
                    elif last in ("тариф1", "тариф 1", "2"):
                        tariff_name, percent = "Тариф1", 2.0
                    else:
                        try:
                            percent = float(last.replace(",", "."))
                            tariff_name = "Кастом"
                        except ValueError:
                            print(f"SKIP (bad percent): {line}")
                            continue
                    set_tariff(fio, tariff_name, percent, None)
                    print(f"OK settariff: {fio} -> {tariff_name} ({percent}%)")
                    count += 1
                elif parts[0] == "/setara":
                    import datetime as dt_mod
                    last = parts[-1]
                    try:
                        ara_date = dt_mod.date.fromisoformat(last)
                        fio = " ".join(parts[1:-1])
                    except ValueError:
                        fio = " ".join(parts[1:])
                        ara_date = datetime.now(DUBAI_TZ).date()
                    set_tariff(fio, "АРА", None, ara_date)
                    print(f"OK setara: {fio} -> {ara_date}")
                    count += 1
                else:
                    print(f"SKIP (unknown command): {line}")
                time.sleep(0.3)  # не долбить Google Sheets API слишком быстро
        print(f"\nDone. Imported {count} tariff entries.")
        raise SystemExit(0)

    # CLI: отправить отчёт в CHAT_ID и выйти
    if args.run_now or args.run_date:
        if CHAT_ID_INT is None:
            raise SystemExit("CHAT_ID is not set. Put CHAT_ID in .env to use --run-now/--run-date.")

        day = datetime.now(DUBAI_TZ).date() - timedelta(days=1)
        if args.run_date:
            day = datetime.fromisoformat(args.run_date).date()

        import asyncio

        async def _cli_run():
            app = build_app()
            await app.initialize()
            try:
                await _send_report_for_date(day, CHAT_ID_INT, app.bot)
            finally:
                await app.shutdown()

        asyncio.run(_cli_run())
        raise SystemExit(0)

    # PROD: run bot polling
    app = build_app()
    log.info("Bot started. Кнопки без /start. Авто-отчёт в 09:00 Dubai.")
    app.run_polling(close_loop=False)