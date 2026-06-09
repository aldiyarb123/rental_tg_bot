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
    creds = Credentials.from_service_account_file(GOOGLE_SA_JSON, scopes=GS_SCOPES)
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

def build_report_rows(day_str: str, orders: List[dict], directory: Dict[str, dict]) -> Tuple[List[List[Any]], Dict[str, DriverAgg]]:
    agg: Dict[str, DriverAgg] = {}

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

def format_quick_summary(day_str: str, agg: Dict[str, DriverAgg]) -> str:
    total_orders = sum(a.done for a in agg.values())
    total_net = sum(a.net for a in agg.values())
    park_income = total_net * (PARK_COMMISSION_PERCENT / 100)
    avg_check = round(total_net / total_orders, 2) if total_orders else 0.0

    # топы
    drivers = list(agg.values())
    top_best = sorted(drivers, key=lambda x: x.net, reverse=True)[:5]
    top_worst = sorted(drivers, key=lambda x: x.net)[:5]

    text = (
        f"📅 <b>Отчёт за {day_str}</b>\n"
        f"✅ Заказов: <b>{total_orders}</b>\n"
        f"💰 Выручка: <b>{total_net:.2f}</b>\n"
        f"🏦 Доход таксопарка: <b>{park_income:.2f}</b>\n"
        f"📊 Средний чек: <b>{avg_check:.2f}</b>\n"
        f"👤 Водителей: <b>{len(agg)}</b>\n\n"
    )

    text += "🏆 <b>ТОП‑5 водителей</b>\n"
    for d in top_best:
        text += f"• {d.fio} — {d.net:.0f}\n"

    text += "\n📉 <b>Худшие 5</b>\n"
    for d in top_worst:
        text += f"• {d.fio} — {d.net:.0f}\n"

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
        directory = load_driver_directory()
    except Exception as e:
        log.warning("Driver directory unavailable (Google Sheets). Continue without it. Error: %s", e)
        directory = {}
    report_rows, agg = build_report_rows(day_str, orders, directory)
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

    await bot.send_message(chat_id=send_to_chat_id, text=format_quick_summary(day_str, agg) + warn, parse_mode=ParseMode.HTML)
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
            drivers = {}

            with requests.Session() as session:
                for i in range(10):
                    d = start_date + timedelta(days=i)
                    time_from, time_to, from_dt, to_dt = dubai_day_range(d)
                    orders = orders_list_all(session, time_from, time_to, from_dt, to_dt)

                    try:
                        directory = load_driver_directory()
                    except Exception:
                        directory = {}

                    rows, agg = build_report_rows(str(d), orders, directory)

                    for a in agg.values():
                        total_orders += a.done
                        total_net += a.net

                        if a.fio not in drivers:
                            drivers[a.fio] = 0
                        drivers[a.fio] += a.net

            park_income = total_net * (PARK_COMMISSION_PERCENT / 100)

            text = (
                f"📅 <b>Отчёт за 10 дней</b>\n\n"
                f"✅ Заказов: <b>{total_orders}</b>\n"
                f"💰 Выручка: <b>{total_net:.2f}</b>\n"
                f"🏦 Доход таксопарка: <b>{park_income:.2f}</b>\n"
                f"👤 Водителей: <b>{len(drivers)}</b>\n"
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
        [KeyboardButton("📅 Отчёт за 10 дней"), KeyboardButton("🏆 ТОП водителей")],
        [KeyboardButton("🆚 Сравнить вчера и сегодня")],
    ],
    resize_keyboard=True,

)

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Выбери действие:", reply_markup=MAIN_KEYBOARD)

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
            drivers = {}
            start_date = now_dubai - timedelta(days=9)
            with requests.Session() as session:
                for i in range(10):
                    d = start_date + timedelta(days=i)
                    time_from, time_to, from_dt, to_dt = dubai_day_range(d)
                    orders = orders_list_all(session, time_from, time_to, from_dt, to_dt)
                    try:
                        directory = load_driver_directory()
                    except Exception:
                        directory = {}
                    rows, agg = build_report_rows(str(d), orders, directory)
                    for a in agg.values():
                        total_orders += a.done
                        total_net += a.net
                        drivers[a.fio] = drivers.get(a.fio, 0) + a.net
            park_income = total_net * (PARK_COMMISSION_PERCENT / 100)
            await update.message.reply_text(
                f"📅 <b>Отчёт за 10 дней</b>\n\n"
                f"✅ Заказов: <b>{total_orders}</b>\n"
                f"💰 Выручка: <b>{total_net:.2f}</b>\n"
                f"🏦 Доход таксопарка: <b>{park_income:.2f}</b>\n"
                f"👤 Водителей: <b>{len(drivers)}</b>",
                parse_mode=ParseMode.HTML,
                reply_markup=MAIN_KEYBOARD,
            )

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

def build_app() -> Application:
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
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
    args = parser.parse_args()

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