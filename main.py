"""Temir Daftar API — personal debt tracker for a Telegram Mini App.

Runs on Vercel (FastAPI) with Neon Postgres. Every /api request must carry the
Telegram WebApp initData in the X-Init-Data header; it is verified with the bot
token, so a user can only ever read and change their own data.
"""
import hashlib, hmac, io, json, logging, os, time
from datetime import date, datetime, timedelta, timezone
from typing import List, Optional
from urllib.parse import parse_qsl

import httpx
import psycopg
from psycopg.rows import dict_row
from fastapi import Depends, FastAPI, Header, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field, field_validator
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

logging.getLogger("httpx").setLevel(logging.WARNING)  # its request log would print the bot token

app = FastAPI(title="Temir Daftar API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://vimovich.github.io"],
    allow_origin_regex=r"https://.*\.vercel\.app|http://localhost(:\d+)?",
    allow_methods=["*"],
    allow_headers=["*"],
)

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
CRON_SECRET = os.environ.get("CRON_SECRET", "")
APP_URL = os.environ.get("APP_URL", "https://vimovich.github.io/temirkassa/")
API_URL = os.environ.get("API_URL", "https://temir-kassa-api.vercel.app")
TZ = timezone(timedelta(hours=5))  # Tashkent
DATABASE_URL = (os.environ.get("DATABASE_URL") or os.environ.get("POSTGRES_URL") or "")
WEBHOOK_SECRET = hashlib.sha256(("wh:" + BOT_TOKEN).encode()).hexdigest()[:32]
SETUP_KEY = hashlib.sha256(("setup:" + BOT_TOKEN).encode()).hexdigest()[:16]

HUES = ["blue", "green", "amber", "rose", "teal", "violet"]
STRATEGIES = {"small", "urgent", "even"}
MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август",
          "сентябрь", "октябрь", "ноябрь", "декабрь"]
MONTHS_IN = ["январе", "феврале", "марте", "апреле", "мае", "июне", "июле", "августе",
             "сентябре", "октябре", "ноябре", "декабре"]
MONTHS_G = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа",
            "сентября", "октября", "ноября", "декабря"]

# ─── Database ─────────────────────────────────────────────
_initialized = False

def db():
    global _initialized
    if not DATABASE_URL:
        raise HTTPException(500, "База данных не подключена")
    conn = psycopg.connect(DATABASE_URL, row_factory=dict_row, prepare_threshold=None)
    if not _initialized:
        init_db(conn)
        _initialized = True
    return conn

def init_db(conn):
    conn.execute("""
    CREATE TABLE IF NOT EXISTS tk_users (
        telegram_id   BIGINT PRIMARY KEY,
        first_name    TEXT,
        monthly_plan  DOUBLE PRECISION DEFAULT 0,
        strategy      TEXT DEFAULT 'small',
        remind_on     BOOLEAN DEFAULT TRUE,
        report_on     BOOLEAN DEFAULT TRUE,
        due_alert_on  BOOLEAN DEFAULT TRUE,
        theme         TEXT DEFAULT 'auto',
        created_at    TIMESTAMPTZ DEFAULT now()
    )""")
    conn.execute("""
    CREATE TABLE IF NOT EXISTS tk_debts (
        id           SERIAL PRIMARY KEY,
        telegram_id  BIGINT NOT NULL,
        who          TEXT NOT NULL,
        kind         TEXT DEFAULT '',
        cur          TEXT NOT NULL DEFAULT 'UZS',
        amount       DOUBLE PRECISION NOT NULL,
        rate         DOUBLE PRECISION,
        taken        DATE NOT NULL,
        due          DATE,
        fixed        DOUBLE PRECISION,
        pay_day      INTEGER,
        hue          TEXT DEFAULT 'blue',
        note         TEXT DEFAULT '',
        created_at   TIMESTAMPTZ DEFAULT now()
    )""")
    conn.execute("""
    CREATE TABLE IF NOT EXISTS tk_payments (
        id           SERIAL PRIMARY KEY,
        telegram_id  BIGINT NOT NULL,
        debt_id      INTEGER NOT NULL REFERENCES tk_debts(id) ON DELETE CASCADE,
        date         DATE NOT NULL,
        amount       DOUBLE PRECISION NOT NULL,
        rate         DOUBLE PRECISION,
        created_at   TIMESTAMPTZ DEFAULT now()
    )""")
    conn.execute("CREATE INDEX IF NOT EXISTS tk_debts_user ON tk_debts(telegram_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS tk_pay_user ON tk_payments(telegram_id)")
    conn.commit()

# ─── Telegram auth ────────────────────────────────────────
def verify_init_data(init_data: str) -> dict:
    if not BOT_TOKEN:
        raise HTTPException(500, "BOT_TOKEN не задан")
    if not init_data:
        raise HTTPException(401, "Откройте приложение через Telegram")
    pairs = dict(parse_qsl(init_data, keep_blank_values=True))
    received = pairs.pop("hash", "")
    check = "\n".join(f"{k}={v}" for k, v in sorted(pairs.items()))
    secret = hmac.new(b"WebAppData", BOT_TOKEN.encode(), hashlib.sha256).digest()
    calc = hmac.new(secret, check.encode(), hashlib.sha256).hexdigest()
    if not received or not hmac.compare_digest(calc, received):
        raise HTTPException(401, "Подпись Telegram не прошла проверку")
    if time.time() - int(pairs.get("auth_date", "0")) > 7 * 86400:
        raise HTTPException(401, "Сессия устарела, откройте приложение заново")
    try:
        return json.loads(pairs["user"])
    except Exception:
        raise HTTPException(401, "Нет данных пользователя")

def current_user(x_init_data: str = Header(default="")) -> dict:
    return verify_init_data(x_init_data)

# ─── Models ───────────────────────────────────────────────
class DebtIn(BaseModel):
    who: str = Field(min_length=1, max_length=60)
    kind: str = Field(default="", max_length=60)
    cur: str = "UZS"
    amount: float = Field(gt=0, lt=1e13)
    rate: Optional[float] = Field(default=None, gt=0, lt=1e6)
    taken: date
    due: Optional[date] = None
    fixed: Optional[float] = Field(default=None, gt=0, lt=1e13)
    pay_day: Optional[int] = Field(default=None, ge=1, le=31)
    hue: str = "blue"
    note: str = Field(default="", max_length=300)

    @field_validator("cur")
    @classmethod
    def _cur(cls, v):
        if v not in ("UZS", "USD"):
            raise ValueError("Валюта: UZS или USD")
        return v

    @field_validator("hue")
    @classmethod
    def _hue(cls, v):
        return v if v in HUES else "blue"

class PaymentIn(BaseModel):
    debt_id: int
    date: date
    amount: float = Field(gt=0, lt=1e13)       # in the debt's own currency
    rate: Optional[float] = Field(default=None, gt=0, lt=1e6)  # soum per $ that day, USD debts

class SettingsIn(BaseModel):
    monthly_plan: Optional[float] = Field(default=None, ge=0, lt=1e13)
    strategy: Optional[str] = None
    remind_on: Optional[bool] = None
    report_on: Optional[bool] = None
    due_alert_on: Optional[bool] = None
    theme: Optional[str] = None

# ─── Helpers ──────────────────────────────────────────────
def today() -> date:
    return datetime.now(TZ).date()

def ym(d: date) -> int:
    return d.year * 12 + d.month - 1

def fmt(n: float) -> str:
    return f"{round(n):,}".replace(",", " ")

def usd(n: float) -> str:
    n = round(n, 2)
    return "$" + (f"{n:,.2f}" if n % 1 else f"{n:,.0f}")

def plural(n, a, b, c):
    m, h = n % 10, n % 100
    return a if m == 1 and h != 11 else b if 2 <= m <= 4 and not 10 <= h < 20 else c

def ensure_user(conn, tg: dict):
    conn.execute("""INSERT INTO tk_users(telegram_id, first_name) VALUES(%s,%s)
                    ON CONFLICT(telegram_id) DO UPDATE SET first_name=excluded.first_name""",
                 (tg["id"], (tg.get("first_name") or "")[:60]))
    conn.commit()

def load_state(conn, uid: int):
    user = conn.execute("SELECT * FROM tk_users WHERE telegram_id=%s", (uid,)).fetchone()
    debts = conn.execute("SELECT * FROM tk_debts WHERE telegram_id=%s ORDER BY id", (uid,)).fetchall()
    pays = conn.execute("SELECT * FROM tk_payments WHERE telegram_id=%s ORDER BY date, id", (uid,)).fetchall()
    by = {d["id"]: [] for d in debts}
    for p in pays:
        by.setdefault(p["debt_id"], []).append(p)
    for d in debts:
        d["payments"] = by.get(d["id"], [])
        r = d["rate"] if d["cur"] == "USD" else 1
        d["paid_cur"] = sum(p["amount"] for p in d["payments"])
        d["left_cur"] = max(0.0, d["amount"] - d["paid_cur"])
        d["left_uzs"] = d["left_cur"] * r
        d["closed"] = d["left_cur"] <= 0.005
    return user, debts

def jsonable(o):
    if isinstance(o, datetime):
        return o.isoformat()
    if isinstance(o, date):
        return o.isoformat()
    raise TypeError(type(o))

def split_month(debts, monthly: float, strategy: str, month_ym: int):
    """Mirror of the app's split: this month's fixed instalments first, then by strategy."""
    pool = monthly
    alloc = {d["id"]: 0.0 for d in debts}
    open_ = [d for d in debts if not d["closed"]]
    for d in open_:
        if d["fixed"]:
            paid_this = sum(p["amount"] for p in d["payments"] if ym(p["date"]) == month_ym)
            need = max(0.0, d["fixed"] - paid_this) * (d["rate"] if d["cur"] == "USD" else 1)
            a = min(need, d["left_uzs"], pool)
            alloc[d["id"]] += a; pool -= a
    rest = [d for d in open_ if not d["fixed"]]
    if strategy == "even":
        while pool > 1 and rest:
            share = max(1000, (pool / len(rest)) // 1000 * 1000)
            for d in rest:
                a = min(share, d["left_uzs"] - alloc[d["id"]], pool)
                alloc[d["id"]] += a; pool -= a
            rest = [d for d in rest if d["left_uzs"] - alloc[d["id"]] > 0.5]
    else:
        key = (lambda d: d["left_uzs"]) if strategy == "small" else (lambda d: d["due"] or date(9999, 1, 1))
        for d in sorted(rest, key=key):
            a = min(pool, d["left_uzs"] - alloc[d["id"]])
            alloc[d["id"]] += a; pool -= a
    return alloc

def paid_in_month_uzs(debts, month_ym: int) -> float:
    s = 0.0
    for d in debts:
        r = d["rate"] if d["cur"] == "USD" else 1
        s += sum(p["amount"] * r for p in d["payments"] if ym(p["date"]) == month_ym)
    return s

def free_month(debts, monthly: float, now_ym: int):
    left = sum(d["left_uzs"] for d in debts)
    if left <= 0.5:
        return now_ym
    if monthly <= 0:
        return None
    left -= max(0.0, monthly - paid_in_month_uzs(debts, now_ym))
    m = now_ym
    while left > 0.5:
        m += 1; left -= monthly
        if m - now_ym > 600:
            return None
    return m

def ym_label(m): return f"{MONTHS_IN[m % 12]} {m // 12}"

# ─── API ──────────────────────────────────────────────────
@app.get("/")
def health():
    return {"ok": True, "service": "Temir Daftar API", "version": 2}

@app.get("/api/state")
def get_state(tg=Depends(current_user)):
    conn = db()
    try:
        ensure_user(conn, tg)
        user, debts = load_state(conn, tg["id"])
    finally:
        conn.close()
    for d in debts:
        d.pop("telegram_id", None); d.pop("created_at", None)
        for p in d["payments"]:
            p.pop("telegram_id", None); p.pop("created_at", None)
    user.pop("created_at", None)
    return json.loads(json.dumps({"user": user, "debts": debts, "today": today().isoformat()}, default=jsonable))

@app.post("/api/debts")
def add_debt(d: DebtIn, tg=Depends(current_user)):
    if d.cur == "USD" and not d.rate:
        raise HTTPException(422, "Для долга в долларах укажите курс, по которому брали")
    conn = db()
    try:
        ensure_user(conn, tg)
        n = conn.execute("SELECT count(*) AS n FROM tk_debts WHERE telegram_id=%s", (tg["id"],)).fetchone()["n"]
        if n >= 200:
            raise HTTPException(400, "Слишком много долгов")
        row = conn.execute("""INSERT INTO tk_debts(telegram_id,who,kind,cur,amount,rate,taken,due,fixed,pay_day,hue,note)
            VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
            (tg["id"], d.who.strip(), d.kind.strip(), d.cur, d.amount, d.rate if d.cur == "USD" else None,
             d.taken, d.due, d.fixed, d.pay_day, d.hue, d.note.strip())).fetchone()
        conn.commit()
        return {"ok": True, "id": row["id"]}
    finally:
        conn.close()

@app.put("/api/debts/{debt_id}")
def edit_debt(debt_id: int, d: DebtIn, tg=Depends(current_user)):
    if d.cur == "USD" and not d.rate:
        raise HTTPException(422, "Для долга в долларах укажите курс, по которому брали")
    conn = db()
    try:
        cur = conn.execute("""UPDATE tk_debts SET who=%s,kind=%s,cur=%s,amount=%s,rate=%s,taken=%s,due=%s,
            fixed=%s,pay_day=%s,hue=%s,note=%s WHERE id=%s AND telegram_id=%s""",
            (d.who.strip(), d.kind.strip(), d.cur, d.amount, d.rate if d.cur == "USD" else None, d.taken, d.due,
             d.fixed, d.pay_day, d.hue, d.note.strip(), debt_id, tg["id"]))
        conn.commit()
        if cur.rowcount == 0:
            raise HTTPException(404, "Долг не найден")
        return {"ok": True}
    finally:
        conn.close()

@app.delete("/api/debts/{debt_id}")
def delete_debt(debt_id: int, tg=Depends(current_user)):
    conn = db()
    try:
        conn.execute("DELETE FROM tk_debts WHERE id=%s AND telegram_id=%s", (debt_id, tg["id"]))
        conn.commit()
        return {"ok": True}
    finally:
        conn.close()

@app.post("/api/payments")
def add_payment(p: PaymentIn, tg=Depends(current_user)):
    conn = db()
    try:
        d = conn.execute("SELECT cur FROM tk_debts WHERE id=%s AND telegram_id=%s", (p.debt_id, tg["id"])).fetchone()
        if not d:
            raise HTTPException(404, "Долг не найден")
        row = conn.execute("""INSERT INTO tk_payments(telegram_id,debt_id,date,amount,rate)
            VALUES(%s,%s,%s,%s,%s) RETURNING id""",
            (tg["id"], p.debt_id, p.date, p.amount, p.rate if d["cur"] == "USD" else None)).fetchone()
        conn.commit()
        return {"ok": True, "id": row["id"]}
    finally:
        conn.close()

@app.delete("/api/payments/{pay_id}")
def delete_payment(pay_id: int, tg=Depends(current_user)):
    conn = db()
    try:
        conn.execute("DELETE FROM tk_payments WHERE id=%s AND telegram_id=%s", (pay_id, tg["id"]))
        conn.commit()
        return {"ok": True}
    finally:
        conn.close()

@app.patch("/api/settings")
def update_settings(s: SettingsIn, tg=Depends(current_user)):
    fields = {k: v for k, v in s.model_dump().items() if v is not None}
    if "strategy" in fields and fields["strategy"] not in STRATEGIES:
        raise HTTPException(422, "Неизвестная стратегия")
    if "theme" in fields and fields["theme"] not in ("auto", "light", "dark"):
        raise HTTPException(422, "Неизвестная тема")
    if not fields:
        return {"ok": True}
    conn = db()
    try:
        ensure_user(conn, tg)
        sets = ", ".join(f"{k}=%s" for k in fields)  # keys come from the model, not the client
        conn.execute(f"UPDATE tk_users SET {sets} WHERE telegram_id=%s", (*fields.values(), tg["id"]))
        conn.commit()
        return {"ok": True}
    finally:
        conn.close()

@app.delete("/api/all")
def delete_all(tg=Depends(current_user)):
    conn = db()
    try:
        conn.execute("DELETE FROM tk_payments WHERE telegram_id=%s", (tg["id"],))
        conn.execute("DELETE FROM tk_debts WHERE telegram_id=%s", (tg["id"],))
        conn.execute("DELETE FROM tk_users WHERE telegram_id=%s", (tg["id"],))
        conn.commit()
        return {"ok": True}
    finally:
        conn.close()

# ─── Excel, sent to the user's chat by the bot ────────────
def build_excel(debts) -> bytes:
    wb = Workbook()
    head_fill = PatternFill("solid", fgColor="1B1D1F")
    head_font = Font(color="FFFFFF", bold=True)

    def header(ws, cols, widths):
        ws.append(cols)
        for i, w in enumerate(widths, 1):
            c = ws.cell(1, i); c.fill = head_fill; c.font = head_font
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    ws = wb.active; ws.title = "Долги"
    header(ws, ["Кому", "Тип", "Валюта", "Сумма", "Курс взятия", "Взял", "Закрыть до",
                "Отдано", "Осталось", "Осталось, сум", "Статус"], [22, 22, 9, 16, 13, 12, 12, 16, 16, 18, 12])
    for d in debts:
        ws.append([d["who"], d["kind"], d["cur"], d["amount"], d["rate"] or "", d["taken"], d["due"] or "",
                   d["paid_cur"], d["left_cur"], d["left_uzs"], "Закрыт" if d["closed"] else "Открыт"])
    for row in ws.iter_rows(min_row=2):
        for i in (3, 4, 7, 8, 9):
            row[i].number_format = "#,##0.##"
        for i in (5, 6):
            row[i].number_format = "DD.MM.YYYY"
    n = ws.max_row
    ws.append(["Итого", "", "", "", "", "", "", "", "", f"=SUM(J2:J{n})" if n > 1 else 0, ""])
    ws.cell(n + 1, 1).font = Font(bold=True); ws.cell(n + 1, 10).font = Font(bold=True)
    ws.cell(n + 1, 10).number_format = "#,##0"

    ws2 = wb.create_sheet("Платежи")
    header(ws2, ["Дата", "Кому", "Сумма", "Валюта", "Курс дня", "Сумов потрачено"], [12, 22, 16, 9, 12, 18])
    rows = [(p, d) for d in debts for p in d["payments"]]
    for p, d in sorted(rows, key=lambda x: x[0]["date"]):
        spent = p["amount"] * (p["rate"] or d["rate"] or 1) if d["cur"] == "USD" else p["amount"]
        ws2.append([p["date"], d["who"], p["amount"], d["cur"], p["rate"] or "", spent])
    for row in ws2.iter_rows(min_row=2):
        row[0].number_format = "DD.MM.YYYY"
        for i in (2, 4, 5):
            row[i].number_format = "#,##0.##"
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()

@app.post("/api/export")
def export_excel(tg=Depends(current_user)):
    conn = db()
    try:
        _, debts = load_state(conn, tg["id"])
    finally:
        conn.close()
    data = build_excel(debts)
    name = f"temir_daftar_{today().strftime('%Y-%m-%d')}.xlsx"
    r = httpx.post(f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument",
                   data={"chat_id": tg["id"], "caption": "Все долги и платежи за всё время"},
                   files={"document": (name, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                   timeout=20)
    if not r.json().get("ok"):
        raise HTTPException(502, "Бот не смог отправить файл. Нажмите Start в чате с ботом.")
    return {"ok": True}

# ─── Bot ──────────────────────────────────────────────────
def tg_call(method: str, **payload):
    try:
        return httpx.post(f"https://api.telegram.org/bot{BOT_TOKEN}/{method}", json=payload, timeout=15).json()
    except Exception as e:
        return {"ok": False, "description": str(e)}

def app_button(text: str, hash_: str = ""):
    return {"text": text, "web_app": {"url": APP_URL + (("#" + hash_) if hash_ else "")}}

def send(chat_id, text, buttons=None):
    payload = {"chat_id": chat_id, "text": text, "parse_mode": "HTML"}
    if buttons:
        payload["reply_markup"] = {"inline_keyboard": buttons}
    return tg_call("sendMessage", **payload)

def send_animated(chat_id, lines: List[str], buttons=None):
    """Show the report building up line by line (Bot API drafts), then send it for real."""
    draft_id = int(time.time()) % 2_000_000_000 or 1
    shown = []
    for line in lines:
        shown.append(line)
        if not line.strip():
            continue
        res = tg_call("sendMessageDraft", chat_id=chat_id, draft_id=draft_id, text="\n".join(shown), parse_mode="HTML")
        if not res.get("ok"):
            break
        time.sleep(0.45)
    return send(chat_id, "\n".join(lines), buttons)

@app.post("/bot/webhook")
async def webhook(request: Request, x_telegram_bot_api_secret_token: str = Header(default="")):
    if not hmac.compare_digest(x_telegram_bot_api_secret_token, WEBHOOK_SECRET):
        raise HTTPException(403)
    upd = await request.json()
    msg = upd.get("message") or {}
    chat = (msg.get("chat") or {}).get("id")
    text = (msg.get("text") or "").strip()
    if not chat or (msg.get("chat") or {}).get("type") != "private":
        return {"ok": True}
    if text.startswith("/start"):
        send(chat, "<b>Temir Daftar</b> — учёт долгов.\n\nЗапиши долги, отмечай платежи, а я напомню 10 и 20 числа, "
                   "сколько осталось отдать в этом месяце.", [[app_button("Открыть Temir Daftar")]])
    elif text.startswith("/report"):
        conn = db()
        try:
            user, debts = load_state(conn, chat)
        finally:
            conn.close()
        if user and debts:
            send_animated(chat, summary_lines(user, debts, today()), [[app_button("Открыть приложение")]])
        else:
            send(chat, "Пока нет данных. Откройте приложение и добавьте первый долг.", [[app_button("Открыть Temir Daftar")]])
    elif text:
        send(chat, "Всё управление — в приложении.", [[app_button("Открыть Temir Daftar")]])
    return {"ok": True}

@app.get("/bot/setup")
def bot_setup(key: str = ""):
    if not BOT_TOKEN or not hmac.compare_digest(key, SETUP_KEY):
        raise HTTPException(403)
    a = tg_call("setWebhook", url=f"{API_URL}/bot/webhook", secret_token=WEBHOOK_SECRET,
                allowed_updates=["message"], drop_pending_updates=True)
    b = tg_call("setMyCommands", commands=[{"command": "start", "description": "Открыть Temir Daftar"},
                                           {"command": "report", "description": "Сводка по долгам"}])
    c = tg_call("setChatMenuButton", menu_button={"type": "web_app", "text": "Открыть", "web_app": {"url": APP_URL}})
    info = tg_call("getWebhookInfo")
    return {"webhook": a, "commands": b, "menu": c, "info": info}

# ─── Messages ─────────────────────────────────────────────
def reminder_lines(user, debts, t: date):
    open_ = [d for d in debts if not d["closed"]]
    plan = user["monthly_plan"] or 0
    paid = paid_in_month_uzs(debts, ym(t))
    left_month = max(0.0, plan - paid)
    lines = [f"<b>Напоминание · {t.day} {MONTHS_G[t.month - 1]}</b>", ""]
    if plan <= 0:
        lines.append("План на месяц не задан. Укажите в приложении, сколько готовы отдавать, и я посчитаю, кому и сколько.")
        return lines
    lines.append(f"План на {MONTHS[t.month - 1]}: {fmt(plan)} сум")
    lines.append(f"Отдано: {fmt(paid)} сум")
    if left_month <= 0.5:
        lines.append("✅ План этого месяца выполнен")
    else:
        lines.append(f"<b>Осталось отдать: {fmt(left_month)} сум</b>")
        alloc = split_month(open_, left_month, user["strategy"] or "small", ym(t))
        parts = [(d, alloc[d["id"]]) for d in open_ if alloc[d["id"]] > 0.5]
        if parts:
            lines += ["", "Как разделить:"]
            for d, a in parts:
                extra = f" (≈ {usd(a / d['rate'])})" if d["cur"] == "USD" else ""
                lines.append(f"• {esc(d['who'])} — {fmt(a)} сум{extra}")
    fm = free_month(debts, plan, ym(t))
    if fm is not None:
        lines += ["", f"Свободен от долгов в {ym_label(fm)}"]
    return lines

def esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def totals(debts):
    left = sum(d["left_uzs"] for d in debts)
    total = sum(d["amount"] * (d["rate"] if d["cur"] == "USD" else 1) for d in debts)
    return left, total

def report_lines(user, debts, t: date):
    """Report for the month that just ended (sent on the 1st)."""
    prev = ym(t) - 1
    paid_prev = paid_in_month_uzs(debts, prev)
    left, total = totals(debts)
    pct = round((total - left) / total * 100) if total else 0
    lines = [f"<b>Итоги: {MONTHS[prev % 12]} {prev // 12}</b>", "",
             f"Отдано за месяц: {fmt(paid_prev)} сум",
             f"Осталось всего: {fmt(left)} сум",
             f"Закрыто: {pct}% всех долгов"]
    closed = [esc(d["who"]) for d in debts if d["closed"] and d["payments"] and ym(d["payments"][-1]["date"]) == prev]
    if closed:
        lines.append("Закрыл полностью: " + ", ".join(closed) + " 🎉")
    return lines + free_line(user, debts, t, left)

def summary_lines(user, debts, t: date):
    """On-demand summary (/report)."""
    left, total = totals(debts)
    pct = round((total - left) / total * 100) if total else 0
    plan = user["monthly_plan"] or 0
    paid = paid_in_month_uzs(debts, ym(t))
    lines = [f"<b>Сводка на {t.day} {MONTHS_G[t.month - 1]}</b>", "",
             f"Осталось всего: {fmt(left)} сум",
             f"Закрыто: {pct}% всех долгов",
             f"Отдано в {MONTHS_IN[t.month - 1]}: {fmt(paid)} сум" + (f" из {fmt(plan)}" if plan else "")]
    return lines + free_line(user, debts, t, left)

def free_line(user, debts, t, left):
    if left <= 0.5:
        return ["", "<b>Ты свободен от долгов!</b>"]
    fm = free_month(debts, user["monthly_plan"] or 0, ym(t))
    if fm is None:
        return ["", "Задай сумму в месяц в разделе «План», и я посчитаю дату свободы"]
    return ["", f"<b>Свободен в {ym_label(fm)}</b> при {fmt(user['monthly_plan'])} сум в месяц"]

def due_lines(d, t: date):
    days = (d["due"] - t).days
    amt = usd(d["left_cur"]) if d["cur"] == "USD" else f"{fmt(d['left_cur'])} сум"
    return [f"<b>Срок долга через {days} {plural(days, 'день', 'дня', 'дней')}</b>", "",
            f"{esc(d['who'])}: осталось {amt}",
            f"Закрыть до {d['due'].day} {MONTHS_G[d['due'].month - 1]} {d['due'].year}"]

@app.get("/cron/daily")
def cron_daily(authorization: str = Header(default=""), day: Optional[int] = None, only: Optional[int] = None):
    if not CRON_SECRET or not hmac.compare_digest(authorization, f"Bearer {CRON_SECRET}"):
        raise HTTPException(401)
    t = today()
    if day:  # manual test: pretend it is this day of the current month
        t = t.replace(day=day)
    conn = db()
    sent = 0
    try:
        users = conn.execute("SELECT * FROM tk_users").fetchall()
        for u in users:
            if only and u["telegram_id"] != only:
                continue
            _, debts = load_state(conn, u["telegram_id"])
            if not debts:
                continue
            open_ = [d for d in debts if not d["closed"]]
            if t.day in (10, 20) and u["remind_on"] and open_:
                send(u["telegram_id"], "\n".join(reminder_lines(u, debts, t)),
                     [[app_button("Отметить оплату", "pay")], [app_button("Открыть приложение")]]); sent += 1
            if t.day == 1 and u["report_on"]:
                send_animated(u["telegram_id"], report_lines(u, debts, t), [[app_button("Открыть приложение")]]); sent += 1
            if u["due_alert_on"]:
                for d in open_:
                    if d["due"] and (d["due"] - t).days == 7:
                        send(u["telegram_id"], "\n".join(due_lines(d, t)),
                             [[app_button("Отметить оплату", f"debt-{d['id']}")]]); sent += 1
    finally:
        conn.close()
    return {"ok": True, "date": t.isoformat(), "sent": sent}
