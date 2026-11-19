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
    conn.execute("ALTER TABLE tk_users ADD COLUMN IF NOT EXISTS card_msg_id BIGINT")
    conn.execute("ALTER TABLE tk_users ADD COLUMN IF NOT EXISTS card_ym INTEGER")
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
    finally:
        conn.close()
    refresh_month_card(tg["id"])
    return {"ok": True, "id": row["id"]}

@app.delete("/api/payments/{pay_id}")
def delete_payment(pay_id: int, tg=Depends(current_user)):
    conn = db()
    try:
        conn.execute("DELETE FROM tk_payments WHERE id=%s AND telegram_id=%s", (pay_id, tg["id"]))
        conn.commit()
    finally:
        conn.close()
    refresh_month_card(tg["id"])
    return {"ok": True}

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
            print("sendMessageDraft failed:", res.get("description"))
            break
        time.sleep(0.45)
    return send(chat_id, "\n".join(lines), buttons)


# ─── Rich cards ───────────────────────────────────────────
EFFECT_PARTY = "5046509860389126442"  # 🎉

def mdx(s: str) -> str:
    return re.sub(r"([\\`*_\[\]|#>~])", r"\\\1", str(s))

def pbar(p: float, n: int = 10) -> str:
    k = max(0, min(n, round(p * n)))
    return "🟩" * k + "⬜" * (n - k)

def send_rich(chat_id, md: str, keyboard=None, effect=None, stream=False, fallback=None):
    """Stream the card block by block as a draft, then send it. Falls back to plain HTML."""
    if stream:
        blocks = md.split("\n\n")
        draft_id = int(time.time() * 10) % 2_000_000_000 or 1
        for i in range(1, len(blocks)):
            res = tg_call("sendRichMessageDraft", chat_id=chat_id, draft_id=draft_id, rich_message={"markdown": "\n\n".join(blocks[:i])})
            if not res.get("ok"):
                print("rich draft failed:", res.get("description"))
                break
            time.sleep(0.35)
    payload = {"chat_id": chat_id, "rich_message": {"markdown": md}}
    if keyboard:
        payload["reply_markup"] = {"inline_keyboard": keyboard}
    if effect:
        payload["message_effect_id"] = effect
    res = tg_call("sendRichMessage", **payload)
    if not res.get("ok"):
        print("sendRichMessage failed:", res.get("description"))
        if fallback:
            return send(chat_id, "\n".join(fallback), keyboard)
    return res

def edit_rich(chat_id, message_id, md: str, keyboard=None):
    payload = {"chat_id": chat_id, "message_id": message_id, "rich_message": {"markdown": md}}
    if keyboard is not None:
        payload["reply_markup"] = {"inline_keyboard": keyboard}
    return tg_call("editMessageText", **payload)

def month_card_md(user, debts, t: date) -> str:
    open_ = [d for d in debts if not d["closed"]]
    plan = user["monthly_plan"] or 0
    paid = paid_in_month_uzs(debts, ym(t))
    left_all = sum(d["left_uzs"] for d in debts)
    parts = [f"## План на {MONTHS[t.month - 1]}"]
    if not open_:
        return "\n\n".join(parts + ["🎉 **Открытых долгов нет. Ты свободен!**"])
    if plan <= 0:
        return "\n\n".join(parts + ["План не задан. Открой приложение → «План» и укажи, сколько готов отдавать в месяц."])
    p = min(1, paid / plan)
    parts.append(f"**{fmt(paid)}** из {fmt(plan)} сум")
    parts.append(f"{pbar(p)} {round(p * 100)}%")
    # per-debt: what the plan asked at the start of the month vs. paid this month
    start = []  # debts as they stood on the 1st, so the plan column does not shift as you pay
    for d in debts:
        r_ = d["rate"] if d["cur"] == "USD" else 1
        back = sum(x["amount"] for x in d["payments"] if ym(x["date"]) == ym(t))
        lc = d["left_cur"] + back
        if lc > 0.005:
            start.append({**d, "left_cur": lc, "left_uzs": lc * r_, "closed": False})
    base = split_month(start, plan, user["strategy"] or "small", -1)
    rows = ["| Кому | По плану | Отдано | |", "| --- | ---: | ---: | :---: |"]
    for d in sorted(debts, key=lambda d: d["due"] or date(9999, 1, 1)):
        r_ = d["rate"] if d["cur"] == "USD" else 1
        got = sum(x["amount"] * r_ for x in d["payments"] if ym(x["date"]) == ym(t))
        need = max(base.get(d["id"], 0.0), 0.0)
        if need < 0.5 and got < 0.5:
            continue
        mark = "✅" if got >= need - 0.5 else ("◐" if got > 0.5 else "⏳")
        rows.append(f"| {mdx(d['who'])} | {fmt(need)} | {fmt(got)} | {mark} |")
    if len(rows) > 2:
        parts.append("\n".join(rows))
    left_m = max(0.0, plan - paid)
    parts.append("✅ **План месяца выполнен**" if left_m <= 0.5 else f"Осталось в этом месяце: **{fmt(left_m)} сум**")
    fm = free_month(debts, plan, ym(t))
    parts.append("---")
    parts.append(f"> Всего осталось {fmt(left_all)} сум" + (f" · свободен в {ym_label(fm)}" if fm is not None else ""))
    return "\n\n".join(parts)

def month_card_kb():
    return [[app_button("Отметить оплату", "pay"), app_button("Открыть")]]

def payment_card_md(d, amount_cur, pdate, note, user, debts):
    t = today()
    when = "сегодня" if pdate == t else f"{pdate.day} {MONTHS_G[pdate.month - 1]}"
    pct = (d["amount"] - d["left_cur"]) / d["amount"]
    parts = ["## 🎉 Долг закрыт!" if d["closed"] else "## ✅ Платёж записан",
             f"**{mdx(d['who'])}** −{short_money(d, amount_cur)}, {when}" + (f"\n_{mdx(note)}_" if note else ""),
             f"{pbar(pct)} {round(pct * 100)}%"]
    rows = ["| | |", "| --- | ---: |",
            f"| Отдано по долгу | {short_money(d, d['paid_cur'])} |",
            f"| Осталось | {short_money(d, d['left_cur'])} |"]
    if d["cur"] == "USD" and not d["closed"]:
        rows.append(f"| Осталось в сумах | ≈ {fmt(d['left_uzs'])} |")
    plan = user["monthly_plan"] or 0
    if plan > 0:
        paid = paid_in_month_uzs(debts, ym(t))
        rows.append(f"| План на {MONTHS[t.month - 1]} | {fmt(paid)} из {fmt(plan)} |")
    parts.append("\n".join(rows))
    left = sum(x["left_uzs"] for x in debts)
    fm = free_month(debts, plan, ym(t))
    parts.append("> " + ("**Ты свободен от долгов!**" if left <= 0.5 else
                         (f"Свободен в {ym_label(fm)}" if fm is not None else f"Всего осталось {fmt(left)} сум")))
    return "\n\n".join(parts)

def debts_card_md(debts, t: date):
    open_ = sorted([d for d in debts if not d["closed"]], key=lambda d: d["due"] or date(9999, 1, 1))
    if not open_:
        return "## Мои долги\n\n🎉 **Открытых долгов нет. Ты свободен!**"
    rows = ["| Кому | Осталось | % | Срок |", "| --- | ---: | ---: | --- |"]
    for d in open_:
        pct = round((d["amount"] - d["left_cur"]) / d["amount"] * 100)
        due = "—"
        if d["due"]:
            due = "⚠️ просрочен" if d["due"] < t else f"{d['due'].day:02d}.{d['due'].month:02d}.{str(d['due'].year)[2:]}"
        rows.append(f"| {mdx(d['who'])} | {short_money(d, d['left_cur'])} | {pct}% | {due} |")
    left = sum(d["left_uzs"] for d in debts)
    total = sum(d["amount"] * (d["rate"] if d["cur"] == "USD" else 1) for d in debts)
    p = (total - left) / total if total else 0
    parts = [f"## Мои долги · {len(open_)}", "\n".join(rows),
             f"Всего осталось: **{fmt(left)} сум**", f"{pbar(p)} закрыто {round(p * 100)}%"]
    usd_ = [d for d in open_ if d["cur"] == "USD"]
    if usd_:
        parts.append("> Долларовые долги в сумах считаются по курсу, когда брал")
    return "\n\n".join(parts)

def report_card_md(user, debts, t: date, prev_month: bool = False):
    left, total = totals(debts)
    p = (total - left) / total if total else 0
    plan = user["monthly_plan"] or 0
    if prev_month:
        pm = ym(t) - 1
        title = f"## Итоги: {MONTHS[pm % 12]} {pm // 12}"
        paid = paid_in_month_uzs(debts, pm)
        paid_row = f"| Отдано за месяц | {fmt(paid)} сум |"
    else:
        title = f"## Сводка на {t.day} {MONTHS_G[t.month - 1]}"
        paid = paid_in_month_uzs(debts, ym(t))
        paid_row = f"| Отдано в {MONTHS_IN[t.month - 1]} | {fmt(paid)}" + (f" из {fmt(plan)}" if plan else "") + " сум |"
    rows = ["| | |", "| --- | ---: |", paid_row, f"| Осталось всего | {fmt(left)} сум |",
            f"| Открытых долгов | {len([d for d in debts if not d['closed']])} |"]
    parts = [title, f"{pbar(p)} закрыто {round(p * 100)}%", "\n".join(rows)]
    if prev_month:
        closed = [mdx(d["who"]) for d in debts if d["closed"] and d["payments"] and ym(d["payments"][-1]["date"]) == ym(t) - 1]
        if closed:
            parts.append("🎉 Закрыл полностью: " + ", ".join(closed))
    fm = free_month(debts, plan, ym(t))
    parts.append("> " + ("**Ты свободен от долгов!**" if left <= 0.5 else
                         (f"**Свободен в {ym_label(fm)}** при {fmt(plan)} сум в месяц" if fm is not None else
                          "Задай сумму в месяц в разделе «План», и я посчитаю дату свободы")))
    return "\n\n".join(parts)


# ─── Card images ──────────────────────────────────────────
import cards

def tg_photo(chat_id, png: bytes, caption: str, keyboard=None, message_id=None, effect=None):
    """Send a card image, or replace the image of an existing message."""
    try:
        files = {"card": ("card.png", png, "image/png")}
        if message_id:
            data = {"chat_id": chat_id, "message_id": message_id,
                    "media": json.dumps({"type": "photo", "media": "attach://card", "caption": caption, "parse_mode": "HTML"})}
            method = "editMessageMedia"
        else:
            data = {"chat_id": chat_id, "caption": caption, "parse_mode": "HTML"}
            files = {"photo": ("card.png", png, "image/png")}
            method = "sendPhoto"
            if effect:
                data["message_effect_id"] = effect
        if keyboard is not None:
            data["reply_markup"] = json.dumps({"inline_keyboard": keyboard})
        res = httpx.post(f"https://api.telegram.org/bot{BOT_TOKEN}/{method}", data=data, files=files, timeout=30).json()
        if not res.get("ok"):
            print(method, "failed:", res.get("description"))
        return res
    except Exception as e:
        print("tg_photo error:", e)
        return {"ok": False, "description": str(e)}

def month_rows(user, debts, t):
    plan = user["monthly_plan"] or 0
    start = []
    for d in debts:
        r_ = d["rate"] if d["cur"] == "USD" else 1
        back = sum(x["amount"] for x in d["payments"] if ym(x["date"]) == ym(t))
        lc = d["left_cur"] + back
        if lc > 0.005:
            start.append({**d, "left_cur": lc, "left_uzs": lc * r_, "closed": False})
    base = split_month(start, plan, user["strategy"] or "small", -1) if plan > 0 else {}
    rows = []
    for d in sorted(debts, key=lambda d: d["due"] or date(9999, 1, 1)):
        r_ = d["rate"] if d["cur"] == "USD" else 1
        got = sum(x["amount"] * r_ for x in d["payments"] if ym(x["date"]) == ym(t))
        need = base.get(d["id"], 0.0)
        if need < 0.5 and got < 0.5:
            continue
        rows.append({"who": d["who"], "hue": d["hue"], "need": need, "need_s": fmt(need), "got": got, "got_s": fmt(got),
                     "sub": "сверх плана" if need < 0.5 else ""})
    return rows

def month_card_png(user, debts, t):
    plan = user["monthly_plan"] or 0
    paid = paid_in_month_uzs(debts, ym(t))
    left_all = sum(d["left_uzs"] for d in debts)
    fm = free_month(debts, plan, ym(t))
    rows = month_rows(user, debts, t)
    free = "уже сейчас" if left_all <= 0.5 else (f"в {ym_label(fm)}" if fm is not None else "задай план")
    return cards.month_card({"title": f"План на {MONTHS[t.month - 1]}", "period": f"{MONTHS[t.month - 1].capitalize()} {t.year}",
                             "paid": paid, "paid_s": fmt(paid), "plan": plan, "plan_s": fmt(plan), "rows": rows[:8],
                             "left_month": max(0.0, plan - paid), "left_month_s": fmt(max(0.0, plan - paid)),
                             "left_all_s": fmt(left_all), "free": free})

def month_caption(user, debts, t):
    plan = user["monthly_plan"] or 0
    paid = paid_in_month_uzs(debts, ym(t))
    left = max(0.0, plan - paid)
    return (f"<b>План на {MONTHS[t.month - 1]}</b>: {fmt(paid)} из {fmt(plan)} сум"
            + (" · выполнен ✅" if left <= 0.5 else f" · осталось {fmt(left)}"))

def payment_card_png(d, amount_cur, pdate, note, user, debts):
    t = today()
    plan = user["monthly_plan"] or 0
    paid = paid_in_month_uzs(debts, ym(t))
    left = sum(x["left_uzs"] for x in debts)
    fm = free_month(debts, plan, ym(t))
    free = "Ты свободен от долгов!" if left <= 0.5 else (f"Свободен в {ym_label(fm)}" if fm is not None else f"Всего осталось {fmt(left)} сум")
    is_usd = d["cur"] == "USD"
    return cards.payment_card({
        "who": d["who"], "hue": d["hue"], "amount_s": short_money(d, amount_cur),
        "when": "сегодня" if pdate == t else f"{pdate.day} {MONTHS_G[pdate.month - 1]}",
        "closed": d["closed"], "pct": (d["amount"] - d["left_cur"]) / d["amount"], "paid_s": short_money(d, d["paid_cur"]),
        "left_s": (usd(d["left_cur"]) if is_usd else fmt(d["left_cur"])),
        "left_alt": (f"≈ {fmt(d['left_uzs'])} сум" if is_usd and not d["closed"] else ("" if d["closed"] else "сум")),
        "plan_val": fmt(paid) if plan > 0 else "не задан",
        "plan_sub": f"из {fmt(plan)} · {min(100, round(paid / plan * 100))}%" if plan > 0 else "",
        "free": free, "note": note})

def debts_card_png(debts, t):
    open_ = sorted([d for d in debts if not d["closed"]], key=lambda d: d["due"] or date(9999, 1, 1))
    left, total = totals(debts)
    rows = []
    for d in open_[:10]:
        pct = (d["amount"] - d["left_cur"]) / d["amount"]
        sub = (f"{d['kind']} · " if d["kind"] else "") + f"{round(pct * 100)}%" + (f" · ≈ {fmt(d['left_uzs'])} сум" if d["cur"] == "USD" else "")
        due, late, soon = "без срока", False, False
        if d["fixed"] and d["pay_day"]:
            due = f"{d['pay_day']}-го числа"
        if d["due"]:
            days = (d["due"] - t).days
            late, soon = days < 0, 0 <= days <= 45
            due = "просрочен" if late else (f"до {d['due'].day} {MSHORT_RU[d['due'].month - 1]}" if soon else f"до {MSHORT_RU[d['due'].month - 1]} {d['due'].year}")
        rows.append({"who": d["who"], "hue": d["hue"], "left_s": short_money(d, d["left_cur"]), "sub": sub, "pct": pct,
                     "due": due, "late": late, "soon": soon})
    return cards.debts_card({"count": len(open_), "left_all_s": fmt(left), "pct": (total - left) / total if total else 0,
                             "date": f"{t.day} {MONTHS_G[t.month - 1]}", "rows": rows})

MSHORT_RU = ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]

def post_month_card(uid, user, debts, t, stream=True):
    tg_call("sendChatAction", chat_id=uid, action="upload_photo")
    res = tg_photo(uid, month_card_png(user, debts, t), month_caption(user, debts, t), month_card_kb())
    if not res.get("ok"):
        return send_rich(uid, month_card_md(user, debts, t), month_card_kb(), stream=False, fallback=plan_lines(user, debts, t))
    new_id = res["result"]["message_id"]
    if user.get("card_msg_id") and user["card_msg_id"] != new_id:
        tg_call("unpinChatMessage", chat_id=uid, message_id=user["card_msg_id"])
    tg_call("pinChatMessage", chat_id=uid, message_id=new_id, disable_notification=True)
    conn = db()
    try:
        conn.execute("UPDATE tk_users SET card_msg_id=%s, card_ym=%s WHERE telegram_id=%s", (new_id, ym(t), uid))
        conn.commit()
    finally:
        conn.close()
    return res

def refresh_month_card(uid):
    try:
        conn = db()
        try:
            user, debts = load_state(conn, uid)
        finally:
            conn.close()
        t = today()
        if not user or not user.get("card_msg_id") or user.get("card_ym") != ym(t):
            return
        res = tg_photo(uid, month_card_png(user, debts, t), month_caption(user, debts, t), month_card_kb(), message_id=user["card_msg_id"])
        if not res.get("ok"):  # older live card was a text card
            edit_rich(uid, user["card_msg_id"], month_card_md(user, debts, t), month_card_kb())
    except Exception as e:
        print("refresh card failed:", e)

# ─── Bot: chat payments and commands ──────────────────────
import re

WELCOME = ("<b>Temir Daftar</b> — тетрадь твоих долгов.\n\n"
           "Записывай, кому и сколько должен, и отмечай платежи. Я посчитаю, когда ты закроешь все долги, "
           "и подскажу, как делить деньги каждый месяц.\n\n"
           "Напоминания и отчёты включаются в настройках приложения.\n"
           "Все быстрые команды и примеры — /help")

HELP = "\n".join([
    "💰 <b>Отдал часть долга — напиши мне одной строкой:</b>",
    "кому и сколько",
    "<code>Акмаль 500 000</code>",
    "<i>→ запишу платёж Акмалю, 500 000 сум, сегодня</i>",
    "",
    "📅 <b>Отдал в другой день</b> — добавь дату:",
    "<code>Акмаль 500к 12.11</code>",
    "",
    "💵 <b>Долг в долларах:</b>",
    "<code>Джасур 200$</code>",
    "или сумами по курсу того дня:",
    "<code>Джасур 2,6 млн по 12 900</code>",
    "",
    "Сумму можно писать как удобно: <code>5 млн</code>, <code>5м</code>, <code>500к</code> или <code>5 000 000</code>.",
    "Ошибся — нажми «↩️ Отменить» под ответом.",
    "",
    "<b>Команды</b>",
    "/plan — карточка месяца: кому и сколько отдать. Закрепится вверху чата и будет обновляться сама",
    "/dolgi — все долги с остатками",
    "/report — сводка: сколько отдано и когда свободен",
    "/help — эта подсказка",
])

COMMANDS = [
    {"command": "plan", "description": "Живая карточка месяца: кому и сколько"},
    {"command": "dolgi", "description": "Все долги с остатками"},
    {"command": "report", "description": "Сводка: отдано и когда свободен"},
    {"command": "help", "description": "Как записать платёж и все команды"},
    {"command": "start", "description": "Открыть Temir Daftar"},
]

MULT = {"млн": 1e6, "миллион": 1e6, "mln": 1e6, "м": 1e6, "m": 1e6, "тыс": 1e3, "т": 1e3, "к": 1e3, "k": 1e3}
NUM = r"(\d{1,3}(?:[  ]\d{3})+|\d+(?:[.,]\d+)?)"

def _num(s: str) -> float:
    return float(s.replace(" ", "").replace(" ", "").replace(",", "."))

def parse_payment_text(text: str, t: date):
    """Pull date, rate, amount and the remaining words (the name) out of a chat message."""
    s = " " + text.lower().replace("ё", "е") + " "
    out = {"date": t, "rate": None, "usd": False, "amount": None}
    md = re.search(r"\s(\d{1,2})[./](\d{1,2})(?:[./](\d{2,4}))?(?=\s)", s)
    if md:
        d_, m_, y_ = int(md.group(1)), int(md.group(2)), md.group(3)
        y = int(y_) + (2000 if y_ and len(y_) == 2 else 0) if y_ else t.year
        try:
            dd = date(y, m_, d_)
            if not y_ and dd > t:
                dd = date(y - 1, m_, d_)
            out["date"] = dd
        except ValueError:
            return {"error": "Не понял дату. Пиши так: <code>12.11</code> или <code>12.11.2026</code>"}
        s = s[:md.start()] + " " + s[md.end():]
    mr = re.search(r"\s(?:по|курс)\s*" + NUM + r"(?=\s)", s)
    if mr:
        out["rate"] = _num(mr.group(1))
        s = s[:mr.start()] + " " + s[mr.end():]
    ma = re.search(r"(\$)?\s*" + NUM + r"\s*(млн|миллион\w*|mln|тыс\w*|м|m|т|к|k)?(?![a-zа-я0-9])\s*(\$|долл\w*|usd|у\.?е\.?)?", s)
    if not ma:
        return {"error": None}
    v = _num(ma.group(2))
    suf = (ma.group(3) or "")
    for k, mult in MULT.items():
        if suf.startswith(k):
            v *= mult
            break
    out["amount"] = v
    out["usd"] = bool(ma.group(1) or ma.group(4))
    s = s[:ma.start()] + " " + s[ma.end():]
    out["name"] = re.sub(r"[^\w\s-]", " ", s).strip()
    return out

def match_debts(name: str, debts):
    words = [w for w in re.split(r"\s+", name) if len(w) >= 2]
    scored = []
    for d in debts:
        who = d["who"].lower().replace("ё", "е")
        sc = 0
        if name and name == who:
            sc = 1000
        elif name and (name in who or who in name):
            sc = 100
        else:
            dw = [w for w in re.split(r"[^\w]+", who) if len(w) >= 2]
            for w in words:
                for x in dw:
                    if x == w:
                        sc = max(sc, 50 + len(w))
                    elif len(w) >= 3 and (x.startswith(w) or w.startswith(x)):
                        sc = max(sc, 20 + min(len(w), len(x)))
        if sc:
            scored.append((sc, d))
    if not scored:
        return []
    top = max(s for s, _ in scored)
    return [d for s, d in scored if s == top]

def payment_from(d, amount, usd, rate, uzs_input: bool):
    """Return (amount in debt currency, rate stored, note) or error text."""
    if d["cur"] == "UZS":
        if usd:
            return None, None, "Этот долг в сумах. Напиши сумму в сумах."
        return amount, None, ""
    if usd or (not rate and amount < 10000):
        return amount, rate, ""
    r = rate or d["rate"]
    note = "" if rate else f"по курсу взятия {fmt(d['rate'])}"
    return round(amount / r, 2), r, note

def short_money(d, v):
    return usd(v) if d["cur"] == "USD" else f"{fmt(v)} сум"

def record_payment(chat, d, amount_cur, rate, pdate, note=""):
    if amount_cur <= 0:
        return send(chat, "Сумма должна быть больше нуля.")
    if amount_cur > d["left_cur"] + 0.01:
        return send(chat, f"Это больше остатка: по долгу «{esc(d['who'])}» осталось {short_money(d, d['left_cur'])}. Не записал.",
                    [[app_button("Открыть долг", f"debt-{d['id']}")]])
    if pdate > today():
        return send(chat, "Дата в будущем. Запиши платёж, когда отдашь.")
    conn = db()
    try:
        row = conn.execute("INSERT INTO tk_payments(telegram_id,debt_id,date,amount,rate) VALUES(%s,%s,%s,%s,%s) RETURNING id",
                           (chat, d["id"], pdate, amount_cur, rate if d["cur"] == "USD" else None)).fetchone()
        conn.commit()
        user, debts = load_state(conn, chat)
    finally:
        conn.close()
    nd = next(x for x in debts if x["id"] == d["id"])
    kb = [[{"text": "↩️ Отменить", "callback_data": f"undo:{row['id']}"}, app_button("Открыть долг", f"debt-{nd['id']}")]]
    when = "" if pdate == today() else f", {pdate.day} {MONTHS_G[pdate.month - 1]}"
    if nd["closed"]:
        line = f"🎉 <b>{esc(nd['who'])}</b> −{short_money(nd, amount_cur)}{when} · долг закрыт полностью"
    else:
        extra = f" (≈ {fmt(nd['left_uzs'])} сум)" if nd["cur"] == "USD" else ""
        line = f"✅ <b>{esc(nd['who'])}</b> −{short_money(nd, amount_cur)}{when} · осталось {short_money(nd, nd['left_cur'])}{extra}"
    if note:
        line += f"\n<i>{esc(note)}</i>"
    send(chat, line, [[{"text": "↩️ Отменить", "callback_data": f"undo:{row['id']}"}]])
    refresh_month_card(chat)

def plan_lines(user, debts, t: date):
    open_ = [d for d in debts if not d["closed"]]
    if not open_:
        return ["🎉 Открытых долгов нет. Ты свободен!"]
    plan = user["monthly_plan"] or 0
    if plan <= 0:
        return ["<b>План не задан</b>", "", "Открой приложение → «План» и укажи, сколько готов отдавать в месяц. Тогда я посчитаю, кому и сколько."]
    paid = paid_in_month_uzs(debts, ym(t))
    left = max(0.0, plan - paid)
    lines = [f"<b>План на {MONTHS[t.month - 1]}</b>", "",
             f"Отдано: {fmt(paid)} из {fmt(plan)} сум"]
    if left <= 0.5:
        lines.append("✅ План этого месяца выполнен")
    else:
        lines.append(f"<b>Осталось отдать: {fmt(left)} сум</b>")
        alloc = split_month(open_, left, user["strategy"] or "small", ym(t))
        parts = [(d, alloc[d["id"]]) for d in open_ if alloc[d["id"]] > 0.5]
        if parts:
            lines += ["", "Кому и сколько:"]
            for d, a in parts:
                extra = f" (≈ {usd(a / d['rate'])})" if d["cur"] == "USD" else ""
                lines.append(f"• {esc(d['who'])} — {fmt(a)} сум{extra}")
    return lines

def debts_lines(debts, t: date):
    open_ = sorted([d for d in debts if not d["closed"]], key=lambda d: d["due"] or date(9999, 1, 1))
    if not open_:
        return ["🎉 Открытых долгов нет. Ты свободен!"]
    lines = [f"<b>Мои долги · {len(open_)}</b>", ""]
    for d in open_:
        pct = round((d["amount"] - d["left_cur"]) / d["amount"] * 100)
        due = ""
        if d["due"]:
            days = (d["due"] - t).days
            due = " · ⚠️ просрочен" if days < 0 else f" · до {d['due'].day} {MONTHS_G[d['due'].month - 1]}" + (f" {d['due'].year}" if d["due"].year != t.year else "")
        extra = f" (≈ {fmt(d['left_uzs'])} сум)" if d["cur"] == "USD" else ""
        lines.append(f"• <b>{esc(d['who'])}</b> — {short_money(d, d['left_cur'])}{extra}")
        lines.append(f"   отдано {pct}%{due}")
    left = sum(d["left_uzs"] for d in debts)
    lines += ["", f"<b>Всего осталось: {fmt(left)} сум</b>"]
    return lines

def handle_text(chat, text):
    conn = db()
    try:
        user, debts = load_state(conn, chat)
    finally:
        conn.close()
    open_ = [d for d in debts if not d["closed"]]
    if not open_:
        return send(chat, "Сначала добавь долг в приложении, потом сможешь отмечать платежи прямо здесь.\n\nПодсказка — /help",
                    [[app_button("Открыть Temir Daftar")]])
    p = parse_payment_text(text, today())
    if p.get("error"):
        return send(chat, p["error"])
    if not p.get("amount"):
        return send(chat, "Не нашёл сумму. Напиши так: <code>Акмаль 500 000</code>\n\nВсе примеры — /help")
    cands = match_debts(p["name"], open_)
    if len(cands) != 1:
        opts = cands or open_
        title = "Кому из них?" if cands else ("Не нашёл такое имя. Кому отдал?" if p["name"] else "Кому отдал?")
        rate = int(p["rate"]) if p["rate"] else 0
        kb = [[{"text": f"{d['who']} · осталось {short_money(d, d['left_cur'])}",
                 "callback_data": f"p:{d['id']}:{p['amount']:g}:{p['date'].isoformat()}:{rate}:{1 if p['usd'] else 0}"}] for d in opts[:8]]
        return send(chat, f"{title}\nСумма: {usd(p['amount']) if p['usd'] else fmt(p['amount']) + ' сум'}", kb)
    d = cands[0]
    amt, rate, note = payment_from(d, p["amount"], p["usd"], p["rate"], True)
    if amt is None:
        return send(chat, note)
    record_payment(chat, d, amt, rate, p["date"], note)

def handle_callback(cb):
    chat = cb["message"]["chat"]["id"]
    data = cb.get("data", "")
    tg_call("answerCallbackQuery", callback_query_id=cb["id"])
    if data.startswith("undo:"):
        pid = int(data[5:])
        conn = db()
        try:
            row = conn.execute("DELETE FROM tk_payments WHERE id=%s AND telegram_id=%s RETURNING id", (pid, chat)).fetchone()
            conn.commit()
        finally:
            conn.close()
        mid = cb["message"]["message_id"]
        md = "## ↩️ Платёж отменён\n\nСумма вернулась в остаток долга." if row else "## Этот платёж уже удалён"
        if cb["message"].get("photo"):
            tg_call("editMessageCaption", chat_id=chat, message_id=mid, parse_mode="HTML",
                    caption="↩️ <b>Платёж отменён.</b> Сумма вернулась в остаток долга." if row else "Этот платёж уже удалён.",
                    reply_markup={"inline_keyboard": [[app_button("Открыть приложение")]]})
        elif cb["message"].get("rich_message"):
            edit_rich(chat, mid, md, [[app_button("Открыть приложение")]])
        else:
            tg_call("editMessageText", chat_id=chat, message_id=mid, text="↩️ <b>Платёж отменён.</b> Сумма вернулась в остаток долга." if row else "Этот платёж уже удалён.", parse_mode="HTML")
        if row:
            refresh_month_card(chat)
        return
    if data.startswith("p:"):
        _, did, amount, d_, rate, is_usd = data.split(":")
        conn = db()
        try:
            _, debts = load_state(conn, chat)
        finally:
            conn.close()
        d = next((x for x in debts if x["id"] == int(did) and not x["closed"]), None)
        tg_call("editMessageReplyMarkup", chat_id=chat, message_id=cb["message"]["message_id"], reply_markup={"inline_keyboard": []})
        if not d:
            return send(chat, "Этот долг уже закрыт или удалён.")
        amt, r, note = payment_from(d, float(amount), is_usd == "1", float(rate) or None, True)
        if amt is None:
            return send(chat, note)
        record_payment(chat, d, amt, r, date.fromisoformat(d_), note)

@app.post("/bot/webhook")
async def webhook(request: Request, x_telegram_bot_api_secret_token: str = Header(default="")):
    if not hmac.compare_digest(x_telegram_bot_api_secret_token, WEBHOOK_SECRET):
        raise HTTPException(403)
    upd = await request.json()
    if upd.get("callback_query"):
        cb = upd["callback_query"]
        if (cb.get("message") or {}).get("chat", {}).get("type") == "private":
            handle_callback(cb)
        return {"ok": True}
    msg = upd.get("message") or {}
    chat = (msg.get("chat") or {}).get("id")
    text = (msg.get("text") or "").strip()
    if not chat or (msg.get("chat") or {}).get("type") != "private" or not text:
        return {"ok": True}
    cmd = text.split()[0].split("@")[0].lower() if text.startswith("/") else ""
    if cmd == "/start":
        send(chat, WELCOME, [[app_button("Открыть Temir Daftar")]])
    elif cmd == "/help":
        send(chat, HELP, [[app_button("Открыть приложение")]])
    elif cmd in ("/report", "/plan", "/dolgi", "/dolg"):
        conn = db()
        try:
            user, debts = load_state(conn, chat)
        finally:
            conn.close()
        if not user or not debts:
            send(chat, "Пока нет долгов. Добавь первый в приложении.", [[app_button("Открыть Temir Daftar")]])
        else:
            t = today()
            if cmd == "/plan":
                post_month_card(chat, user, debts, t)
            elif cmd == "/report":
                send(chat, "\n".join(summary_lines(user, debts, t)), [[app_button("Открыть приложение")]])
            else:
                kb = [[app_button("Отметить оплату", "pay"), app_button("Открыть")]]
                if not [d for d in debts if not d["closed"]] or not tg_photo(chat, debts_card_png(debts, t), "<b>Мои долги</b>", kb).get("ok"):
                    send_rich(chat, debts_card_md(debts, t), kb, fallback=debts_lines(debts, t))
    elif cmd:
        send(chat, "Не знаю такой команды.\n\n" + HELP)
    else:
        handle_text(chat, text)
    return {"ok": True}

@app.get("/bot/setup")
def bot_setup(key: str = ""):
    if not BOT_TOKEN or not hmac.compare_digest(key, SETUP_KEY):
        raise HTTPException(403)
    a = tg_call("setWebhook", url=f"{API_URL}/bot/webhook", secret_token=WEBHOOK_SECRET,
                allowed_updates=["message", "callback_query"], drop_pending_updates=True)
    b = tg_call("setMyCommands", commands=COMMANDS)
    c = tg_call("setChatMenuButton", menu_button={"type": "web_app", "text": "Открыть", "web_app": {"url": APP_URL}})
    names = tg_call("setMyName", name="Temir Daftar")
    about = tg_call("setMyShortDescription", short_description="Система учёта долгов, платежей и плана погашения")
    desc = tg_call("setMyDescription", description="Записывай долги в сумах и долларах, отмечай платежи и смотри, когда будешь свободен. "
                                                    "План подскажет, сколько отдавать каждый месяц, а напоминания можно включить в настройках.")
    try:
        import base64
        from avatar_data import AVATAR_JPG_B64
        photo = httpx.post(f"https://api.telegram.org/bot{BOT_TOKEN}/setMyProfilePhoto",
                           data={"photo": json.dumps({"type": "static", "photo": "attach://avatar"})},
                           files={"avatar": ("avatar.jpg", base64.b64decode(AVATAR_JPG_B64), "image/jpeg")}, timeout=30).json()
    except Exception as e:
        photo = {"ok": False, "description": str(e)}
    info = tg_call("getWebhookInfo")
    return {"webhook": a, "commands": b, "menu": c, "name": names, "about": about, "description": desc, "photo": photo, "info": info}

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
            if t.day in (1, 10, 20) and u["remind_on"] and open_:
                post_month_card(u["telegram_id"], u, debts, t, stream=False); sent += 1
            if t.day == 1 and u["report_on"]:
                send(u["telegram_id"], "\n".join(report_lines(u, debts, t)), [[app_button("Открыть приложение")]]); sent += 1
            if u["due_alert_on"]:
                for d in open_:
                    if d["due"] and (d["due"] - t).days == 7:
                        send(u["telegram_id"], "\n".join(due_lines(d, t)),
                             [[app_button("Отметить оплату", f"debt-{d['id']}")]]); sent += 1
    finally:
        conn.close()
    return {"ok": True, "date": t.isoformat(), "sent": sent}
