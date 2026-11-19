"""Card images for the Temir Daftar bot, drawn in the app's style.

Every function takes plain data and returns PNG bytes. Drawn at 2x and
scaled down so edges stay smooth.
"""
import io
import os
from PIL import Image, ImageDraw, ImageFont

S = 2                      # supersampling factor
W = 1080                   # output width
HERE = os.path.dirname(os.path.abspath(__file__))
_fonts = {}

GROUND = (243, 245, 247)
SURFACE = (255, 255, 255)
TEXT = (22, 24, 26)
MUTED = (86, 90, 95)
FAINT = (163, 167, 171)
RULE = (238, 240, 243)
TRACK = (234, 237, 241)
ACCENT = (29, 95, 214)
ACCENT_SOFT = (220, 241, 255)
ACCENT_INK = (10, 58, 128)
POS = (47, 122, 69)
POS_SOFT = (220, 250, 229)
POS_INK = (20, 83, 45)
WARN_SOFT = (255, 239, 214)
WARN_INK = (122, 58, 6)
NEU_SOFT = (238, 241, 244)
NEU_INK = (58, 61, 64)

HUES = {
    "blue": ((220, 241, 255), (10, 58, 128)), "green": ((220, 250, 229), (20, 83, 45)),
    "amber": ((255, 239, 214), (122, 58, 6)), "rose": ((255, 228, 236), (138, 21, 56)),
    "teal": ((213, 245, 242), (15, 79, 74)), "violet": ((239, 230, 255), (74, 35, 144)),
}


def font(weight: int, size: int):
    name = {500: "Medium", 600: "SemiBold", 700: "Bold", 800: "ExtraBold"}[weight]
    key = (name, size)
    if key not in _fonts:
        _fonts[key] = ImageFont.truetype(os.path.join(HERE, "fonts", f"Manrope_{weight}{name}.ttf"), size * S)
    return _fonts[key]


class Canvas:
    def __init__(self, height: int, bg=GROUND):
        self.h = height
        self.im = Image.new("RGB", (W * S, height * S), bg)
        self.d = ImageDraw.Draw(self.im)

    def rect(self, x, y, w, h, r, fill):
        self.d.rounded_rectangle([x * S, y * S, (x + w) * S, (y + h) * S], radius=r * S, fill=fill)

    def text(self, x, y, s, weight, size, fill, anchor="la"):
        self.d.text((x * S, y * S), s, font=font(weight, size), fill=fill, anchor=anchor)

    def width(self, s, weight, size):
        return self.d.textlength(s, font=font(weight, size)) / S

    def bar(self, x, y, w, h, p, fill, track=TRACK):
        self.rect(x, y, w, h, h / 2, track)
        if p > 0:
            self.rect(x, y, max(h, w * min(1, p)), h, h / 2, fill)

    def avatar(self, x, y, size, letter, hue):
        bg, ink = HUES.get(hue, HUES["blue"])
        self.rect(x, y, size, size, size * 0.32, bg)
        self.text(x + size / 2, y + size / 2 + 1, letter, 700, int(size * 0.42), ink, anchor="mm")

    def check(self, cx, cy, r, bg, ink):
        self.d.ellipse([(cx - r) * S, (cy - r) * S, (cx + r) * S, (cy + r) * S], fill=bg)
        k = r * 0.42
        self.d.line([((cx - k) * S, cy * S), ((cx - k * 0.25) * S, (cy + k * 0.75) * S), ((cx + k) * S, (cy - k * 0.7) * S)],
                    fill=ink, width=int(r * 0.22 * S), joint="curve")

    def pill(self, x, y, s, bg, ink, size=26, align="left"):
        w = self.width(s, 700, size) + 32
        if align == "right":
            x -= w
        self.rect(x, y, w, size + 22, (size + 22) / 2, bg)
        self.text(x + w / 2, y + (size + 22) / 2, s, 700, size, ink, anchor="mm")
        return w

    def png(self) -> bytes:
        out = self.im.resize((W, self.h), Image.LANCZOS)
        buf = io.BytesIO()
        out.save(buf, "PNG", optimize=True)
        return buf.getvalue()


_logo = None

def _logo_img(size):
    global _logo
    if _logo is None:
        import base64
        from avatar_data import AVATAR_JPG_B64
        _logo = Image.open(io.BytesIO(base64.b64decode(AVATAR_JPG_B64))).convert("RGB")
    im = _logo.resize((size, size), Image.LANCZOS)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, size - 1, size - 1], fill=255)
    return im, mask

def _brand(c: Canvas, x, y, right_text):
    im, mask = _logo_img(44 * S)
    c.im.paste(im, (int(x * S), int(y * S)), mask)
    c.text(x + 58, y + 22, "Temir Daftar", 700, 26, TEXT, anchor="lm")
    c.text(W - x, y + 22, right_text, 600, 26, MUTED, anchor="rm")


def month_card(d: dict) -> bytes:
    """d: title, period, paid, plan, rows[{who, hue, need, got}], left_month, left_all, free"""
    rows = d["rows"]
    pad, cw = 48, W - 96
    h = 520 + len(rows) * 132 + 210
    c = Canvas(h)
    _brand(c, pad, 44, d["period"])
    y = 120
    c.rect(pad, y, cw, h - y - 48, 40, SURFACE)
    x = pad + 44
    c.text(x, y + 48, d["title"], 700, 34, MUTED)
    paid_s = d["paid_s"]
    c.text(x, y + 96, paid_s, 800, 84, TEXT)
    c.text(x + c.width(paid_s, 800, 84) + 16, y + 140, f"из {d['plan_s']} сум", 600, 34, MUTED)
    p = d["paid"] / d["plan"] if d["plan"] else 0
    c.bar(x, y + 222, cw - 88, 22, p, POS)
    c.text(x, y + 266, f"{round(min(1, p) * 100)}% плана", 700, 28, POS_INK)
    status = "План выполнен" if d["left_month"] <= 0.5 else f"Осталось {d['left_month_s']} сум"
    c.pill(pad + cw - 44, y + 254, status, POS_SOFT if d["left_month"] <= 0.5 else NEU_SOFT,
           POS_INK if d["left_month"] <= 0.5 else NEU_INK, 26, align="right")
    yy = y + 350
    c.rect(x, yy - 18, cw - 88, 2, 1, RULE)
    for i, r in enumerate(rows):
        ry = yy + i * 132
        c.avatar(x, ry + 10, 84, r["who"][:1].upper(), r["hue"])
        c.text(x + 108, ry + 22, r["who"], 700, 36, TEXT)
        c.text(x + 108, ry + 70, r.get("sub") or f"по плану {r['need_s']}", 600, 28, MUTED)
        right = pad + cw - 44
        if r["got"] >= r["need"] - 0.5 and r["got"] > 0.5:
            c.check(right - 26, ry + 52, 26, POS_SOFT, POS)
            c.text(right - 70, ry + 52, r["got_s"], 800, 34, TEXT, anchor="rm")
        elif r["got"] > 0.5:
            c.pill(right, ry + 28, "частично", WARN_SOFT, WARN_INK, 24, align="right")
            c.text(right - c.width("частично", 700, 24) - 52, ry + 52, r["got_s"], 800, 34, TEXT, anchor="rm")
        else:
            c.pill(right, ry + 28, "ждёт", NEU_SOFT, NEU_INK, 24, align="right")
        if i < len(rows) - 1:
            c.rect(x + 108, ry + 124, cw - 196, 2, 1, RULE)
    fy = yy + len(rows) * 132 + 14
    c.rect(x, fy, cw - 88, 128, 28, ACCENT_SOFT)
    c.text(x + 32, fy + 24, "Свободен от долгов", 600, 26, ACCENT_INK)
    c.text(x + 32, fy + 92, d["free"], 800, 38, ACCENT_INK, anchor="lm")
    c.text(pad + cw - 44 - 32, fy + 24, "Осталось всего", 600, 26, ACCENT_INK, anchor="ra")
    c.text(pad + cw - 44 - 32, fy + 92, f"{d['left_all_s']} сум", 800, 34, ACCENT_INK, anchor="rm")
    return c.png()


def payment_card(d: dict) -> bytes:
    """d: who, hue, amount_s, when, closed, pct, paid_s, left_s, left_alt, plan_line, free, note"""
    pad, cw = 48, W - 96
    h = 880 + (40 if d.get("note") else 0)
    c = Canvas(h)
    _brand(c, pad, 44, d["when"])
    y = 120
    c.rect(pad, y, cw, h - y - 48, 40, SURFACE)
    x = pad + 44
    if d["closed"]:
        for cx, cy, col, rr in [(860, 190, (255, 207, 135), 10), (930, 240, (185, 208, 255), 8), (800, 250, (166, 232, 187), 9),
                                (960, 170, (255, 179, 200), 7), (890, 290, (211, 191, 255), 6), (760, 180, (147, 226, 217), 7)]:
            c.d.ellipse([(cx - rr) * S, (cy - rr) * S, (cx + rr) * S, (cy + rr) * S], fill=col)
    c.check(x + 30, y + 74, 30, POS_SOFT, POS)
    c.text(x + 78, y + 74, "Долг закрыт полностью" if d["closed"] else "Платёж записан", 800, 36, POS_INK, anchor="lm")
    c.avatar(x, y + 140, 96, d["who"][:1].upper(), d["hue"])
    c.text(x + 120, y + 152, d["who"], 700, 38, TEXT)
    c.text(x + 120, y + 204, "отдал", 600, 28, MUTED)
    c.text(pad + cw - 44, y + 188, f"−{d['amount_s']}", 800, 54, TEXT, anchor="rm")
    if d.get("note"):
        c.text(pad + cw - 44, y + 240, d["note"], 600, 24, MUTED, anchor="ra")
    by = y + 290 + (40 if d.get("note") else 0)
    _, ink = HUES.get(d["hue"], HUES["blue"])
    c.bar(x, by, cw - 88, 20, d["pct"], ink)
    c.text(x, by + 44, f"Отдано {d['paid_s']}", 700, 28, TEXT)
    c.text(pad + cw - 44, by + 44, f"{round(d['pct'] * 100)}%", 700, 28, MUTED, anchor="ra")
    ty = by + 112
    tw = (cw - 88 - 20) / 2
    for i, (lbl, val, sub) in enumerate([("Осталось по долгу", d["left_s"], d.get("left_alt", "")),
                                          ("План месяца", d.get("plan_val", "не задан"), d.get("plan_sub", ""))]):
        tx = x + i * (tw + 20)
        c.rect(tx, ty, tw, 168, 24, (242, 248, 253))
        c.text(tx + 26, ty + 26, lbl, 600, 26, MUTED)
        c.text(tx + 26, ty + 66, val, 800, 40, TEXT)
        if sub:
            c.text(tx + 26, ty + 140, sub, 600, 24, MUTED, anchor="lm")
    fy = ty + 194
    c.rect(x, fy, cw - 88, 96, 26, ACCENT_SOFT)
    c.text(x + 30, fy + 48, d["free"], 800, 32, ACCENT_INK, anchor="lm")
    return c.png()


def debts_card(d: dict) -> bytes:
    """d: count, left_all_s, pct, date, rows[{who, hue, left_s, sub, pct, due, late, soon}]"""
    rows = d["rows"]
    pad, cw = 48, W - 96
    h = 400 + len(rows) * 140 + 110
    c = Canvas(h)
    _brand(c, pad, 44, d["date"])
    y = 120
    c.rect(pad, y, cw, h - y - 48, 40, SURFACE)
    x = pad + 44
    c.text(x, y + 48, f"Мои долги · {d['count']}", 700, 34, MUTED)
    c.text(x, y + 96, d["left_all_s"], 800, 84, TEXT)
    c.text(x + c.width(d["left_all_s"], 800, 84) + 16, y + 140, "сум осталось", 600, 34, MUTED)
    c.bar(x, y + 222, cw - 88, 20, d["pct"], ACCENT)
    c.text(x, y + 264, f"Закрыто {round(d['pct'] * 100)}% всех долгов", 700, 28, ACCENT_INK)
    yy = y + 330
    c.rect(x, yy - 14, cw - 88, 2, 1, RULE)
    for i, r in enumerate(rows):
        ry = yy + i * 140
        c.avatar(x, ry + 14, 84, r["who"][:1].upper(), r["hue"])
        c.text(x + 108, ry + 22, r["who"], 700, 36, TEXT)
        c.text(pad + cw - 44, ry + 22, r["left_s"], 800, 36, TEXT, anchor="ra")
        c.text(x + 108, ry + 70, r["sub"], 600, 26, MUTED)
        if r["late"] or r["soon"]:
            c.pill(pad + cw - 44, ry + 64, r["due"], WARN_SOFT, WARN_INK, 22, align="right")
        else:
            c.text(pad + cw - 44, ry + 72, r["due"], 600, 26, MUTED, anchor="ra")
        _, ink = HUES.get(r["hue"], HUES["blue"])
        c.bar(x + 108, ry + 114, cw - 88 - 108, 8, r["pct"], ink)
    return c.png()
