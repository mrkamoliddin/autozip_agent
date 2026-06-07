"""
Email Agent - Yandex Mail → OpenAI tahlil → Telegram
Yangi xatlar kelganda tahlil qilib Telegramga yuboradi.
Javob tugmasi orqali mijozga to'g'ridan-to'g'ri pochta yuborish mumkin.
Preview + tasdiqlash + imzo bilan.

YANGI FUNKSIYALAR (v2):
- Agent ishga tushganda inbox ro'yxati ko'rsatiladi (7 kun / 30 kun tugmalari)
- Har bir xat bosilganda to'liq tarix + har xat uchun javob tugmasi
- Sahifama-sahifa navigatsiya (10 tadan, « » tugmalar)
- /inbox buyrug'i bilan istalgan vaqt ro'yxatni chaqirish

v3 TUZATISHLAR:
1. send_email_detail har doim pending qaytaradi (None emas)
2. cleanup_pending hist_ kalitlarini o'chirmaydi
3. int() konversiyalarida try/except — crash bo'lmaydi
4. SINCE filtriga +1 kun buffer — aniqroq sana oralig'i
5. safe_callback() — 64 bayt limitini kafolatlaydi
6. save_json tarix xatlari uchun faqat bir marta chaqiriladi
7. /inbox buyrug'ida keraksiz "yuklanmoqda" xabari olib tashlandi

FIX LOG (v1):
1. waiting_edit / waiting_reply konflikti tuzatildi
2. Ikki admin bir vaqtda bosishi (race condition) tuzatildi
3. Bo'sh preview_text tekshiruvi qo'shildi
4. last_update_id faylga saqlanadi (agent qayta ishlaganda)
5. seen_ids AVVAL saqlanadi, keyin xabar yuboriladi
6. IMAP search noto'g'ri format tuzatildi
7. Telegram Markdown maxsus belgilar ekranlanadi
8. cleanup_pending faqat kutayotmagan xatlarni o'chiradi
9. loop_count bilan tsikl hisoblagich tuzatildi
10. send_telegram retry barcha xatolar uchun ishlaydi
"""

import imaplib
import smtplib
import email
import time
import json
import os
import re
import logging
import requests
import sys
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from datetime import datetime, timedelta
from pathlib import Path

# BeautifulSoup HTML parsing uchun
try:
    from bs4 import BeautifulSoup
    BS4_AVAILABLE = True
except ImportError:
    BS4_AVAILABLE = False

# ─── .env O'QISH ──────────────────────────────────────────────────────────────
env_path = Path(__file__).parent / ".env"
if env_path.exists():
    with open(env_path) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, _, value = line.partition("=")
                os.environ.setdefault(key.strip(), value.strip())

# ─── LOGGING ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("agent.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger(__name__)

# ─── SOZLAMALAR ───────────────────────────────────────────────────────────────
YANDEX_EMAIL      = os.environ["YANDEX_EMAIL"]
YANDEX_PASSWORD   = os.environ["YANDEX_PASSWORD"]
TELEGRAM_TOKEN    = os.environ["TELEGRAM_TOKEN"]

# FIX 1: Bo'sh string filtrlash — verguldan oldin bo'sh joy bo'lsa ham ishlaydi
TELEGRAM_CHAT_IDS = [
    cid.strip()
    for cid in os.environ.get("TELEGRAM_CHAT_IDS", os.environ.get("TELEGRAM_CHAT_ID", "")).split(",")
    if cid.strip()
]

OPENAI_KEY        = os.environ["OPENAI_API_KEY"]
CHECK_INTERVAL    = int(os.environ.get("CHECK_INTERVAL", "60"))

SEEN_IDS_FILE       = "seen_ids.json"
PENDING_FILE        = "pending_replies.json"
LAST_UPDATE_FILE    = "last_update_id.json"   # FIX 4: last_update_id saqlash uchun
IMAP_HOST           = "imap.yandex.ru"
IMAP_PORT           = 993
SMTP_HOST           = "smtp.yandex.ru"
SMTP_PORT           = 465

SEEN_IDS_MAX        = 1000
PENDING_MAX_DAYS    = 30
OPENAI_MAX_RETRIES  = 5
OPENAI_RETRY_DELAY  = 3
INBOX_PAGE_SIZE     = 10   # Sahifada nechta xat ko'rsatilsin

LOGO_PATH = Path(__file__).parent / "autozip_logo_email.png"

EMAIL_SIGNATURE_TEXT = """
--
Askar Mukhamed
Director, AutoZIP
Tel.: +998 90 168 15 13
Email: info@autozip.uz
Website: www.autozip.uz
"""

EMAIL_SIGNATURE_HTML = """
<br><br>
<table style="border-top: 2px solid #cc0000; padding-top: 12px; font-family: Arial, sans-serif; font-size: 13px; color: #333;">
  <tr>
    <td style="padding-right: 20px; vertical-align: top;">
      <img src="cid:autozip_logo" width="180" alt="AutoZIP" style="display:block;">
    </td>
    <td style="vertical-align: top; border-left: 2px solid #cc0000; padding-left: 16px;">
      <strong style="font-size:15px; color:#cc0000;">Askar Mukhamed</strong><br>
      <span style="color:#555;">Director, AutoZIP</span><br><br>
      <span>&#128222; <a href="tel:+998901681513" style="color:#333; text-decoration:none;">+998 90 168 15 13</a></span><br>
      <span>&#9993; <a href="mailto:info@autozip.uz" style="color:#cc0000;">info@autozip.uz</a></span><br>
      <span>&#127760; <a href="https://www.autozip.uz" style="color:#cc0000;">www.autozip.uz</a></span>
    </td>
  </tr>
</table>
"""

# ─── YORDAMCHI ────────────────────────────────────────────────────────────────

def load_json(path: str, default):
    if Path(path).exists():
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            log.warning(f"JSON o'qishda xato ({path}): {e} — standart qiymat ishlatiladi")
    return default

def save_json(path: str, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError as e:
        log.error(f"JSON saqlashda xato ({path}): {e}")

def decode_str(value: str) -> str:
    if not value:
        return ""
    parts = decode_header(value)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(part)
    return " ".join(result)

def get_email_body(msg: email.message.Message) -> str:
    """Email tanasini olish — HTML dan BeautifulSoup bilan matn ajratish."""
    plain_body = ""
    html_body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            # FIX 9: text/plain bo'sh bo'lsa ham break qilmaymiz
            if ct == "text/plain" and not plain_body:
                charset = part.get_content_charset() or "utf-8"
                text = part.get_payload(decode=True).decode(charset, errors="replace").strip()
                if text:
                    plain_body = text
            elif ct == "text/html" and not html_body:
                charset = part.get_content_charset() or "utf-8"
                raw_html = part.get_payload(decode=True).decode(charset, errors="replace")
                if BS4_AVAILABLE:
                    soup = BeautifulSoup(raw_html, "html.parser")
                    html_body = soup.get_text(separator=" ", strip=True)
                else:
                    html_body = re.sub(r"<[^>]+>", " ", raw_html)
                    html_body = re.sub(r"\s+", " ", html_body).strip()
    else:
        charset = msg.get_content_charset() or "utf-8"
        plain_body = msg.get_payload(decode=True).decode(charset, errors="replace").strip()

    # plain_body ustunlik qiladi, yo'q bo'lsa html_body
    return (plain_body or html_body).strip()

def extract_email_address(sender: str) -> str:
    match = re.search(r'<(.+?)>', sender)
    if match:
        return match.group(1).strip()
    return sender.strip()

def clean_subject_for_reply(subject: str) -> str:
    if re.match(r'^re\s*:', subject.strip(), re.IGNORECASE):
        return subject
    return f"Re: {subject}"

# FIX 7: Telegram MarkdownV2 uchun maxsus belgilarni ekranlash
def escape_markdown(text: str) -> str:
    """Telegram MarkdownV2 uchun xavfli belgilarni ekranlash."""
    escape_chars = r'_*[]()~`>#+-=|{}.!'
    return re.sub(f'([{re.escape(escape_chars)}])', r'\\\1', text)

# FIX 8: cleanup_pending — faqat kutayotmagan xatlarni o'chiradi
def cleanup_pending(pending: dict) -> dict:
    """Eski pending yozuvlarni tozalash — lekin kutayotgan va hist_ xatlarni saqlab qoladi."""
    now = datetime.now()
    to_delete = []
    for eid, em in pending.items():
        # Kimdir javob kutayotgan bo'lsa — o'chirmaymiz
        if em.get("waiting_reply") or em.get("waiting_edit"):
            continue
        # FIX 2: hist_ prefiksli kalitlar — tarix xatlari, o'chirmaymiz
        if eid.startswith("hist_"):
            continue
        created = em.get("created_at")
        if created:
            try:
                created_dt = datetime.fromisoformat(created)
                days_old = (now - created_dt).days
                if days_old > PENDING_MAX_DAYS:
                    to_delete.append(eid)
            except ValueError:
                pass
    for eid in to_delete:
        log.info(f"Eski pending o'chirildi: {eid}")
        del pending[eid]
    return pending

def trim_seen_ids(seen_ids: set) -> set:
    if len(seen_ids) > SEEN_IDS_MAX:
        sorted_ids = sorted(seen_ids)
        seen_ids = set(sorted_ids[-SEEN_IDS_MAX:])
        log.info(f"seen_ids {SEEN_IDS_MAX} taga qisqartirildi")
    return seen_ids

# ─── OPENAI ───────────────────────────────────────────────────────────────────

def openai_request(prompt: str, max_tokens: int = 1500) -> str:
    url = "https://api.openai.com/v1/chat/completions"
    headers = {"Authorization": f"Bearer {OPENAI_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": "gpt-4o-mini",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens
    }

    last_error = None
    for attempt in range(1, OPENAI_MAX_RETRIES + 1):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=60)
            resp.raise_for_status()
            return resp.json()["choices"][0]["message"]["content"]
        except requests.exceptions.Timeout as e:
            last_error = e
            log.warning(f"OpenAI timeout, urinish {attempt}/{OPENAI_MAX_RETRIES}...")
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response else None
            if status and (status == 429 or status >= 500):
                last_error = e
                log.warning(f"OpenAI HTTP {status}, urinish {attempt}/{OPENAI_MAX_RETRIES}...")
            else:
                raise
        except Exception as e:
            last_error = e
            log.warning(f"OpenAI xato ({type(e).__name__}), urinish {attempt}/{OPENAI_MAX_RETRIES}...")

        if attempt < OPENAI_MAX_RETRIES:
            time.sleep(OPENAI_RETRY_DELAY * attempt)

    raise Exception(f"OpenAI {OPENAI_MAX_RETRIES} urinishdan keyin ham ishlamadi: {last_error}")

def analyze_email(sender: str, subject: str, body: str, history: list = None) -> str:
    history_text = ""
    if history:
        history_text = "\n\nПРЕДЫДУЩАЯ ПЕРЕПИСКА С ЭТИМ КЛИЕНТОМ:\n"
        for i, h in enumerate(history, 1):
            history_text += f"\n[Письмо {i}] Тема: {h['subject']}\n{h['body'][:500]}\n---"

    prompt = f"""Вы — бизнес-ассистент, который анализирует письма от потенциальных клиентов.
Прочитайте письмо и дайте владельцу бизнеса краткую и чёткую сводку на русском языке.
Учитывайте историю переписки с этим клиентом, если она есть.

ОТ КОГО: {sender}
ТЕМА: {subject}

ТЕКСТ ПИСЬМА:
{body[:3000]}
{history_text}

---
Ответьте строго в следующем формате:

🎯 ЦЕЛЬ: (что хочет клиент — 1-2 предложения)
📦 ИНТЕРЕС К ТОВАРУ: (какой товар/услуга его интересует)
💰 НАМЕРЕНИЕ КУПИТЬ: (Высокое / Среднее / Низкое — с обоснованием)
🔄 ИСТОРИЯ: (если клиент писал раньше — кратко что было, если нет — "Первое обращение")
❓ ОСНОВНЫЕ ВОПРОСЫ:
• ...
• ...
• ...
⚡ РЕКОМЕНДАЦИЯ: (что делать дальше — как ответить клиенту)"""
    return openai_request(prompt)

def generate_reply(original_body: str, original_subject: str, hint: str, history: list = None) -> str:
    history_text = ""
    if history:
        history_text = "\n\nПРЕДЫДУЩАЯ ПЕРЕПИСКА С ЭТИМ КЛИЕНТОМ:\n"
        for i, h in enumerate(history, 1):
            history_text += f"\n[Письмо {i}] Тема: {h['subject']}\n{h['body'][:500]}\n---"

    prompt = f"""Вы — менеджер по продажам компании AutoZIP (автозапчасти). Напишите профессиональный ответ клиенту НА РУССКОМ ЯЗЫКЕ.

- Тон: вежливый, профессиональный, дружелюбный
- Учитывайте историю переписки если она есть
- НЕ добавляйте подпись в конце — она будет добавлена автоматически
- Пишите ТОЛЬКО на русском языке

ИСХОДНОЕ ПИСЬМО КЛИЕНТА:
Тема: {original_subject}
{original_body[:3000]}
{history_text}

---
КРАТКИЕ ТЕЗИСЫ ДЛЯ ОТВЕТА (от менеджера):
{hint}

---
Напишите готовое письмо-ответ на русском языке. Только текст письма, без пояснений и без подписи."""
    return openai_request(prompt, max_tokens=1500)

def translate_reply(russian_text: str, original_body: str) -> str:
    prompt = f"""Определите язык следующего письма клиента и переведите готовый ответ на тот же язык.

ПИСЬМО КЛИЕНТА (для определения языка):
{original_body[:1000]}

---
ГОТОВЫЙ ОТВЕТ НА РУССКОМ (переведите на язык клиента):
{russian_text}

---
Важно:
- Если клиент писал на русском — оставьте текст БЕЗ ИЗМЕНЕНИЙ
- Если на узбекском — переведите на узбекский
- Если на английском — переведите на английский
- Если на другом языке — переведите на тот язык
- Возвращайте ТОЛЬКО переведённый текст, без пояснений"""
    return openai_request(prompt, max_tokens=1500)

# ─── SMTP ─────────────────────────────────────────────────────────────────────

def send_email_reply(to_address: str, subject: str, body: str):
    msg = MIMEMultipart("related")
    msg["From"]    = YANDEX_EMAIL
    msg["To"]      = to_address
    msg["Subject"] = clean_subject_for_reply(subject)

    html_body = body.replace("\n", "<br>")
    html_content = f"""
<html><body>
<div style="font-family: Arial, sans-serif; font-size: 14px; color: #222; line-height: 1.6;">
{html_body}
</div>
{EMAIL_SIGNATURE_HTML}
</body></html>"""

    alternative = MIMEMultipart("alternative")
    alternative.attach(MIMEText(body + EMAIL_SIGNATURE_TEXT, "plain", "utf-8"))
    alternative.attach(MIMEText(html_content, "html", "utf-8"))
    msg.attach(alternative)

    if LOGO_PATH.exists():
        with open(LOGO_PATH, "rb") as f:
            logo_img = MIMEImage(f.read())
            logo_img.add_header("Content-ID", "<autozip_logo>")
            logo_img.add_header("Content-Disposition", "inline", filename="autozip_logo.png")
            msg.attach(logo_img)

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(YANDEX_EMAIL, YANDEX_PASSWORD)
        server.sendmail(YANDEX_EMAIL, to_address, msg.as_string())
    log.info(f"Pochta yuborildi: {to_address}")

# ─── TELEGRAM ─────────────────────────────────────────────────────────────────

def send_telegram(text: str, chat_id: str = None, reply_markup: dict = None):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    max_len = 4000
    chunks = [text[i:i+max_len] for i in range(0, len(text), max_len)]
    targets = [chat_id] if chat_id else TELEGRAM_CHAT_IDS

    for cid in targets:
        for i, chunk in enumerate(chunks):
            payload = {
                "chat_id": cid,
                "text": chunk,
                "parse_mode": "Markdown"
            }
            if reply_markup and i == len(chunks) - 1:
                payload["reply_markup"] = reply_markup

            # FIX 10: Barcha xatolar uchun retry ishlaydi
            for attempt in range(3):
                try:
                    resp = requests.post(url, json=payload, timeout=30)
                    if not resp.ok:
                        log.error(f"Telegram xato ({cid}): {resp.text}")
                    break
                except requests.exceptions.Timeout:
                    log.warning(f"Telegram timeout, urinish {attempt+1}/3...")
                    time.sleep(5)
                except Exception as e:
                    log.error(f"Telegram xato (urinish {attempt+1}/3): {e}")
                    if attempt < 2:
                        time.sleep(3)
                    # break o'rniga davom etamiz — 3 marta urinib ko'ramiz

def format_telegram_message(sender: str, subject: str, analysis: str) -> str:
    now = datetime.now().strftime("%d.%m.%Y %H:%M")
    # FIX 7: sender va subject ichidagi maxsus belgilarni ekranlash
    safe_sender  = escape_markdown(sender)
    safe_subject = escape_markdown(subject)
    return (
        f"📧 *НОВОЕ ПИСЬМО ОТ КЛИЕНТА*\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 *От кого:* `{safe_sender}`\n"
        f"📌 *Тема:* {safe_subject}\n"
        f"🕐 *Время:* {now}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        f"{analysis}"
    )

# ─── INBOX RO'YXATI ───────────────────────────────────────────────────────────

def fetch_inbox_emails(days: int = 7) -> list:
    """
    Inboxdagi barcha xatlarni (o'qilgan + o'qilmagan) sana bo'yicha olish.
    Qaytaradi: [{"uid", "sender", "subject", "date", "seen", "date_obj"}, ...]
    Eng yangilar birinchi.
    """
    emails = []
    mail = None
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(YANDEX_EMAIL, YANDEX_PASSWORD)
        mail.select("INBOX")

        # FIX 4: +1 kun buffer — IMAP SINCE kun boshidan hisoblagani uchun
        since_date = (datetime.now() - timedelta(days=days + 1)).strftime("%d-%b-%Y")
        _, data = mail.search(None, f'SINCE {since_date}')
        ids = data[0].split()

        for uid in ids:
            uid_str = uid.decode()
            # Faqat sarlavha va flags — tezroq
            _, msg_data = mail.fetch(uid, "(FLAGS BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
            raw_headers = msg_data[0][1]
            msg = email.message_from_bytes(raw_headers)

            sender  = decode_str(msg.get("From", ""))
            subject = decode_str(msg.get("Subject", "(без темы)"))
            date_str = msg.get("Date", "")

            # O'qilgan/o'qilmagan
            flags_data = msg_data[0][0].decode() if isinstance(msg_data[0][0], bytes) else str(msg_data[0][0])
            seen = "\\Seen" in flags_data

            # Sana parse
            date_obj = None
            try:
                from email.utils import parsedate_to_datetime
                date_obj = parsedate_to_datetime(date_str)
            except Exception:
                date_obj = datetime.now()

            emails.append({
                "uid":      uid_str,
                "sender":   sender,
                "subject":  subject,
                "date":     date_str,
                "date_obj": date_obj,
                "seen":     seen,
            })

    except Exception as e:
        log.error(f"fetch_inbox_emails xato: {e}")
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass

    # Eng yangilar birinchi
    emails.sort(key=lambda x: x["date_obj"] or datetime.min, reverse=True)
    return emails


def fetch_email_full(uid_str: str) -> dict | None:
    """Bitta xatning to'liq tanasini olish (uid bo'yicha)."""
    mail = None
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(YANDEX_EMAIL, YANDEX_PASSWORD)
        mail.select("INBOX")
        _, msg_data = mail.fetch(uid_str.encode(), "(RFC822)")
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        sender  = decode_str(msg.get("From", ""))
        subject = decode_str(msg.get("Subject", "(без темы)"))
        date    = msg.get("Date", "")
        body    = get_email_body(msg)
        sender_addr = extract_email_address(sender)
        history = fetch_sender_history(mail, sender_addr)
        return {
            "uid": uid_str, "sender": sender, "subject": subject,
            "date": date, "body": body, "history": history
        }
    except Exception as e:
        log.error(f"fetch_email_full xato: {e}")
        return None
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass


def safe_callback(data: str) -> str:
    """Telegram callback_data 64 bayt limitini kafolatlash."""
    encoded = data.encode("utf-8")
    if len(encoded) > 64:
        # Oxirgi qismni kesib, xavfsiz qilamiz
        log.warning(f"callback_data 64 baytdan uzun, qisqartirildi: {data}")
        return encoded[:64].decode("utf-8", errors="ignore")
    return data


def send_inbox_list(chat_id: str, emails: list, page: int = 0, days: int = 7):
    """
    Inbox ro'yxatini Telegramga yuborish.
    Sahifama-sahifa: 10 tadan, « » navigatsiya tugmalari.
    """
    total = len(emails)
    if total == 0:
        send_telegram(
            f"📭 *Inbox bo'sh*\nOxirgi {days} kunda xat topilmadi.",
            chat_id=chat_id
        )
        return

    start = page * INBOX_PAGE_SIZE
    end   = min(start + INBOX_PAGE_SIZE, total)
    page_emails = emails[start:end]
    total_pages = (total + INBOX_PAGE_SIZE - 1) // INBOX_PAGE_SIZE

    unseen_count = sum(1 for e in emails if not e["seen"])
    lines = [
        f"📬 *INBOX — oxirgi {days} kun*",
        f"Jami: *{total}* xat  |  O'qilmagan: *{unseen_count}*",
        f"Sahifa {page+1}/{total_pages}",
        "━━━━━━━━━━━━━━━━━━━━",
    ]
    for i, em in enumerate(page_emails, start=start+1):
        icon    = "🔴" if not em["seen"] else "⚪"
        try:
            d = em["date_obj"].strftime("%d.%m %H:%M") if em["date_obj"] else "—"
        except Exception:
            d = "—"
        sender_short = escape_markdown(extract_email_address(em["sender"]))
        subject_short = escape_markdown(em["subject"][:40])
        lines.append(f"{icon} *{i}.* {subject_short}\n    👤 `{sender_short}` · {d}")

    text = "\n".join(lines)

    # Inline keyboard: har xat uchun tugma + navigatsiya
    keyboard = []
    for i, em in enumerate(page_emails, start=start+1):
        icon = "🔴" if not em["seen"] else "⚪"
        btn_label = f"{icon} {i}. {em['subject'][:30]}"
        cb = safe_callback(f"view_email:{em['uid']}:{days}:{page}")  # FIX 5
        keyboard.append([{"text": btn_label, "callback_data": cb}])

    # Navigatsiya qatori
    nav_row = []
    if page > 0:
        nav_row.append({"text": "« Oldingi", "callback_data": f"inbox_page:{days}:{page-1}"})
    if end < total:
        nav_row.append({"text": "Keyingi »", "callback_data": f"inbox_page:{days}:{page+1}"})
    if nav_row:
        keyboard.append(nav_row)

    # Sana filtri tugmalari
    keyboard.append([
        {"text": "📅 7 kun", "callback_data": f"inbox_page:7:0"},
        {"text": "📅 30 kun", "callback_data": f"inbox_page:30:0"},
    ])

    send_telegram(text, chat_id=chat_id, reply_markup={"inline_keyboard": keyboard})


def send_email_detail(chat_id: str, uid_str: str, back_days: int = 7, back_page: int = 0, pending: dict = None):
    """
    Xatni to'liq ko'rsatish: tarix + har xat uchun javob tugmasi.
    Har doim pending dict qaytaradi (hech qachon None emas).
    """
    if pending is None:
        pending = {}

    send_telegram("⏳ *Xat yuklanmoqda...*", chat_id=chat_id)
    em = fetch_email_full(uid_str)
    if not em:
        send_telegram("❌ *Xatni yuklashda xato.*", chat_id=chat_id)
        return pending  # FIX 1: None emas, pending qaytaramiz

    sender_safe  = escape_markdown(em["sender"])
    subject_safe = escape_markdown(em["subject"])

    # Asosiy xat
    body_preview = em["body"][:2000] if em["body"] else "_(bo'sh)_"
    msg_text = (
        f"📧 *XAT TAFSILOTI*\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 *Kimdan:* `{sender_safe}`\n"
        f"📌 *Mavzu:* {subject_safe}\n"
        f"🕐 *Sana:* {escape_markdown(em['date'][:30])}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        f"{body_preview}"
    )

    # Asosiy xat uchun tugmalar
    # pending ga qo'shib qo'yamiz (agar yo'q bo'lsa)
    if uid_str not in pending:
        pending[uid_str] = {
            "sender":     em["sender"],
            "subject":    em["subject"],
            "body":       em["body"],
            "history":    em.get("history", []),
            "created_at": datetime.now().isoformat()
        }
        save_json(PENDING_FILE, pending)

    main_keyboard = {
        "inline_keyboard": [
            [
                {"text": "✍️ Javob yozish", "callback_data": f"reply:{uid_str}"},
            ],
            [
                {"text": "◀ Orqaga", "callback_data": f"inbox_page:{back_days}:{back_page}"}
            ]
        ]
    }
    send_telegram(msg_text, chat_id=chat_id, reply_markup=main_keyboard)

    # Tarix xatlarini ham ko'rsatamiz
    history = em.get("history", [])
    if history:
        send_telegram(
            f"🔄 *Bu mijozning avvalgi {len(history)} ta xati:*",
            chat_id=chat_id
        )
        pending_changed = False
        for i, h in enumerate(history, 1):
            h_subject = escape_markdown(h.get("subject", ""))
            h_date    = escape_markdown(str(h.get("date", ""))[:30])
            h_body    = h.get("body", "")[:1500]

            h_text = (
                f"📨 *{i}-xat*\n"
                f"📌 {h_subject}\n"
                f"🕐 {h_date}\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"{h_body}"
            )
            h_key = f"hist_{uid_str}_{i}"
            if h_key not in pending:
                pending[h_key] = {
                    "sender":     em["sender"],
                    "subject":    h.get("subject", em["subject"]),
                    "body":       h.get("body", ""),
                    "history":    [],
                    "created_at": datetime.now().isoformat()
                }
                pending_changed = True  # FIX 6: o'zgarish borligini belgilaymiz

            h_keyboard = {
                "inline_keyboard": [[
                    {"text": f"✍️ {i}-xatga javob", "callback_data": f"reply:{h_key}"}
                ]]
            }
            send_telegram(h_text, chat_id=chat_id, reply_markup=h_keyboard)

        # FIX 6: barcha tarix kalitlari qo'shilgandan keyin faqat bir marta saqlaymiz
        if pending_changed:
            save_json(PENDING_FILE, pending)

    return pending



def fetch_sender_history(mail_conn, sender_email: str, limit: int = 5) -> list:
    history = []
    try:
        # FIX 6: IMAP search to'g'ri formatda — alohida argumentlar
        _, data = mail_conn.search(None, "FROM", f'"{sender_email}"')
        ids = data[0].split()
        recent_ids = ids[-limit-1:-1] if len(ids) > 1 else []
        for uid in reversed(recent_ids):
            _, msg_data = mail_conn.fetch(uid, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)
            subject = decode_str(msg.get("Subject", "(без темы)"))
            body    = get_email_body(msg)
            date    = msg.get("Date", "")
            history.append({"subject": subject, "body": body[:1000], "date": date})
    except Exception as e:
        log.warning(f"История писем: {e}")
    return history

def fetch_new_emails(seen_ids: set) -> list:
    new_emails = []
    mail = None
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(YANDEX_EMAIL, YANDEX_PASSWORD)
        mail.select("INBOX")
        _, data = mail.search(None, "UNSEEN")
        ids = data[0].split()
        for uid in ids:
            uid_str = uid.decode()
            if uid_str in seen_ids:
                continue
            _, msg_data = mail.fetch(uid, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)
            sender  = decode_str(msg.get("From", ""))
            subject = decode_str(msg.get("Subject", "(без темы)"))
            body    = get_email_body(msg)

            sender_addr = extract_email_address(sender)
            history = fetch_sender_history(mail, sender_addr)

            new_emails.append({
                "uid": uid_str,
                "sender": sender,
                "subject": subject,
                "body": body,
                "history": history
            })
            log.info(f"Новое письмо: {sender} — {subject} (история: {len(history)} писем)")
    except Exception as e:
        log.error(f"IMAP ошибка: {e}")
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass
    return new_emails

# ─── TELEGRAM UPDATE HANDLER ──────────────────────────────────────────────────

def get_telegram_updates(offset: int = 0) -> list:
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getUpdates"
    try:
        resp = requests.get(url, params={"offset": offset, "timeout": 5}, timeout=30)
        if resp.ok:
            return resp.json().get("result", [])
        log.warning(f"getUpdates muvaffaqiyatsiz: {resp.status_code}")
    except requests.exceptions.Timeout:
        log.warning("Telegram getUpdates timeout")
    except Exception as e:
        log.error(f"Telegram getUpdates ошибка: {e}")
    return []

def answer_callback_query(callback_query_id: str):
    try:
        resp = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/answerCallbackQuery",
            json={"callback_query_id": callback_query_id},
            timeout=5
        )
        if not resp.ok:
            log.warning(f"answerCallbackQuery xato: {resp.text}")
    except Exception as e:
        log.warning(f"answerCallbackQuery exception: {e}")

def handle_telegram_updates(pending: dict, last_update_id: int) -> tuple:
    updates = get_telegram_updates(last_update_id + 1)

    for update in updates:
        last_update_id = update["update_id"]

        # Tugma bosilishi
        if "callback_query" in update:
            cq = update["callback_query"]
            chat_id = str(cq["message"]["chat"]["id"])
            data = cq["data"]

            answer_callback_query(cq["id"])

            if chat_id not in TELEGRAM_CHAT_IDS:
                continue

            # ── INBOX SAHIFASI NAVIGATSIYA ────────────────────────────
            if data.startswith("inbox_page:"):
                parts = data.split(":")
                try:
                    days  = int(parts[1]) if len(parts) > 1 else 7
                    page  = int(parts[2]) if len(parts) > 2 else 0
                except ValueError:
                    days, page = 7, 0  # FIX 3: buzilgan data — default qiymatlar
                inbox_emails = fetch_inbox_emails(days=days)
                send_inbox_list(chat_id, inbox_emails, page=page, days=days)

            # ── XAT TO'LIQ KO'RISH ───────────────────────────────────
            elif data.startswith("view_email:"):
                parts = data.split(":")
                try:
                    uid_str   = parts[1] if len(parts) > 1 else ""
                    back_days = int(parts[2]) if len(parts) > 2 else 7
                    back_page = int(parts[3]) if len(parts) > 3 else 0
                except ValueError:
                    uid_str, back_days, back_page = "", 7, 0  # FIX 3
                if uid_str:
                    result = send_email_detail(
                        chat_id, uid_str,
                        back_days=back_days, back_page=back_page,
                        pending=pending
                    )
                    pending = result  # FIX 1: har doim pending qaytariladi

            # Javob yozish tugmasi
            elif data.startswith("reply:"):
                email_id = data.split(":", 1)[1]
                if email_id in pending:
                    # FIX 2: Avval boshqa adminning sessiyasini tozalaymiz
                    pending[email_id]["waiting_reply"] = chat_id
                    # FIX 1: waiting_edit ni ham tozalaymiz — konflikt yo'q
                    pending[email_id].pop("waiting_edit", None)
                    pending[email_id].pop("preview_text", None)
                    save_json(PENDING_FILE, pending)
                    em = pending[email_id]
                    send_telegram(
                        f"✍️ *Напишите тезисы для ответа:*\n\n"
                        f"Клиент: `{escape_markdown(em['sender'])}`\n"
                        f"Тема: {escape_markdown(em['subject'])}\n\n"
                        f"_Напишите кратко — GPT составит полное письмо_",
                        chat_id=chat_id
                    )

            elif data.startswith("send:"):
                email_id = data.split(":", 1)[1]
                if email_id in pending:
                    em = pending[email_id]
                    # FIX 3: preview_text mavjud VA bo'sh emasligini tekshiramiz
                    preview_text = em.get("preview_text", "").strip()
                    if not preview_text:
                        send_telegram(
                            "⚠️ *Yuborish uchun matn yo'q. Avval javob yozing.*",
                            chat_id=chat_id
                        )
                        continue

                    russian_text = preview_text
                    try:
                        to_addr = extract_email_address(em["sender"])
                        send_telegram(f"⏳ *Перевожу и отправляю...*", chat_id=chat_id)
                        final_text = translate_reply(russian_text, em["body"])
                        send_email_reply(to_addr, em["subject"], final_text)
                        send_telegram(
                            f"✅ *Письмо отправлено!*\n📬 *Кому:* `{escape_markdown(to_addr)}`",
                            chat_id=chat_id
                        )
                        log.info(f"Письмо отправлено: {to_addr}")
                        del pending[email_id]
                        save_json(PENDING_FILE, pending)
                    except Exception as e:
                        log.error(f"Ошибка отправки: {e}")
                        send_telegram(f"❌ *Xat yuborishda xato yuz berdi. Qayta urinib ko'ring.*", chat_id=chat_id)

            # Qayta yozish tugmasi
            elif data.startswith("rewrite:"):
                email_id = data.split(":", 1)[1]
                if email_id in pending:
                    pending[email_id]["waiting_reply"] = chat_id
                    # FIX 1: waiting_edit ni ham tozalaymiz
                    pending[email_id].pop("waiting_edit", None)
                    pending[email_id].pop("preview_text", None)
                    save_json(PENDING_FILE, pending)
                    send_telegram(
                        f"🔄 *Напишите новые тезисы:*\n\n"
                        f"_GPT составит новый вариант письма_",
                        chat_id=chat_id
                    )

            # Tahrirlash tugmasi
            elif data.startswith("edit:"):
                email_id = data.split(":", 1)[1]
                if email_id in pending:
                    pending[email_id]["waiting_edit"] = chat_id
                    # FIX 1: waiting_reply ni ham tozalaymiz
                    pending[email_id].pop("waiting_reply", None)
                    save_json(PENDING_FILE, pending)
                    em = pending[email_id]
                    current_text = em.get("preview_text", "")
                    send_telegram(
                        f"✏️ *Отредактируйте письмо:*\n\n"
                        f"Скопируйте текст ниже, внесите изменения и отправьте:\n\n"
                        f"━━━━━━━━━━━━━━━━━━━━\n"
                        f"{current_text[:3000]}\n"
                        f"━━━━━━━━━━━━━━━━━━━━",
                        chat_id=chat_id
                    )

        # Matn xabari
        elif "message" in update:
            msg = update["message"]
            chat_id = str(msg["chat"]["id"])
            text = msg.get("text", "").strip()

            if not text:
                continue
            if chat_id not in TELEGRAM_CHAT_IDS:
                continue

            # /inbox buyrug'i — inbox ro'yxatini ko'rsatish
            if text.lower() in ("/inbox", "/start"):
                # FIX 7: "yuklanmoqda" xabari keraksiz — faqat sana tugmalarini ko'rsatamiz
                send_telegram(
                    "📬 *Qaysi davr uchun inbox ko'rsatilsin?*",
                    chat_id=chat_id,
                    reply_markup={"inline_keyboard": [[
                        {"text": "📅 Oxirgi 7 kun",  "callback_data": "inbox_page:7:0"},
                        {"text": "📅 Oxirgi 30 kun", "callback_data": "inbox_page:30:0"},
                    ]]}
                )
                continue

            # Boshqa / buyruqlarni e'tiborsiz qoldiramiz
            if text.startswith("/"):
                continue

            # Javob kutayotgan xat bormi?
            waiting_email_id = None
            waiting_mode = None
            for eid, em in pending.items():
                # FIX 2: faqat shu chat_id ga tegishli kutuvlarni topamiz
                if em.get("waiting_reply") == chat_id:
                    waiting_email_id = eid
                    waiting_mode = "reply"
                    break
                elif em.get("waiting_edit") == chat_id:
                    waiting_email_id = eid
                    waiting_mode = "edit"
                    break

            if waiting_email_id:
                em = pending[waiting_email_id]

                # Tahrirlash rejimi
                if waiting_mode == "edit":
                    pending[waiting_email_id]["preview_text"] = text
                    pending[waiting_email_id].pop("waiting_edit", None)
                    save_json(PENDING_FILE, pending)

                    to_addr = extract_email_address(em["sender"])
                    preview_msg = (
                        f"📝 *Ваш текст — проверьте:*\n\n"
                        f"📬 *Кому:* `{escape_markdown(to_addr)}`\n"
                        f"━━━━━━━━━━━━━━━━━━━━\n\n"
                        f"{text[:3000]}\n\n"
                        f"━━━━━━━━━━━━━━━━━━━━\n"
                        f"_Подпись AutoZIP будет добавлена автоматически_"
                    )
                    reply_markup = {
                        "inline_keyboard": [[
                            {"text": "✅ Отправить", "callback_data": f"send:{waiting_email_id}"},
                            {"text": "✏️ Редактировать", "callback_data": f"edit:{waiting_email_id}"},
                            {"text": "🔄 Переписать", "callback_data": f"rewrite:{waiting_email_id}"}
                        ]]
                    }
                    send_telegram(preview_msg, chat_id=chat_id, reply_markup=reply_markup)

                # Tezislar rejimi — GPT yozadi
                elif waiting_mode == "reply":
                    send_telegram(f"⏳ *Составляю письмо...*", chat_id=chat_id)
                    try:
                        reply_text = generate_reply(em["body"], em["subject"], text, em.get("history", []))

                        pending[waiting_email_id]["preview_text"] = reply_text
                        pending[waiting_email_id].pop("waiting_reply", None)
                        save_json(PENDING_FILE, pending)

                        to_addr = extract_email_address(em["sender"])
                        preview_msg = (
                            f"📝 *Готовое письмо — проверьте:*\n\n"
                            f"📬 *Кому:* `{escape_markdown(to_addr)}`\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n\n"
                            f"{reply_text[:3000]}\n\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"_Подпись AutoZIP будет добавлена автоматически_"
                        )
                        reply_markup = {
                            "inline_keyboard": [[
                                {"text": "✅ Отправить", "callback_data": f"send:{waiting_email_id}"},
                                {"text": "✏️ Редактировать", "callback_data": f"edit:{waiting_email_id}"},
                                {"text": "🔄 Переписать", "callback_data": f"rewrite:{waiting_email_id}"}
                            ]]
                        }
                        send_telegram(preview_msg, chat_id=chat_id, reply_markup=reply_markup)

                    except Exception as e:
                        log.error(f"Ошибка генерации: {e}")
                        send_telegram(f"❌ *GPT javob yozishda xato. Qayta urinib ko'ring.*", chat_id=chat_id)
                        pending[waiting_email_id].pop("waiting_reply", None)
                        save_json(PENDING_FILE, pending)

    # FIX 4: last_update_id ni faylga saqlaymiz
    if updates:
        save_json(LAST_UPDATE_FILE, {"last_update_id": last_update_id})

    return pending, last_update_id

# ─── ASOSIY TSIKL ─────────────────────────────────────────────────────────────

def main():
    if not TELEGRAM_CHAT_IDS:
        log.error("TELEGRAM_CHAT_IDS yoki TELEGRAM_CHAT_ID .env da topilmadi!")
        sys.exit(1)

    log.info("Email Agent запущен")
    log.info(f"Telegram foydalanuvchilar: {len(TELEGRAM_CHAT_IDS)} ta")
    send_telegram("🤖 *Email-агент запущен!*\nНовые письма от клиентов будут автоматически анализироваться.")

    # Ishga tushganda inbox ko'rsatish — sana tanlash tugmalari
    send_telegram(
        "📬 *Inbox qaysi davr uchun ko'rsatilsin?*",
        reply_markup={"inline_keyboard": [[
            {"text": "📅 Oxirgi 7 kun",  "callback_data": "inbox_page:7:0"},
            {"text": "📅 Oxirgi 30 kun", "callback_data": "inbox_page:30:0"},
        ]]}
    )

    seen_ids = set(load_json(SEEN_IDS_FILE, []))
    pending  = load_json(PENDING_FILE, {})

    # FIX 4: last_update_id ni fayldan o'qiymiz — agent qayta ishlaganda eski updatelar qayta ishlanmaydi
    saved = load_json(LAST_UPDATE_FILE, {})
    last_update_id = saved.get("last_update_id", 0)

    # FIX 9 (loop_count): tsikl hisoblagichi — last_update_id o'rniga ishlatiladi
    loop_count = 0

    while True:
        try:
            pending, last_update_id = handle_telegram_updates(pending, last_update_id)

            log.info("Проверка почты...")
            new_emails = fetch_new_emails(seen_ids)

            for em in new_emails:
                log.info(f"Анализируется: {em['sender']}")

                email_id = em["uid"]

                # FIX 5: seen_ids ga AVVAL qo'shamiz — xabar yuborishda xato bo'lsa ikki marta ishlanmaydi
                seen_ids.add(email_id)
                seen_ids = trim_seen_ids(seen_ids)
                save_json(SEEN_IDS_FILE, list(seen_ids))

                analysis = analyze_email(em["sender"], em["subject"], em["body"], em.get("history", []))
                message  = format_telegram_message(em["sender"], em["subject"], analysis)

                pending[email_id] = {
                    "sender":     em["sender"],
                    "subject":    em["subject"],
                    "body":       em["body"],
                    "history":    em.get("history", []),
                    "created_at": datetime.now().isoformat()
                }
                save_json(PENDING_FILE, pending)

                reply_markup = {
                    "inline_keyboard": [[
                        {"text": "✍️ Ответить клиенту", "callback_data": f"reply:{email_id}"}
                    ]]
                }

                send_telegram(message, reply_markup=reply_markup)
                log.info(f"Отправлено: {em['sender']}")
                time.sleep(2)

            # FIX 9: loop_count bilan 100 ta tsiklda tozalash — last_update_id emas
            loop_count += 1
            if loop_count % 100 == 0:
                pending = cleanup_pending(pending)
                save_json(PENDING_FILE, pending)

        except Exception as e:
            log.error(f"Ошибка основного цикла: {e}")
            send_telegram(f"⚠️ *Agent xatosi yuz berdi. Log faylini tekshiring.*")

        time.sleep(10)


if __name__ == "__main__":
    main()
