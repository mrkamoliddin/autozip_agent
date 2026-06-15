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
from datetime import datetime, timedelta, timezone
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
LAST_UPDATE_FILE    = "last_update_id.json"
MENU_MESSAGE_FILE   = "menu_message_ids.json"  # Menyu xabar ID lari
IMAP_HOST           = "imap.yandex.ru"
IMAP_PORT           = 993
SMTP_HOST           = "smtp.yandex.ru"
SMTP_PORT           = 465

SEEN_IDS_MAX        = 1000
PENDING_MAX_DAYS    = 30
OPENAI_MAX_RETRIES  = 5
OPENAI_RETRY_DELAY  = 3
INBOX_PAGE_SIZE     = 3    # Количество писем на странице

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

def edit_telegram_message(chat_id: str, message_id: int, text: str, reply_markup: dict = None):
    """Mavjud Telegram xabarni tahrirlash (badge yangilash uchun)."""
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/editMessageText"
    payload = {
        "chat_id":    chat_id,
        "message_id": message_id,
        "text":       text,
        "parse_mode": "Markdown"
    }
    if reply_markup:
        payload["reply_markup"] = reply_markup
    try:
        resp = requests.post(url, json=payload, timeout=10)
        if not resp.ok:
            # Xabar o'zgarmaganda Telegram 400 qaytaradi — bu normal
            if "message is not modified" not in resp.text:
                log.warning(f"editMessage xato: {resp.text}")
    except Exception as e:
        log.warning(f"editMessage exception: {e}")

def fetch_sent_emails(limit: int = 20) -> list:
    """Sent papkasidan yuborilgan xatlarni olish."""
    sent_emails = []
    mail = None
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(YANDEX_EMAIL, YANDEX_PASSWORD)

        # Barcha papkalarni ko'ramiz
        _, folders = mail.list()
        log.info(f"IMAP papkalar: {[f.decode('utf-8', errors='replace') if isinstance(f, bytes) else f for f in folders]}")

        selected = False
        for f in folders:
            f_str = f.decode("utf-8", errors="replace") if isinstance(f, bytes) else f
            # Sent papkasini topamiz
            if any(name.lower() in f_str.lower() for name in ["sent", "отправлен"]):
                # Papka nomini oxirgi qismidan ajratamiz
                # Format: (\HasNoChildren) "|" "Sent"
                if '"|"' in f_str or '" "' in f_str:
                    folder_name = f_str.split('"')[-2]
                else:
                    folder_name = f_str.split()[-1].strip('"')
                log.info(f"Sent papkasi topildi: {folder_name}")
                result, _ = mail.select(f'"{folder_name}"')
                if result == "OK":
                    selected = True
                    break

        if not selected:
            log.error("Sent papkasi topilmadi")
            return []

        _, data = mail.search(None, "ALL")
        ids = data[0].split()
        log.info(f"Sent papkasida {len(ids)} ta xat")
        recent_ids = ids[-limit:] if len(ids) >= limit else ids

        for uid in reversed(recent_ids):
            _, msg_data = mail.fetch(uid, "(FLAGS BODY.PEEK[HEADER.FIELDS (TO SUBJECT DATE)])")
            raw_headers = msg_data[0][1]
            msg = email.message_from_bytes(raw_headers)

            to_addr  = decode_str(msg.get("To", ""))
            subject  = decode_str(msg.get("Subject", "(без темы)"))
            date_str = msg.get("Date", "")

            date_obj = None
            try:
                date_obj = parsedate_to_datetime(date_str)
                if date_obj.tzinfo is None:
                    date_obj = date_obj.replace(tzinfo=timezone.utc)
            except Exception:
                date_obj = datetime.now(tz=timezone.utc)

            sent_emails.append({
                "to":       to_addr,
                "subject":  subject,
                "date":     date_str,
                "date_obj": date_obj,
            })

    except Exception as e:
        log.error(f"fetch_sent_emails xato: {e}")
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass
    return sent_emails

def build_menu_text(counts: dict = None, new_badge: bool = False) -> tuple:
    """Asosiy 4 bo'limli menyu — har bo'lim soni bilan."""
    if counts is None:
        counts = {}
    badge = " 🔴" if new_badge else ""
    unread = counts.get("unread", 0)
    inbox  = counts.get("inbox", 0)
    sent   = counts.get("sent", 0)

    unread_label = f"🔴 Непрочитанные сообщения ({unread}){badge}" if unread > 0 else f"🔴 Непрочитанные сообщения{badge}"
    inbox_label  = f"📥 Входящие сообщения ({inbox})"  if inbox  > 0 else "📥 Входящие сообщения"
    sent_label   = f"📤 Отправленные сообщения ({sent})" if sent  > 0 else "📤 Отправленные сообщения"

    text = "🏠 *AutoZIP Email Agent*\n━━━━━━━━━━━━━━━━━━━━"
    keyboard = {
        "inline_keyboard": [
            [{"text": unread_label, "callback_data": "menu:new"}],
            [{"text": inbox_label,  "callback_data": "menu:inbox"}],
            [{"text": sent_label,   "callback_data": "menu:sent"}],
            [{"text": "🔍 Поиск нового клиента", "callback_data": "menu:search"}],
        ]
    }
    return text, keyboard

def send_main_menu(chat_id: str, counts: dict = None) -> int | None:
    """Asosiy menyuni yuboradi va message_id ni qaytaradi."""
    text, keyboard = build_menu_text(counts)
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id":      chat_id,
        "text":         text,
        "parse_mode":   "Markdown",
        "reply_markup": keyboard
    }
    try:
        resp = requests.post(url, json=payload, timeout=10)
        if resp.ok:
            return resp.json()["result"]["message_id"]
        else:
            log.error(f"send_main_menu xato: {resp.text}")
    except Exception as e:
        log.error(f"send_main_menu exception: {e}")
    return None

def update_menu_badge(menu_msg_ids: dict, counts: dict, new_badge: bool = False):
    """Har bir admindagi menyu xabarini yangilaydi."""
    text, keyboard = build_menu_text(counts, new_badge=new_badge)
    for cid, mid in menu_msg_ids.items():
        if isinstance(mid, int):
            edit_telegram_message(cid, mid, text, reply_markup=keyboard)

def fetch_all_counts() -> dict:
    """Bitta IMAP ulanishda barcha sonlarni oladi."""
    mail = None
    counts = {"unread": 0, "inbox": 0, "sent": 0}
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(YANDEX_EMAIL, YANDEX_PASSWORD)

        # INBOX — o'qilmagan va jami
        mail.select("INBOX")
        _, unseen_data = mail.search(None, "UNSEEN")
        _, all_data    = mail.search(None, "ALL")

        # Bo'sh b'' ni 0 ga aylantirish
        unseen_ids = [x for x in unseen_data[0].split() if x]
        all_ids    = [x for x in all_data[0].split()    if x]
        counts["unread"] = len(unseen_ids)
        counts["inbox"]  = len(all_ids)

        # Sent papkasini topamiz
        _, folders = mail.list()
        sent_found = False
        for f in folders:
            f_str = f.decode("utf-8", errors="replace") if isinstance(f, bytes) else f
            if any(name.lower() in f_str.lower() for name in ["sent", "отправлен"]):
                folder_name = f_str.split('"')[-2] if '"' in f_str else f_str.split()[-1].strip('"')
                result, _ = mail.select(f'"{folder_name}"')
                if result == "OK":
                    _, sent_data = mail.search(None, "ALL")
                    sent_ids = [x for x in sent_data[0].split() if x]
                    counts["sent"] = len(sent_ids)
                    sent_found = True
                    break

        if not sent_found:
            log.warning("Sent papkasi topilmadi")

        log.info(f"Counts: {counts}")

    except Exception as e:
        log.warning(f"fetch_all_counts xato: {e}")
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass
    return counts

def count_unread_inbox() -> int:
    """IMAP dan faqat o'qilmagan xatlar sonini olish."""
    mail = None
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(YANDEX_EMAIL, YANDEX_PASSWORD)
        mail.select("INBOX")
        _, data = mail.search(None, "UNSEEN")
        return len(data[0].split()) if data[0] else 0
    except Exception as e:
        log.warning(f"count_unread_inbox xato: {e}")
        return 0
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass





def _fetch_inbox_by_criteria(criteria: str, limit: int = 50) -> list:
    """IMAP search criteria bo'yicha xatlarni oluvchi ichki funksiya."""
    emails = []
    mail = None
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(YANDEX_EMAIL, YANDEX_PASSWORD)
        mail.select("INBOX")
        _, data = mail.search(None, criteria)
        ids = data[0].split()
        recent_ids = ids[-limit:] if len(ids) > limit else ids

        for uid in reversed(recent_ids):
            uid_str = uid.decode()
            _, msg_data = mail.fetch(uid, "(FLAGS BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
            raw_headers = msg_data[0][1]
            msg = email.message_from_bytes(raw_headers)

            sender   = decode_str(msg.get("From", ""))
            subject  = decode_str(msg.get("Subject", "(без темы)"))
            date_str = msg.get("Date", "")

            flags_raw = msg_data[0][0].decode() if isinstance(msg_data[0][0], bytes) else str(msg_data[0][0])
            seen = "\\Seen" in flags_raw

            date_obj = None
            try:
                date_obj = parsedate_to_datetime(date_str)
                if date_obj.tzinfo is None:
                    date_obj = date_obj.replace(tzinfo=timezone.utc)
            except Exception:
                date_obj = datetime.now(tz=timezone.utc)

            emails.append({
                "uid":      uid_str,
                "sender":   sender,
                "subject":  subject,
                "date":     date_str,
                "date_obj": date_obj,
                "seen":     seen,
            })
    except Exception as e:
        log.error(f"_fetch_inbox_by_criteria ({criteria}) xato: {e}")
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass
    emails.sort(key=lambda x: x["date_obj"] or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
    return emails

def fetch_new_inbox_emails() -> list:
    """
    'Непрочитанные' bo'limi: o'qilmagan + bugun kelgan xatlar.
    Bitta IMAP ulanishda bajaradi.
    """
    emails = []
    mail = None
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(YANDEX_EMAIL, YANDEX_PASSWORD)
        mail.select("INBOX")

        today = datetime.now().strftime("%d-%b-%Y")

        # O'qilmagan xatlar
        _, unseen_data = mail.search(None, "UNSEEN")
        unseen_ids = set(unseen_data[0].split())

        # Bugun kelgan o'qilgan xatlar
        _, today_data = mail.search(None, f"SEEN SINCE {today}")
        today_ids = today_data[0].split()

        # Birlashtirish — unseen birinchi
        all_ids = list(unseen_ids) + [uid for uid in today_ids if uid not in unseen_ids]

        for uid in all_ids:
            uid_str = uid.decode() if isinstance(uid, bytes) else uid
            _, msg_data = mail.fetch(uid if isinstance(uid, bytes) else uid.encode(),
                                     "(FLAGS BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
            raw_headers = msg_data[0][1]
            msg = email.message_from_bytes(raw_headers)

            sender   = decode_str(msg.get("From", ""))
            subject  = decode_str(msg.get("Subject", "(без темы)"))
            date_str = msg.get("Date", "")
            flags_raw = msg_data[0][0].decode() if isinstance(msg_data[0][0], bytes) else str(msg_data[0][0])
            seen = "\\Seen" in flags_raw

            date_obj = None
            try:
                date_obj = parsedate_to_datetime(date_str)
                if date_obj.tzinfo is None:
                    date_obj = date_obj.replace(tzinfo=timezone.utc)
            except Exception:
                date_obj = datetime.now(tz=timezone.utc)

            emails.append({
                "uid":      uid_str,
                "sender":   sender,
                "subject":  subject,
                "date":     date_str,
                "date_obj": date_obj,
                "seen":     seen,
            })

    except Exception as e:
        log.error(f"fetch_new_inbox_emails xato: {e}")
    finally:
        if mail:
            try:
                mail.logout()
            except Exception:
                pass

    emails.sort(key=lambda x: x["date_obj"] or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
    return emails

def fetch_read_inbox_emails(days: int = 30) -> list:
    """'Входящие' bo'limi: faqat o'qilgan (SEEN) xatlar."""
    since_date = (datetime.now() - timedelta(days=days + 1)).strftime("%d-%b-%Y")
    return _fetch_inbox_by_criteria(f"SEEN SINCE {since_date}")



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


def send_email_list(chat_id: str, emails: list, page: int = 0,
                    section: str = "inbox", title: str = "📬 *ВХОДЯЩИЕ*"):
    """
    Список писем с пагинацией (3 на страницу).
    Каждое письмо показывается отдельным блоком с кнопкой Открыть.
    """
    total = len(emails)
    if total == 0:
        send_telegram(
            "📭 *Писем не найдено.*",
            chat_id=chat_id,
            reply_markup={"inline_keyboard": [[
                {"text": "🏠 Главное меню", "callback_data": "menu:main"}
            ]]}
        )
        return

    start       = page * INBOX_PAGE_SIZE
    end         = min(start + INBOX_PAGE_SIZE, total)
    page_emails = emails[start:end]
    total_pages = (total + INBOX_PAGE_SIZE - 1) // INBOX_PAGE_SIZE

    # Sarlavha
    send_telegram(
        f"{title}\nВсего: *{total}*  |  Стр. {page+1}/{total_pages}\n━━━━━━━━━━━━━━━━━━━━",
        chat_id=chat_id
    )

    # Har bir xat alohida xabar + Открыть tugmasi
    for i, em in enumerate(page_emails, start=start+1):
        icon = "🔴" if not em.get("seen", True) else "⚪"
        try:
            d = em["date_obj"].strftime("%d.%m %H:%M") if em.get("date_obj") else "—"
        except Exception:
            d = "—"
        sender_short  = escape_markdown(extract_email_address(em.get("sender", em.get("to", ""))))
        subject_short = escape_markdown(em.get("subject", "")[:45])

        text = f"{icon} *{i}.* {subject_short}\n👤 `{sender_short}` · {d}"
        cb   = safe_callback(f"view_email:{em['uid']}:{section}:{page}")

        is_last = (i == start + len(page_emails))
        nav_row = []
        if page > 0:
            nav_row.append({"text": "« Назад", "callback_data": f"list_page:{section}:{page-1}"})
        if end < total:
            nav_row.append({"text": "Вперёд »", "callback_data": f"list_page:{section}:{page+1}"})

        keyboard_rows = [[{"text": "📂 Открыть", "callback_data": cb}]]
        if is_last:
            if nav_row:
                keyboard_rows.append(nav_row)
            keyboard_rows.append([{"text": "🏠 Главное меню", "callback_data": "menu:main"}])

        send_telegram(text, chat_id=chat_id, reply_markup={"inline_keyboard": keyboard_rows})
        time.sleep(0.3)



def send_email_detail(chat_id: str, uid_str: str, back_section: str = "new",
                      back_page: int = 0, pending: dict = None):
    """
    Показать полное письмо: история + кнопка ответа для каждого.
    Всегда возвращает pending dict.
    """
    if pending is None:
        pending = {}

    send_telegram("⏳ *Загружаю письмо...*", chat_id=chat_id)
    em = fetch_email_full(uid_str)
    if not em:
        send_telegram(
            "❌ *Ошибка загрузки письма.*",
            chat_id=chat_id,
            reply_markup={"inline_keyboard": [[
                {"text": "◀ Назад", "callback_data": f"list_page:{back_section}:{back_page}"},
                {"text": "🏠 Главное меню", "callback_data": "menu:main"},
            ]]}
        )
        return pending

    sender_safe  = escape_markdown(em["sender"])
    subject_safe = escape_markdown(em["subject"])
    body_preview = em["body"][:2000] if em["body"] else "_(пусто)_"

    msg_text = (
        f"📧 *ДЕТАЛИ ПИСЬМА*\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 *От кого:* `{sender_safe}`\n"
        f"📌 *Тема:* {subject_safe}\n"
        f"🕐 *Дата:* {escape_markdown(em['date'][:30])}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        f"{body_preview}"
    )

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
            [{"text": "✍️ Ответить", "callback_data": f"reply:{uid_str}"}],
            [
                {"text": "◀ Назад",         "callback_data": f"list_page:{back_section}:{back_page}"},
                {"text": "🏠 Главное меню", "callback_data": "menu:main"},
            ]
        ]
    }
    send_telegram(msg_text, chat_id=chat_id, reply_markup=main_keyboard)

    # Предыдущие письма этого клиента
    history = em.get("history", [])
    if history:
        send_telegram(
            f"🔄 *Предыдущие письма этого клиента ({len(history)} шт.):*",
            chat_id=chat_id
        )
        pending_changed = False
        for i, h in enumerate(history, 1):
            h_subject = escape_markdown(h.get("subject", ""))
            h_date    = escape_markdown(str(h.get("date", ""))[:30])
            h_body    = h.get("body", "")[:1500]

            h_text = (
                f"📨 *Письмо {i}*\n"
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
                pending_changed = True

            h_keyboard = {
                "inline_keyboard": [[
                    {"text": f"✍️ Ответить на письмо {i}", "callback_data": f"reply:{h_key}"}
                ]]
            }
            send_telegram(h_text, chat_id=chat_id, reply_markup=h_keyboard)

        if pending_changed:
            save_json(PENDING_FILE, pending)

    # Eng pastda — asosiy javob tugmasi
    send_telegram(
        f"✉️ *Ответить на это письмо:*",
        chat_id=chat_id,
        reply_markup={
            "inline_keyboard": [
                [{"text": "✍️ Написать ответ", "callback_data": f"reply:{uid_str}"}],
                [
                    {"text": "◀ Назад",         "callback_data": f"list_page:{back_section}:{back_page}"},
                    {"text": "🏠 Главное меню", "callback_data": "menu:main"},
                ]
            ]
        }
    )

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

def handle_telegram_updates(pending: dict, last_update_id: int, menu_msg_ids: dict) -> tuple:
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

            # ── ASOSIY MENYU TUGMALARI ────────────────────────────────
            if data.startswith("menu:"):
                section = data.split(":", 1)[1]

                if section == "main":
                    cached = load_json(MENU_MESSAGE_FILE, {})
                    counts = cached.get("last_counts", {"unread": 0, "inbox": 0, "sent": 0})
                    mid = send_main_menu(chat_id, counts=counts)
                    if mid:
                        menu_msg_ids[chat_id] = mid
                        cached[chat_id] = mid
                        save_json(MENU_MESSAGE_FILE, cached)

                elif section == "new":
                    send_telegram("⏳ *Загружаю новые сообщения...*", chat_id=chat_id)
                    new_emails = fetch_new_inbox_emails()

                    if not new_emails:
                        send_telegram(
                            "📭 *Новых сообщений нет.*",
                            chat_id=chat_id,
                            reply_markup={"inline_keyboard": [[
                                {"text": "🏠 Главное меню", "callback_data": "menu:main"}
                            ]]}
                        )
                    else:
                        # Sarlavha xabari
                        send_telegram(
                            f"🔴 *НЕПРОЧИТАННЫЕ СООБЩЕНИЯ ({len(new_emails)})*\n━━━━━━━━━━━━━━━━━━━━",
                            chat_id=chat_id
                        )
                        # Har bir xat uchun alohida xabar + Открыть tugmasi
                        for i, em in enumerate(new_emails, 1):
                            icon = "🔴" if not em.get("seen") else "⚪"
                            try:
                                d = em["date_obj"].strftime("%d.%m %H:%M")
                            except Exception:
                                d = "—"
                            sender_short  = escape_markdown(extract_email_address(em["sender"]))
                            subject_short = escape_markdown(em["subject"][:45])

                            text = (
                                f"{icon} *{i}.* {subject_short}\n"
                                f"👤 `{sender_short}` · {d}"
                            )
                            cb = safe_callback(f"open_new:{em['uid']}")
                            is_last = (i == len(new_emails))
                            keyboard = {
                                "inline_keyboard": [
                                    [{"text": "📂 Открыть", "callback_data": cb}],
                                ] + ([[{"text": "🏠 Главное меню", "callback_data": "menu:main"}]] if is_last else [])
                            }
                            send_telegram(text, chat_id=chat_id, reply_markup=keyboard)
                            time.sleep(0.3)

                elif section == "inbox":
                    send_telegram("⏳ *Загружаю входящие...*", chat_id=chat_id)
                    emails = fetch_read_inbox_emails(days=30)
                    send_email_list(chat_id, emails, page=0, section="inbox",
                                   title="📥 *ВХОДЯЩИЕ СООБЩЕНИЯ* (прочитанные)")

                elif section == "sent":
                    send_telegram("⏳ *Загружаю отправленные...*", chat_id=chat_id)
                    sent = fetch_sent_emails(limit=50)
                    if not sent:
                        send_telegram(
                            "📭 *Отправленных писем не найдено.*",
                            chat_id=chat_id,
                            reply_markup={"inline_keyboard": [[
                                {"text": "🏠 Главное меню", "callback_data": "menu:main"}
                            ]]}
                        )
                    else:
                        total_sent  = len(sent)
                        total_pages = (total_sent + INBOX_PAGE_SIZE - 1) // INBOX_PAGE_SIZE
                        page_sent   = sent[:INBOX_PAGE_SIZE]

                        send_telegram(
                            f"📤 *ОТПРАВЛЕННЫЕ СООБЩЕНИЯ*\nВсего: *{total_sent}*  |  Стр. 1/{total_pages}\n━━━━━━━━━━━━━━━━━━━━",
                            chat_id=chat_id
                        )
                        for i, s in enumerate(page_sent, 1):
                            try:
                                d = s["date_obj"].strftime("%d.%m %H:%M")
                            except Exception:
                                d = "—"
                            to_name = escape_markdown(s["to"][:45])
                            subj    = escape_markdown(s["subject"][:45])
                            text    = f"📨 *{i}.* {subj}\n👤 `{to_name}` · {d}"

                            is_last = (i == len(page_sent))
                            keyboard_rows = [[{"text": "📂 Открыть", "callback_data": safe_callback(f"view_sent:{i-1}:0")}]]
                            if is_last:
                                if total_sent > INBOX_PAGE_SIZE:
                                    keyboard_rows.append([{"text": "Вперёд »", "callback_data": "sent_page:1"}])
                                keyboard_rows.append([{"text": "🏠 Главное меню", "callback_data": "menu:main"}])

                            send_telegram(text, chat_id=chat_id, reply_markup={"inline_keyboard": keyboard_rows})
                            time.sleep(0.3)

                elif section == "search":
                    send_telegram(
                        "🔍 *Поиск нового клиента*\n\n_Этот раздел будет добавлен в ближайшее время._",
                        chat_id=chat_id,
                        reply_markup={"inline_keyboard": [[
                            {"text": "🏠 Главное меню", "callback_data": "menu:main"}
                        ]]}
                    )

            # ── RO'YXAT SAHIFASI NAVIGATSIYA ─────────────────────────
            elif data.startswith("list_page:"):
                parts = data.split(":")
                try:
                    sec  = parts[1] if len(parts) > 1 else "new"
                    page = int(parts[2]) if len(parts) > 2 else 0
                except ValueError:
                    sec, page = "new", 0
                if sec == "new":
                    emails = fetch_new_inbox_emails()
                    send_email_list(chat_id, emails, page=page, section="new",
                                   title="🔴 *НЕПРОЧИТАННЫЕ СООБЩЕНИЯ*")
                elif sec == "inbox":
                    emails = fetch_read_inbox_emails(days=30)
                    send_email_list(chat_id, emails, page=page, section="inbox",
                                   title="📥 *ВХОДЯЩИЕ СООБЩЕНИЯ* (прочитанные)")

            # ── XAT TO'LIQ KO'RISH ───────────────────────────────────
            elif data.startswith("view_email:"):
                parts = data.split(":")
                try:
                    uid_str      = parts[1] if len(parts) > 1 else ""
                    back_section = parts[2] if len(parts) > 2 else "new"
                    back_page    = int(parts[3]) if len(parts) > 3 else 0
                except ValueError:
                    uid_str, back_section, back_page = "", "new", 0
                if uid_str:
                    pending = send_email_detail(
                        chat_id, uid_str,
                        back_section=back_section,
                        back_page=back_page,
                        pending=pending
                    )

            # ── ESKI inbox_page KOLBEK (muvofiqlashuv) ───────────────
            elif data.startswith("inbox_page:"):
                parts = data.split(":")
                try:
                    page = int(parts[2]) if len(parts) > 2 else 0
                except ValueError:
                    page = 0
                emails = fetch_new_inbox_emails()
                send_email_list(chat_id, emails, page=page, section="new",
                               title="🔴 *НЕПРОЧИТАННЫЕ СООБЩЕНИЯ*")

            # ── YANGI XATNI OCHISH — GPT TAHLIL ─────────────────────
            elif data.startswith("open_new:"):
                email_id = data.split(":", 1)[1]
                send_telegram("⏳ *Анализирую письмо...*", chat_id=chat_id)

                # To'liq xatni olamiz
                if email_id not in pending:
                    full = fetch_email_full(email_id)
                    if not full:
                        send_telegram(
                            "❌ *Не удалось загрузить письмо.*",
                            chat_id=chat_id,
                            reply_markup={"inline_keyboard": [[
                                {"text": "◀ Назад", "callback_data": "menu:new"},
                                {"text": "🏠 Главное меню", "callback_data": "menu:main"},
                            ]]}
                        )
                        continue
                    pending[email_id] = {
                        "sender":     full["sender"],
                        "subject":    full["subject"],
                        "body":       full["body"],
                        "history":    full.get("history", []),
                        "created_at": datetime.now().isoformat()
                    }
                    save_json(PENDING_FILE, pending)
                    em_data = full
                else:
                    em_data = pending[email_id]

                # GPT tahlil
                try:
                    analysis = analyze_email(
                        em_data["sender"],
                        em_data["subject"],
                        em_data["body"],
                        em_data.get("history", [])
                    )
                    message = format_telegram_message(
                        em_data["sender"],
                        em_data["subject"],
                        analysis
                    )
                    reply_markup = {
                        "inline_keyboard": [
                            [{"text": "✍️ Ответить клиенту", "callback_data": f"reply:{email_id}"}],
                            [
                                {"text": "◀ Назад",         "callback_data": "menu:new"},
                                {"text": "🏠 Главное меню", "callback_data": "menu:main"},
                            ],
                        ]
                    }
                    send_telegram(message, chat_id=chat_id, reply_markup=reply_markup)
                except Exception as e:
                    log.error(f"Ошибка анализа {email_id}: {e}")
                    send_telegram(
                        "❌ *Ошибка анализа. Попробуйте ещё раз.*",
                        chat_id=chat_id,
                        reply_markup={"inline_keyboard": [[
                            {"text": "◀ Назад", "callback_data": "menu:new"},
                            {"text": "🏠 Главное меню", "callback_data": "menu:main"},
                        ]]}
                    )

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

            # /start va /menu — asosiy menyuni ko'rsatish
            if text.lower() in ("/start", "/menu"):
                counts = fetch_all_counts()
                mid = send_main_menu(chat_id, counts=counts)
                if mid:
                    cached = load_json(MENU_MESSAGE_FILE, {})
                    cached[chat_id] = mid
                    cached["last_counts"] = counts
                    save_json(MENU_MESSAGE_FILE, cached)
                    menu_msg_ids[chat_id] = mid
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

    return pending, last_update_id, menu_msg_ids

# ─── ASOSIY TSIKL ─────────────────────────────────────────────────────────────

def main():
    if not TELEGRAM_CHAT_IDS:
        log.error("TELEGRAM_CHAT_IDS yoki TELEGRAM_CHAT_ID .env da topilmadi!")
        sys.exit(1)

    log.info("Email-агент запущен")
    log.info(f"Пользователей Telegram: {len(TELEGRAM_CHAT_IDS)} шт.")
    send_telegram("🤖 *Email-агент запущен!*\nНовые письма от клиентов будут автоматически анализироваться.")

    # Ishga tushganda har bir admin uchun asosiy menyuni yuboramiz
    counts = fetch_all_counts()
    menu_msg_ids = load_json(MENU_MESSAGE_FILE, {})
    menu_msg_ids["last_counts"] = counts
    for cid in TELEGRAM_CHAT_IDS:
        mid = send_main_menu(cid, counts=counts)
        if mid:
            menu_msg_ids[cid] = mid
    save_json(MENU_MESSAGE_FILE, menu_msg_ids)

    seen_ids = set(load_json(SEEN_IDS_FILE, []))
    pending  = load_json(PENDING_FILE, {})

    saved = load_json(LAST_UPDATE_FILE, {})
    last_update_id = saved.get("last_update_id", 0)

    loop_count = 0
    # Badge holati: True bo'lsa 🔴 ko'rsatilmoqda
    badge_active = False

    while True:
        try:
            pending, last_update_id, menu_msg_ids = handle_telegram_updates(
                pending, last_update_id, menu_msg_ids
            )

            log.info("Проверяю почту...")
            new_emails = fetch_new_emails(seen_ids)

            if new_emails:
                for em in new_emails:
                    log.info(f"Yangi xat: {em['sender']}")
                    email_id = em["uid"]

                    seen_ids.add(email_id)
                    seen_ids = trim_seen_ids(seen_ids)
                    save_json(SEEN_IDS_FILE, list(seen_ids))

                    # pending ga saqlaymiz — "Открыть" bosilganda tahlil qilinadi
                    if email_id not in pending:
                        pending[email_id] = {
                            "sender":     em["sender"],
                            "subject":    em["subject"],
                            "body":       em["body"],
                            "history":    em.get("history", []),
                            "created_at": datetime.now().isoformat()
                        }
                        save_json(PENDING_FILE, pending)

                # Faqat badge yangilaymiz — alohida xabar yuborilmaydi
                counts = fetch_all_counts()
                menu_msg_ids["last_counts"] = counts
                save_json(MENU_MESSAGE_FILE, menu_msg_ids)
                update_menu_badge(menu_msg_ids, counts, new_badge=True)
                badge_active = True
                log.info(f"Badge yangilandi: {counts}")

            else:
                if badge_active:
                    counts = fetch_all_counts()
                    menu_msg_ids["last_counts"] = counts
                    save_json(MENU_MESSAGE_FILE, menu_msg_ids)
                    update_menu_badge(menu_msg_ids, counts, new_badge=False)
                    badge_active = False

            loop_count += 1
            if loop_count % 100 == 0:
                pending = cleanup_pending(pending)
                save_json(PENDING_FILE, pending)

            # Har daqiqada menyuni yangilaymiz — yangi xat bo'lmasa ham
            if not new_emails and not badge_active:
                counts = fetch_all_counts()
                menu_msg_ids["last_counts"] = counts
                save_json(MENU_MESSAGE_FILE, menu_msg_ids)
                update_menu_badge(menu_msg_ids, counts, new_badge=False)

        except Exception as e:
            log.error(f"Ошибка основного цикла: {e}")
            send_telegram(f"⚠️ *Произошла ошибка агента. Проверьте лог-файл.*")

        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
