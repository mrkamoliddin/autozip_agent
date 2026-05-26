"""
Email Agent - Yandex Mail → OpenAI tahlil → Telegram
Yangi xatlar kelganda tahlil qilib Telegramga yuboradi.
Javob tugmasi orqali mijozga to'g'ridan-to'g'ri pochta yuborish mumkin.
Preview + tasdiqlash + imzo bilan.
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
from datetime import datetime
from pathlib import Path

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
        logging.StreamHandler(stream=open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1, closefd=False))
    ]
)
log = logging.getLogger(__name__)

# ─── SOZLAMALAR ───────────────────────────────────────────────────────────────
YANDEX_EMAIL    = os.environ["YANDEX_EMAIL"]
YANDEX_PASSWORD = os.environ["YANDEX_PASSWORD"]
TELEGRAM_TOKEN  = os.environ["TELEGRAM_TOKEN"]
TELEGRAM_CHAT_IDS = [
    os.environ["TELEGRAM_CHAT_ID"],   # 1-foydalanuvchi (Ravshan)
    "1326222702",                      # 2-foydalanuvchi (Kamoliddin)
]
OPENAI_KEY      = os.environ["OPENAI_API_KEY"]
CHECK_INTERVAL  = int(os.environ.get("CHECK_INTERVAL", "60"))

SEEN_IDS_FILE   = "seen_ids.json"
PENDING_FILE    = "pending_replies.json"
IMAP_HOST       = "imap.yandex.ru"
IMAP_PORT       = 993
SMTP_HOST       = "smtp.yandex.ru"
SMTP_PORT       = 465

LOGO_PATH       = Path(__file__).parent / "autozip_logo_email.png"

# Email imzosi (signature)
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
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return default

def save_json(path: str, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

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
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/plain":
                charset = part.get_content_charset() or "utf-8"
                body = part.get_payload(decode=True).decode(charset, errors="replace")
                break
            elif ct == "text/html" and not body:
                charset = part.get_content_charset() or "utf-8"
                raw_html = part.get_payload(decode=True).decode(charset, errors="replace")
                body = re.sub(r"<[^>]+>", " ", raw_html)
                body = re.sub(r"\s+", " ", body).strip()
    else:
        charset = msg.get_content_charset() or "utf-8"
        body = msg.get_payload(decode=True).decode(charset, errors="replace")
    return body.strip()

def extract_email_address(sender: str) -> str:
    match = re.search(r'<(.+?)>', sender)
    if match:
        return match.group(1)
    return sender.strip()

# ─── OPENAI ───────────────────────────────────────────────────────────────────

def openai_request(prompt: str, max_tokens: int = 1500) -> str:
    url = "https://api.openai.com/v1/chat/completions"
    headers = {"Authorization": f"Bearer {OPENAI_KEY}", "Content-Type": "application/json"}
    payload = {
        "model": "gpt-4o-mini",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]

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
{body[:50000]}
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
    """Javobni RUS tilida yozish — preview uchun."""
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
    """Rus tilidagi javobni mijoz tiliga tarjima qilish."""
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

# ─── SMTP (POCHTA YUBORISH) ───────────────────────────────────────────────────

def send_email_reply(to_address: str, subject: str, body: str):
    """Mijozga HTML + imzo + logo bilan pochta yuborish."""
    msg = MIMEMultipart("related")
    msg["From"]    = YANDEX_EMAIL
    msg["To"]      = to_address
    msg["Subject"] = f"Re: {subject}" if not subject.startswith("Re:") else subject

    # HTML tana
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

    # Logo qo'shish
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
                    log.error(f"Telegram xato: {e}")
                    break

def format_telegram_message(sender: str, subject: str, analysis: str) -> str:
    now = datetime.now().strftime("%d.%m.%Y %H:%M")
    return (
        f"📧 *НОВОЕ ПИСЬМО ОТ КЛИЕНТА*\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 *От кого:* `{sender}`\n"
        f"📌 *Тема:* {subject}\n"
        f"🕐 *Время:* {now}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        f"{analysis}"
    )

# ─── IMAP ─────────────────────────────────────────────────────────────────────

def fetch_sender_history(mail_conn, sender_email: str, limit: int = 5) -> list[dict]:
    """Bir mijozning oxirgi xatlarini olish (kontekst uchun)."""
    history = []
    try:
        # Ushbu email manzilidan kelgan xatlarni qidirish
        search_query = f'FROM "{sender_email}"'
        _, data = mail_conn.search(None, search_query)
        ids = data[0].split()
        # Oxirgi limit ta xat (yangi xat bundan oldin)
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

def fetch_new_emails(seen_ids: set) -> list[dict]:
    new_emails = []
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

            # Mijozning oldingi xatlarini olish
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
        mail.logout()
    except Exception as e:
        log.error(f"IMAP ошибка: {e}")
    return new_emails

# ─── TELEGRAM UPDATE HANDLER ──────────────────────────────────────────────────

def get_telegram_updates(offset: int = 0) -> list:
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getUpdates"
    try:
        resp = requests.get(url, params={"offset": offset, "timeout": 5}, timeout=30)
        if resp.ok:
            return resp.json().get("result", [])
    except requests.exceptions.Timeout:
        log.warning("Telegram getUpdates timeout")
    except Exception as e:
        log.error(f"Telegram getUpdates ошибка: {e}")
    return []

def handle_telegram_updates(pending: dict, last_update_id: int) -> tuple[dict, int]:
    updates = get_telegram_updates(last_update_id + 1)

    for update in updates:
        last_update_id = update["update_id"]

        # Tugma bosilishi
        if "callback_query" in update:
            cq = update["callback_query"]
            chat_id = str(cq["message"]["chat"]["id"])
            data = cq["data"]

            requests.post(
                f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/answerCallbackQuery",
                json={"callback_query_id": cq["id"]}, timeout=5
            )

            if chat_id not in TELEGRAM_CHAT_IDS:
                continue

            # Javob yozish tugmasi
            if data.startswith("reply:"):
                email_id = data.split(":", 1)[1]
                if email_id in pending:
                    pending[email_id]["waiting_reply"] = chat_id
                    pending[email_id].pop("preview_text", None)
                    save_json(PENDING_FILE, pending)
                    em = pending[email_id]
                    send_telegram(
                        f"✍️ *Напишите тезисы для ответа:*\n\n"
                        f"Клиент: `{em['sender']}`\n"
                        f"Тема: {em['subject']}\n\n"
                        f"_Напишите кратко — GPT составит полное письмо_",
                        chat_id=chat_id
                    )

            if data.startswith("send:"):
                email_id = data.split(":", 1)[1]
                if email_id in pending and "preview_text" in pending[email_id]:
                    em = pending[email_id]
                    russian_text = em["preview_text"]
                    try:
                        to_addr = extract_email_address(em["sender"])
                        send_telegram(f"⏳ *Перевожу и отправляю...*", chat_id=chat_id)
                        # Mijoz tiliga tarjima
                        final_text = translate_reply(russian_text, em["body"])
                        send_email_reply(to_addr, em["subject"], final_text)
                        send_telegram(
                            f"✅ *Письмо отправлено!*\n📬 *Кому:* `{to_addr}`",
                            chat_id=chat_id
                        )
                        log.info(f"Письмо отправлено: {to_addr}")
                        del pending[email_id]
                        save_json(PENDING_FILE, pending)
                    except Exception as e:
                        log.error(f"Ошибка отправки: {e}")
                        send_telegram(f"❌ *Ошибка:* `{e}`", chat_id=chat_id)

            # Qayta yozish tugmasi
            elif data.startswith("rewrite:"):
                email_id = data.split(":", 1)[1]
                if email_id in pending:
                    pending[email_id]["waiting_reply"] = chat_id
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

            if not text or text.startswith("/"):
                continue
            if chat_id not in TELEGRAM_CHAT_IDS:
                continue

            # Javob kutayotgan xat bormi?
            waiting_email_id = None
            waiting_mode = None
            for eid, em in pending.items():
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

                # Tahrirlash rejimi — matnni to'g'ridan preview sifatida saqlash
                if waiting_mode == "edit":
                    pending[waiting_email_id]["preview_text"] = text
                    pending[waiting_email_id].pop("waiting_edit", None)
                    save_json(PENDING_FILE, pending)

                    to_addr = extract_email_address(em["sender"])
                    preview_msg = (
                        f"📝 *Ваш текст — проверьте:*\n\n"
                        f"📬 *Кому:* `{to_addr}`\n"
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
                            f"📬 *Кому:* `{to_addr}`\n"
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
                        send_telegram(f"❌ *Ошибка:* `{e}`", chat_id=chat_id)
                        pending[waiting_email_id].pop("waiting_reply", None)
                        save_json(PENDING_FILE, pending)

    return pending, last_update_id

# ─── ASOSIY TSIKL ─────────────────────────────────────────────────────────────

def main():
    log.info("Email Agent запущен")
    send_telegram("🤖 *Email-агент запущен!*\nНовые письма от клиентов будут автоматически анализироваться.")

    seen_ids = set(load_json(SEEN_IDS_FILE, []))
    pending  = load_json(PENDING_FILE, {})
    last_update_id = 0

    while True:
        try:
            pending, last_update_id = handle_telegram_updates(pending, last_update_id)

            log.info("Проверка почты...")
            new_emails = fetch_new_emails(seen_ids)

            for em in new_emails:
                log.info(f"Анализируется: {em['sender']}")
                analysis = analyze_email(em["sender"], em["subject"], em["body"], em.get("history", []))
                message  = format_telegram_message(em["sender"], em["subject"], analysis)

                email_id = em["uid"]
                pending[email_id] = {
                    "sender":  em["sender"],
                    "subject": em["subject"],
                    "body":    em["body"],
                    "history": em.get("history", [])
                }
                save_json(PENDING_FILE, pending)

                reply_markup = {
                    "inline_keyboard": [[
                        {"text": "✍️ Ответить клиенту", "callback_data": f"reply:{email_id}"}
                    ]]
                }

                send_telegram(message, reply_markup=reply_markup)
                seen_ids.add(email_id)
                save_json(SEEN_IDS_FILE, list(seen_ids))
                log.info(f"Отправлено: {em['sender']}")
                time.sleep(2)

        except Exception as e:
            log.error(f"Ошибка основного цикла: {e}")
            send_telegram(f"⚠️ Ошибка агента: `{e}`")

        time.sleep(10)


if __name__ == "__main__":
    main()
