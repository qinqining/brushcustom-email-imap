import argparse
import email
from email.header import decode_header, make_header
import html
from html.parser import HTMLParser
import imaplib
import json
from pathlib import Path
import re
from datetime import datetime, timedelta


PROJECT_DIR = Path(r"D:\brushcustom-email-imap")
ENV_PATH = PROJECT_DIR / ".env"
WORK_DIR = PROJECT_DIR / "work"
MUTATING_IMAP_METHODS = {
    "append",
    "close",
    "copy",
    "create",
    "delete",
    "expunge",
    "rename",
    "store",
    "subscribe",
    "uid",
    "unsubscribe",
}


class ReadOnlyIMAP:
    """Tiny allow-list wrapper: login, read-only select, search, BODY.PEEK fetch."""

    def __init__(self, host, port):
        self._client = imaplib.IMAP4_SSL(host, port)

    def login(self, address, password):
        return self._client.login(address, password)

    def select(self, mailbox):
        return self._client.select(mailbox, readonly=True)

    def search(self, *args):
        return self._client.search(*args)

    def fetch_raw_message(self, msg_id):
        status, fetched = self._client.fetch(msg_id, "(BODY.PEEK[])")
        if status != "OK":
            return status, None
        for item in fetched:
            if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], bytes):
                return status, item[1]
        return status, None

    def logout(self):
        return self._client.logout()

    def __getattr__(self, name):
        if name.lower() in MUTATING_IMAP_METHODS:
            raise RuntimeError(f"IMAP method '{name}' is disabled by read-only policy")
        raise AttributeError(name)


class HtmlTextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        if data and data.strip():
            self.parts.append(data.strip())

    def get_text(self):
        return "\n".join(self.parts)


def parse_env(path):
    env = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        env[key.strip()] = value.strip().strip('"').strip("'")
    return env


def decode_mime(value):
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def strip_html(value):
    extractor = HtmlTextExtractor()
    extractor.feed(value)
    text = extractor.get_text()
    return html.unescape(text)


def clean_text(value):
    value = value.replace("\r", "\n")
    value = re.sub(r"\n{3,}", "\n\n", value)
    value = re.sub(r"[ \t]{2,}", " ", value)
    return value.strip()


def body_from_message(msg):
    text_parts = []
    html_parts = []

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = (part.get_content_disposition() or "").lower()
            if disposition == "attachment":
                continue
            if content_type not in {"text/plain", "text/html"}:
                continue
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            charset = part.get_content_charset() or "utf-8"
            try:
                decoded = payload.decode(charset, errors="replace")
            except LookupError:
                decoded = payload.decode("utf-8", errors="replace")
            if content_type == "text/plain":
                text_parts.append(decoded)
            else:
                html_parts.append(decoded)
    else:
        payload = msg.get_payload(decode=True)
        if payload is not None:
            charset = msg.get_content_charset() or "utf-8"
            try:
                decoded = payload.decode(charset, errors="replace")
            except LookupError:
                decoded = payload.decode("utf-8", errors="replace")
            if msg.get_content_type() == "text/html":
                html_parts.append(decoded)
            else:
                text_parts.append(decoded)

    if text_parts:
        return clean_text("\n\n".join(text_parts))
    return clean_text(strip_html("\n\n".join(html_parts)))


def imap_date(value):
    return value.strftime("%d-%b-%Y")


def parse_iso_date(value):
    return datetime.strptime(value, "%Y-%m-%d")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", default="2026-07-10")
    parser.add_argument("--end", default=None, help="exclusive end date, YYYY-MM-DD")
    args = parser.parse_args()

    start = parse_iso_date(args.start)
    end = parse_iso_date(args.end) if args.end else start + timedelta(days=1)

    env = parse_env(ENV_PATH)
    address = env.get("EMAIL_ADDRESS", "")
    password = env.get("EMAIL_IMAP_PASSWORD", "")
    host = env.get("IMAP_HOST", "imap.exmail.qq.com")
    port = int(env.get("IMAP_PORT", "993"))
    mailbox = env.get("MAILBOX", "INBOX") or "INBOX"

    if not address or not password:
        raise SystemExit("EMAIL_ADDRESS or EMAIL_IMAP_PASSWORD is empty in .env")

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    client = ReadOnlyIMAP(host, port)
    try:
        client.login(address, password)
        status, _ = client.select(mailbox)
        if status != "OK":
            raise RuntimeError(f"Unable to select mailbox: {mailbox}")

        criteria = f'(SINCE "{imap_date(start)}" BEFORE "{imap_date(end)}")'
        status, data = client.search(None, criteria)
        if status != "OK":
            raise RuntimeError("IMAP SEARCH failed")
        ids = data[0].split() if data and data[0] else []

        messages = []
        for msg_id in ids:
            status, raw = client.fetch_raw_message(msg_id)
            if status != "OK" or raw is None:
                continue
            msg = email.message_from_bytes(raw)
            sender = decode_mime(msg.get("From"))
            subject = decode_mime(msg.get("Subject"))
            date = decode_mime(msg.get("Date"))
            body = body_from_message(msg)
            attachments = []
            for part in msg.walk():
                filename = decode_mime(part.get_filename())
                if filename:
                    attachments.append(filename)
            messages.append(
                {
                    "imap_id": msg_id.decode("ascii", errors="ignore"),
                    "message_id": msg.get("Message-ID", ""),
                    "date": date,
                    "from": sender,
                    "reply_to": decode_mime(msg.get("Reply-To")),
                    "to": decode_mime(msg.get("To")),
                    "subject": subject,
                    "body": body,
                    "attachments": attachments,
                }
            )
    finally:
        client.logout()

    output_path = WORK_DIR / f"emails_{args.start}_to_{end.strftime('%Y-%m-%d')}.json"
    output_path.write_text(json.dumps(messages, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Fetched {len(messages)} message(s) from {args.start} to {end.strftime('%Y-%m-%d')} (exclusive).")
    print(f"Saved: {output_path}")
    for idx, item in enumerate(messages, start=1):
        body_preview = re.sub(r"\s+", " ", item["body"])[:140]
        print(f"{idx}. {item['date']} | {item['from']} | {item['subject']} | {body_preview}")


if __name__ == "__main__":
    main()
