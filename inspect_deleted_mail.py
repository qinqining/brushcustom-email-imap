import argparse
import email
from email.header import decode_header, make_header
from email.utils import parseaddr
import html
from html.parser import HTMLParser
import imaplib
import json
from pathlib import Path
import re


PROJECT_DIR = Path(r"D:\brushcustom-email-imap")
ENV_PATH = PROJECT_DIR / ".env"
WORK_DIR = PROJECT_DIR / "work"


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
    return html.unescape(extractor.get_text())


def clean_text(value):
    value = value.replace("\r", "\n")
    value = re.sub(r"\n{3,}", "\n\n", value)
    value = re.sub(r"[ \t]{2,}", " ", value)
    return value.strip()


def body_from_message(msg):
    text_parts = []
    html_parts = []
    for part in msg.walk() if msg.is_multipart() else [msg]:
        if (part.get_content_disposition() or "").lower() == "attachment":
            continue
        content_type = part.get_content_type()
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
    if text_parts:
        return clean_text("\n\n".join(text_parts))
    return clean_text(strip_html("\n\n".join(html_parts)))


def mailbox_name_from_list_item(item):
    text = item.decode("utf-8", "replace") if isinstance(item, bytes) else str(item)
    match = re.search(r' "([^"]+)"$', text)
    if match:
        return match.group(1)
    parts = text.rsplit(" ", 1)
    return parts[-1].strip('"') if parts else text


def is_deleted_mailbox(name):
    low = name.lower()
    return any(token in low for token in ["trash", "deleted", "delete", "已删除", "deleted messages"])


def fetch_message(client, msg_id):
    status, fetched = client.fetch(msg_id, "(BODY.PEEK[])")
    if status != "OK":
        return None
    for item in fetched or []:
        if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], bytes):
            msg = email.message_from_bytes(item[1])
            sender = decode_mime(msg.get("From"))
            name, address = parseaddr(sender)
            return {
                "imap_id": msg_id.decode("ascii", errors="ignore"),
                "date": decode_mime(msg.get("Date")),
                "from": sender,
                "from_name": name,
                "from_email": address,
                "reply_to": decode_mime(msg.get("Reply-To")),
                "subject": decode_mime(msg.get("Subject")),
                "body": body_from_message(msg),
                "attachments": [decode_mime(part.get_filename()) for part in msg.walk() if part.get_filename()],
            }
    return None


def select_readonly(client, mailbox):
    candidates = [mailbox]
    if not (mailbox.startswith('"') and mailbox.endswith('"')):
        candidates.append(f'"{mailbox}"')
    for candidate in candidates:
        try:
            status, data = client.select(candidate, readonly=True)
        except imaplib.IMAP4.error:
            continue
        if status == "OK":
            return status, data
    return "NO", None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--subject", default="")
    parser.add_argument("--limit", type=int, default=10)
    args = parser.parse_args()

    env = parse_env(ENV_PATH)
    client = imaplib.IMAP4_SSL(env.get("IMAP_HOST", "imap.exmail.qq.com"), int(env.get("IMAP_PORT", "993")))
    messages = []
    mailboxes = []
    try:
        client.login(env["EMAIL_ADDRESS"], env["EMAIL_IMAP_PASSWORD"])
        status, listed = client.list()
        if status != "OK":
            raise RuntimeError("IMAP LIST failed")
        mailboxes = [mailbox_name_from_list_item(item) for item in listed or []]
        deleted_boxes = [name for name in mailboxes if is_deleted_mailbox(name)]
        if not deleted_boxes:
            deleted_boxes = ["Trash", "Deleted Messages", "Deleted", "已删除"]

        for mailbox in deleted_boxes:
            status, _ = select_readonly(client, mailbox)
            if status != "OK":
                continue
            status, data = client.search(None, "ALL")
            if status != "OK":
                continue
            ids = data[0].split() if data and data[0] else []
            ids = ids[-args.limit :]
            for msg_id in reversed(ids):
                message = fetch_message(client, msg_id)
                if not message:
                    continue
                message["mailbox"] = mailbox
                if args.subject and args.subject.lower() not in message["subject"].lower():
                    continue
                messages.append(message)
    finally:
        client.logout()

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    output = WORK_DIR / "deleted_mail_inspection.json"
    output.write_text(json.dumps({"mailboxes": mailboxes, "messages": messages}, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Saved: {output}")
    print(f"Matched messages: {len(messages)}")
    for idx, item in enumerate(messages, start=1):
        preview = re.sub(r"\s+", " ", item["body"])[:500]
        print(f"{idx}. [{item['mailbox']}] {item['date']} | {item['from']} | {item['subject']}")
        print(preview)


if __name__ == "__main__":
    main()
