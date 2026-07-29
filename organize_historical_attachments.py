import email
from email.header import decode_header, make_header
from html.parser import HTMLParser
import html
import imaplib
import json
from pathlib import Path
import re
from urllib.request import Request, urlopen


PROJECT_DIR = Path(r"D:\brushcustom-email-imap")
ENV_PATH = PROJECT_DIR / ".env"
ATTACHMENTS_DIR = PROJECT_DIR / "attachments"


RECORDS = [
    {
        "key": "oyerli@emremakinasan.com.tr",
        "folder": "2026-06-16_Emre-Makina_Mehmet-Okan-Yerli",
        "query": "RFQ 6000274612",
    },
    {
        "key": "glam@tnt-group.com",
        "folder": "2026-06-18_TNT-Group_Gigi-Lam",
        "query": "Body brush and nail brush",
    },
    {
        "key": "admin02@lienminhine.vn",
        "folder": "2026-06-26_Lien-Minh_Trang",
        "query": "Inquiry for Wheel brush",
    },
    {
        "key": "general@advanxtech.com.sg",
        "folder": "2026-07-09_ADVANX_Rena-Goh",
        "query": "RFQ_ 2.60mm Stainless steel wire cleaning brush",
    },
    {
        "key": "mcipriano@acncutting.com",
        "folder": "2026-07-10_ACN-Cutting_Maria-Clara-Cabrita",
        "query": "ACN | QUOTATION REQUEST",
    },
    {
        "key": "rfinale@cisa.net",
        "folder": "2026-07-14_CISA_Rita-Finale",
        "query": "Custom Brush Manufacturing Inquiry",
    },
    {
        "key": "kodchaphan@noventa.com",
        "folder": "2026-07-16_Noventa_Kodchaphan",
        "query": "RFQ ROUND BRUSH",
    },
    {
        "key": "jasper.poppele@assetcool.com",
        "folder": "2026-07-17_AssetCool_Jasper-Poppele",
        "query": "Re: Previous Communication and Request for Collaboration",
    },
    {
        "key": "jfarina@nichiha.com",
        "folder": "2026-07-20_Nichiha_Jim-Farina",
        "urls": [
            "https://brushcustom.com/wp-content/uploads/elementor/forms/6a5e2d7a5e430.jpg",
        ],
    },
    {
        "key": "zmielke@americanflexible.com",
        "folder": "2026-07-23_American-Flexible_Zory",
        "urls": [
            "https://brushcustom.com/wp-content/uploads/elementor/forms/6a6215f58e30f.jpg",
        ],
    },
]


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


def clean_text(value):
    value = value.replace("\r", "\n")
    value = re.sub(r"\n{3,}", "\n\n", value)
    value = re.sub(r"[ \t]{2,}", " ", value)
    return value.strip()


def strip_html(value):
    extractor = HtmlTextExtractor()
    extractor.feed(value)
    return html.unescape(extractor.get_text())


def body_from_message(msg):
    text_parts = []
    html_parts = []
    parts = msg.walk() if msg.is_multipart() else [msg]
    for part in parts:
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


def safe_filename(value):
    value = decode_mime(value) or "attachment.bin"
    value = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "_", value).strip()
    return value or "attachment.bin"


def message_to_text(msg):
    headers = [
        ("Date", decode_mime(msg.get("Date"))),
        ("From", decode_mime(msg.get("From"))),
        ("To", decode_mime(msg.get("To"))),
        ("Cc", decode_mime(msg.get("Cc"))),
        ("Subject", decode_mime(msg.get("Subject"))),
    ]
    body = body_from_message(msg)
    return "\n".join(f"{k}: {v}" for k, v in headers if v) + "\n\n" + body + "\n"


def message_haystack(msg):
    return "\n".join([
        decode_mime(msg.get("From")),
        decode_mime(msg.get("To")),
        decode_mime(msg.get("Cc")),
        decode_mime(msg.get("Subject")),
        body_from_message(msg),
    ]).lower()


def find_message(client, query, limit=800):
    status, data = client.search(None, "ALL")
    if status != "OK":
        return None
    ids = data[0].split() if data and data[0] else []
    query = query.lower()
    for msg_id in reversed(ids[-limit:]):
        status, fetched = client.fetch(msg_id, "(BODY.PEEK[])")
        if status != "OK":
            continue
        raw = None
        for item in fetched or []:
            if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], bytes):
                raw = item[1]
                break
        if raw is None:
            continue
        msg = email.message_from_bytes(raw)
        if query in message_haystack(msg):
            return msg
    return None


def mailbox_name_from_list_item(item):
    text = item.decode("utf-8", "replace") if isinstance(item, bytes) else str(item)
    match = re.search(r' "([^"]+)"$', text)
    if match:
        return match.group(1)
    parts = text.rsplit(" ", 1)
    return parts[-1].strip('"') if parts else text


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


def find_message_all_mailboxes(client, mailboxes, query):
    for mailbox in mailboxes:
        status, _ = select_readonly(client, mailbox)
        if status != "OK":
            continue
        msg = find_message(client, query)
        if msg is not None:
            return msg, mailbox
    return None, None


def save_attachments_from_message(msg, target_dir):
    saved = []
    for index, part in enumerate(msg.walk(), start=1):
        filename = part.get_filename()
        if not filename:
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        target = target_dir / f"{index:02d}_{safe_filename(filename)}"
        target.write_bytes(payload)
        saved.append(target)
    return saved


def download_url(url, target_dir):
    name = safe_filename(url.rstrip("/").split("/")[-1] or "uploaded_file")
    target = target_dir / name
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(req, timeout=30) as response:
        target.write_bytes(response.read())
    return target


def main():
    env = parse_env(ENV_PATH)
    ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)
    client = imaplib.IMAP4_SSL(env.get("IMAP_HOST", "imap.exmail.qq.com"), int(env.get("IMAP_PORT", "993")))
    summary = {}
    try:
        client.login(env["EMAIL_ADDRESS"], env["EMAIL_IMAP_PASSWORD"])
        status, listed = client.list()
        if status != "OK":
            raise RuntimeError("Unable to list mailboxes")
        mailboxes = [mailbox_name_from_list_item(item) for item in listed or []]
        preferred = [env.get("MAILBOX", "INBOX") or "INBOX"]
        mailboxes = preferred + [box for box in mailboxes if box not in preferred]
        for record in RECORDS:
            target_dir = ATTACHMENTS_DIR / record["folder"]
            target_dir.mkdir(parents=True, exist_ok=True)
            saved_files = []
            original_urls = record.get("urls", [])
            msg = None
            matched_mailbox = None
            if record.get("query"):
                msg, matched_mailbox = find_message_all_mailboxes(client, mailboxes, record["query"])
            if msg is not None:
                (target_dir / "inquiry_email.txt").write_text(message_to_text(msg), encoding="utf-8")
                saved_files.extend(save_attachments_from_message(msg, target_dir))
                (target_dir / "matched_mailbox.txt").write_text(str(matched_mailbox), encoding="utf-8")
            else:
                (target_dir / "notes.txt").write_text(
                    "No matching email message was found by the organizer script.\n",
                    encoding="utf-8",
                )
            for url in original_urls:
                try:
                    saved_files.append(download_url(url, target_dir))
                except Exception as exc:
                    (target_dir / "download_errors.txt").write_text(f"{url}\n{exc}\n", encoding="utf-8")
            (target_dir / "original_urls.txt").write_text("\n".join(original_urls) + ("\n" if original_urls else ""), encoding="utf-8")
            summary[record["key"]] = {
                "folder": str(target_dir),
                "files": [str(path.name) for path in saved_files],
                "original_urls": original_urls,
                "matched_mailbox": matched_mailbox,
            }
    finally:
        client.logout()
    summary_path = ATTACHMENTS_DIR / "_attachment_summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"Summary: {summary_path}")


if __name__ == "__main__":
    main()
