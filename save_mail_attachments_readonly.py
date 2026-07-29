import argparse
import email
from email.header import decode_header, make_header
import imaplib
from pathlib import Path
import re


PROJECT_DIR = Path(r"D:\brushcustom-email-imap")
ENV_PATH = PROJECT_DIR / ".env"
WORK_DIR = PROJECT_DIR / "work"


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


def safe_filename(value):
    value = decode_mime(value) or "attachment.bin"
    value = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "_", value).strip()
    return value or "attachment.bin"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mailbox", default="INBOX")
    parser.add_argument("--query", required=True, help="case-insensitive match across from/subject/body headers")
    parser.add_argument("--limit", type=int, default=50)
    args = parser.parse_args()

    env = parse_env(ENV_PATH)
    out_dir = WORK_DIR / "attachments" / re.sub(r"[^A-Za-z0-9_.-]+", "_", args.query)[:80]
    out_dir.mkdir(parents=True, exist_ok=True)

    client = imaplib.IMAP4_SSL(env.get("IMAP_HOST", "imap.exmail.qq.com"), int(env.get("IMAP_PORT", "993")))
    try:
        client.login(env["EMAIL_ADDRESS"], env["EMAIL_IMAP_PASSWORD"])
        status, _ = client.select(args.mailbox, readonly=True)
        if status != "OK":
            raise RuntimeError(f"Unable to select mailbox: {args.mailbox}")
        status, data = client.search(None, "ALL")
        if status != "OK":
            raise RuntimeError("IMAP SEARCH failed")
        ids = data[0].split() if data and data[0] else []
        query = args.query.lower()
        saved = []
        for msg_id in reversed(ids[-args.limit:]):
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
            haystack = "\n".join([
                decode_mime(msg.get("From")),
                decode_mime(msg.get("To")),
                decode_mime(msg.get("Subject")),
            ]).lower()
            if query not in haystack:
                continue
            for index, part in enumerate(msg.walk(), start=1):
                filename = part.get_filename()
                if not filename:
                    continue
                payload = part.get_payload(decode=True)
                if payload is None:
                    continue
                target = out_dir / f"{index:02d}_{safe_filename(filename)}"
                target.write_bytes(payload)
                saved.append(str(target))
            break
    finally:
        client.logout()

    print(f"Saved attachments: {len(saved)}")
    for item in saved:
        print(item)


if __name__ == "__main__":
    main()
