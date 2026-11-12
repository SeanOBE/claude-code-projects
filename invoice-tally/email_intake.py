"""Pull invoice attachments out of a dedicated mailbox into inbox/.

Settings come from environment variables (see .env.example):
    IMAP_HOST, IMAP_USER, IMAP_PASSWORD, [IMAP_FOLDER=INBOX]

Only unread emails are processed, and an email is marked read only after its
attachments are safely saved. Nothing is ever deleted from the mailbox.
"""
import email, hashlib, imaplib, json, os, re
from datetime import datetime
from email import policy
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path

KEEP = {".pdf", ".jpg", ".jpeg", ".png"}
MIN_IMAGE_BYTES = 40_000  # smaller images are almost always logos / signatures
LOG = ".intake_log.json"


def _safe(s):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", s).strip("_")[:80] or "file"


def save_attachments(raw: bytes, inbox: Path, log: dict) -> dict:
    """Save invoice-looking attachments from one raw email. Returns a result dict."""
    msg = email.message_from_bytes(raw, policy=policy.default)
    sender = parseaddr(msg.get("From", ""))[1] or "unknown"
    try:
        when = parsedate_to_datetime(msg["Date"]).strftime("%Y-%m-%d")
    except Exception:
        when = datetime.now().strftime("%Y-%m-%d")
    saved, dupes = [], 0
    for part in msg.walk():
        name = part.get_filename()
        if not name or part.is_multipart():
            continue
        ext = Path(name).suffix.lower()
        if ext not in KEEP:
            continue
        data = part.get_payload(decode=True) or b""
        if ext != ".pdf" and len(data) < MIN_IMAGE_BYTES:
            continue
        digest = hashlib.sha256(data).hexdigest()
        if digest in log["hashes"]:
            dupes += 1  # exact same file already received
            continue
        log["hashes"][digest] = f"{when} {sender}"
        dest = inbox / f"{when}_{_safe(sender.split('@')[-1])}_{_safe(name)}"
        n = 1
        while dest.exists():
            dest = dest.with_name(f"{dest.stem}_{n}{dest.suffix}")
            n += 1
        dest.write_bytes(data)
        saved.append(dest.name)
    if not saved and not dupes:
        log["no_attachment"].append({"date": when, "from": sender, "subject": str(msg.get("Subject", ""))})
    return {"from": sender, "subject": str(msg.get("Subject", "")), "saved": saved, "duplicates": dupes}


def fetch(inbox="inbox", conn=None) -> dict:
    inbox = Path(inbox)
    inbox.mkdir(exist_ok=True)
    logp = inbox / LOG
    log = json.loads(logp.read_text()) if logp.exists() else {"hashes": {}, "no_attachment": []}
    log["no_attachment"] = []  # this run's list only

    own = conn is None
    if own:
        conn = imaplib.IMAP4_SSL(os.environ["IMAP_HOST"])
        conn.login(os.environ["IMAP_USER"], os.environ["IMAP_PASSWORD"])
    conn.select(os.environ.get("IMAP_FOLDER", "INBOX"))
    _, ids = conn.search(None, "UNSEEN")
    results = []
    for num in ids[0].split():
        _, parts = conn.fetch(num, "(BODY.PEEK[])")  # PEEK: don't mark read yet
        raw = parts[0][1]
        results.append(save_attachments(raw, inbox, log))
        logp.write_text(json.dumps(log, indent=1))
        conn.store(num, "+FLAGS", "\\Seen")  # only after the files are on disk
    if own:
        conn.logout()
    return {"emails": len(results), "files": sum(len(r["saved"]) for r in results),
            "results": results, "no_attachment": log["no_attachment"]}


if __name__ == "__main__":
    r = fetch()
    print(f"{r['emails']} new emails, {r['files']} invoice files saved")
    for x in r["no_attachment"]:
        print(f"  no invoice attached: {x['from']} - {x['subject']}")
