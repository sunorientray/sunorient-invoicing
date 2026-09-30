"""Mail watcher: runs on GitHub Actions (no AI, no Claude usage).

Logs in to Gmail over IMAP, looks at every email that arrived since the last run,
and saves the ones that matter for invoicing (with their attachments) into
inbox/pending/<timestamp>_<category>_<subject>/ in this repo. The Claude run then
picks them up from there.

Env: GMAIL_USER, GMAIL_APP_PASSWORD (repo secrets). Optional: START_DAYS (first run lookback, default 3).
"""
import email, imaplib, json, os, re, sys, datetime as dt
from email import policy
from email.utils import parsedate_to_datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATE = os.path.join(ROOT, "inbox", "state.json")
PENDING = os.path.join(ROOT, "inbox", "pending")
RULES = json.load(open(os.path.join(ROOT, "watcher", "rules.json")))["rules"]
PREFIX = re.compile(r"^\s*((fwd?|fw|re)\s*:\s*)+", re.I)


def slug(s, n=60):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", s or "").strip("_")[:n] or "no_subject"


def text_of(msg):
    parts = []
    for p in msg.walk():
        if p.get_content_type() == "text/plain" and not p.get_filename():
            try:
                parts.append(p.get_content())
            except Exception:
                pass
    if not parts:
        for p in msg.walk():
            if p.get_content_type() == "text/html" and not p.get_filename():
                try:
                    parts.append(re.sub(r"<[^>]+>", " ", p.get_content()))
                except Exception:
                    pass
    return "\n".join(parts)


def attachments_of(msg):
    """All attachments, including ones inside forwarded-as-attachment emails."""
    out = []
    for p in msg.walk():
        fn = p.get_filename()
        if fn and not p.is_multipart():
            data = p.get_payload(decode=True)
            if data:
                out.append((fn, data))
    return out


def classify(frm, subject, body, att_names):
    subj = PREFIX.sub("", subject or "")
    for r in RULES:
        ok = True
        if "from" in r and not re.search(r["from"], frm or "", re.I): ok = False
        if ok and "subject" in r and not re.search(r["subject"], subj, re.I): ok = False
        if ok and "body" in r and not re.search(r["body"], body or "", re.I): ok = False
        if ok and "attachment" in r and not any(re.search(r["attachment"], a, re.I) for a in att_names): ok = False
        if ok:
            return r["category"], r.get("save", True)
    return None, False


def main():
    user, pw = os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"]
    state = json.load(open(STATE)) if os.path.exists(STATE) else {}
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(user, pw)
    M.select("INBOX", readonly=True)            # never changes anything in the mailbox
    uidvalidity = M.untagged_responses.get("UIDVALIDITY", [b"0"])[0].decode()
    same_box = state.get("uidvalidity") == uidvalidity and state.get("last_uid")
    if same_box:
        typ, data = M.uid("search", None, f"UID {int(state['last_uid']) + 1}:*")
    else:
        since = (dt.date.today() - dt.timedelta(days=int(os.environ.get("START_DAYS", "3")))).strftime("%d-%b-%Y")
        typ, data = M.uid("search", None, f"SINCE {since}")
    last = int(state["last_uid"]) if same_box else 0
    uids = [int(u) for u in data[0].split() if int(u) > last]   # "N:*" can return the last UID even if < N
    seen = saved = 0
    for uid in sorted(uids):
        typ, msgdata = M.uid("fetch", str(uid), "(RFC822)")
        raw = next((x[1] for x in msgdata if isinstance(x, tuple)), None)
        last = max(last, uid)
        if not raw:
            continue
        seen += 1
        msg = email.message_from_bytes(raw, policy=policy.default)
        frm, subject = str(msg.get("From", "")), str(msg.get("Subject", ""))
        body = text_of(msg)
        atts = attachments_of(msg)
        cat, save = classify(frm, subject, body, [a[0] for a in atts])
        if not (cat and save):
            continue
        try:
            when = parsedate_to_datetime(msg["Date"]).strftime("%Y%m%d-%H%M")
        except Exception:
            when = dt.datetime.utcnow().strftime("%Y%m%d-%H%M")
        folder = os.path.join(PENDING, f"{when}_{cat}_{slug(PREFIX.sub('', subject))}_{uid}")
        os.makedirs(folder, exist_ok=True)
        for fn, data in atts:
            open(os.path.join(folder, slug(fn, 120)), "wb").write(data)
        json.dump({"uid": uid, "category": cat, "from": frm, "to": str(msg.get("To", "")), "cc": str(msg.get("Cc", "")),
                   "subject": subject, "date": str(msg.get("Date", "")), "message_id": str(msg.get("Message-ID", "")),
                   "attachments": [slug(a[0], 120) for a in atts]},
                  open(os.path.join(folder, "meta.json"), "w"), indent=2)
        open(os.path.join(folder, "body.txt"), "w").write(body)
        saved += 1
    M.logout()
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    json.dump({"uidvalidity": uidvalidity, "last_uid": last,
               "last_run": dt.datetime.utcnow().isoformat(timespec="seconds") + "Z"}, open(STATE, "w"), indent=2)
    print(f"checked {seen} new emails, saved {saved} for invoicing")


if __name__ == "__main__":
    sys.exit(main())
