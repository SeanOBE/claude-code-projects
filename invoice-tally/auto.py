#!/usr/bin/env python3
"""The hands-off job. Run it daily (cron / Task Scheduler):

    python auto.py                 # fetch new emails, refresh this month's report
    python auto.py --send          # ...and email the report to REPORT_TO
    python auto.py --month previous --send   # run on the 1st: send last month's final tally

Email settings (env vars): IMAP_HOST IMAP_USER IMAP_PASSWORD  (incoming)
                           SMTP_HOST SMTP_USER SMTP_PASSWORD REPORT_TO  (outgoing, optional)
"""
import argparse, os, smtplib
from datetime import date
from email.message import EmailMessage
from pathlib import Path

import email_intake, tally


def resolve_month(m):
    today = date.today()
    if m in (None, "current"):
        return today.strftime("%Y-%m")
    if m == "previous":
        y, mo = (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)
        return f"{y:04d}-{mo:02d}"
    return m


def build_email(s, intake):
    lines = [f"Invoice tally for {s['month']}", "",
             f"Invoices counted: {s['count']}",
             f"TOTAL: {s['total']:,.2f}", ""]
    if s["flagged"] or s["unreadable"] or intake["no_attachment"]:
        lines.append("Needs a quick look:")
        lines += [f"  - {f}: {why}" for f, why in s["flagged"]]
        lines += [f"  - {f}: could not be read ({e})" for f, e in s["unreadable"]]
        lines += [f"  - email from {x['from']} (\"{x['subject']}\") had no invoice attached" for x in intake["no_attachment"]]
    else:
        lines.append("Nothing needs checking.")
    lines += ["", "Full breakdown by day and supplier is in the attached spreadsheet."]
    return "\n".join(lines)


def send(s, intake):
    msg = EmailMessage()
    msg["Subject"] = f"Invoice tally {s['month']}: {s['total']:,.2f} ({s['count']} invoices)"
    msg["From"] = os.environ["SMTP_USER"]
    msg["To"] = os.environ["REPORT_TO"]
    msg.set_content(build_email(s, intake))
    data = Path(s["report"]).read_bytes()
    msg.add_attachment(data, maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       filename=f"invoices_{s['month']}.xlsx")
    port = int(os.environ.get("SMTP_PORT", "465"))
    cls = smtplib.SMTP_SSL if port == 465 else smtplib.SMTP
    with cls(os.environ["SMTP_HOST"], port) as smtp:
        if port != 465 and os.environ.get("SMTP_STARTTLS", "1") == "1":
            smtp.starttls()
        if os.environ.get("SMTP_PASSWORD"):
            smtp.login(os.environ["SMTP_USER"], os.environ["SMTP_PASSWORD"])
        smtp.send_message(msg)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--inbox", default="inbox")
    ap.add_argument("--month", default="current")
    ap.add_argument("--send", action="store_true")
    ap.add_argument("--offline", action="store_true")
    ap.add_argument("--skip-fetch", action="store_true")
    ap.add_argument("--out", default="monthly_report.xlsx")
    a = ap.parse_args()

    intake = {"emails": 0, "files": 0, "no_attachment": []}
    if not a.skip_fetch:
        intake = email_intake.fetch(a.inbox)
        print(f"Mailbox: {intake['emails']} new emails, {intake['files']} invoice files saved")
    s = tally.run(a.inbox, resolve_month(a.month), a.offline, a.out)
    if a.send:
        send(s, intake)
        print(f"Report emailed to {os.environ['REPORT_TO']}")


if __name__ == "__main__":
    main()
