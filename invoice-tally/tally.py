#!/usr/bin/env python3
"""Invoice tally: read every invoice/receipt in a folder, total up the month.

Usage:
    python tally.py [inbox_folder] [--month YYYY-MM] [--offline]

Reads PDFs / photos (jpg, png) from the inbox folder. With ANTHROPIC_API_KEY set,
Claude reads each document (works on scans, photos and any layout). Without a key
(or with --offline) a basic text reader is used, which only handles clean PDFs.

Output: monthly_report.xlsx  (Invoices sheet, Daily sheet, Suppliers sheet, Summary)
"""
import argparse, base64, json, os, re, sys
from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MODEL = "claude-sonnet-5-5"
SUPPORTED = {".pdf", ".jpg", ".jpeg", ".png"}
CACHE_NAME = ".extracted.json"  # so each file is only read once

PROMPT = (
    "This is an invoice or receipt received by a petrol station business. "
    "Extract the fields by calling record_invoice. Dates must be YYYY-MM-DD. "
    "Amounts are plain numbers (no currency symbols or thousands separators). "
    "total is the final amount payable/paid including tax. If it is a credit note, "
    "make the amounts negative. If a field is not on the document use null."
)

TOOL = {
    "name": "record_invoice",
    "description": "Record the details of one invoice or receipt.",
    "input_schema": {
        "type": "object",
        "properties": {
            "supplier": {"type": "string"},
            "invoice_number": {"type": ["string", "null"]},
            "invoice_date": {"type": ["string", "null"], "description": "YYYY-MM-DD"},
            "subtotal": {"type": ["number", "null"]},
            "tax": {"type": ["number", "null"]},
            "total": {"type": "number"},
            "currency": {"type": ["string", "null"]},
            "description": {"type": ["string", "null"], "description": "few words on what it is for"},
        },
        "required": ["supplier", "total"],
    },
}


# ---------- extraction ----------
def extract_with_claude(path: Path) -> dict:
    import anthropic

    client = anthropic.Anthropic()
    data = base64.standard_b64encode(path.read_bytes()).decode()
    ext = path.suffix.lower()
    if ext == ".pdf":
        block = {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": data}}
    else:
        mt = "image/png" if ext == ".png" else "image/jpeg"
        block = {"type": "image", "source": {"type": "base64", "media_type": mt, "data": data}}
    resp = client.messages.create(
        model=MODEL,
        max_tokens=1024,
        tools=[TOOL],
        tool_choice={"type": "tool", "name": "record_invoice"},
        messages=[{"role": "user", "content": [block, {"type": "text", "text": PROMPT}]}],
    )
    for part in resp.content:
        if part.type == "tool_use":
            return part.input
    raise RuntimeError("Claude returned no invoice data")


MONEY = r"[-]?[€£$]?\s*([\d,]+\.\d{2})"


def _money(s):
    return float(s.replace(",", ""))


def extract_offline(path: Path) -> dict:
    """Basic reader for clean text PDFs. Real-world scans/photos need Claude."""
    if path.suffix.lower() != ".pdf":
        raise RuntimeError("offline mode can only read PDFs (set ANTHROPIC_API_KEY for photos/scans)")
    from pypdf import PdfReader

    text = "\n".join(p.extract_text() or "" for p in PdfReader(path).pages)
    if not text.strip():
        raise RuntimeError("no text in PDF (scanned image) - needs Claude")
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    out = {"supplier": lines[0], "invoice_number": None, "invoice_date": None,
           "subtotal": None, "tax": None, "total": None, "currency": None, "description": None}
    m = re.search(r"(?:invoice|inv)\s*(?:no\.?|number|#)?\s*[:#]?\s*([A-Z0-9][A-Z0-9\-/]{2,})", text, re.I)
    if m:
        out["invoice_number"] = m.group(1)
    m = re.search(r"date\s*[:]?\s*(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})", text, re.I) or None
    if m:
        d, mo, y = map(int, m.groups())
        out["invoice_date"] = f"{y:04d}-{mo:02d}-{d:02d}"
    else:
        m = re.search(r"date\s*[:]?\s*(\d{4})-(\d{2})-(\d{2})", text, re.I)
        if m:
            out["invoice_date"] = "-".join(m.groups())
    for key, pat in (("subtotal", r"sub\s*-?total"), ("tax", r"(?:vat|tax|gst)[^\n\d]*"), ("total", r"(?:total due|amount due|grand total|^total)")):
        for l in lines:
            if re.search(pat, l, re.I) and (key != "total" or not re.search(r"sub", l, re.I)):
                mm = re.search(MONEY, l)
                if mm:
                    out[key] = _money(mm.group(1))
    sym = re.search(r"[€£$]", text)
    out["currency"] = sym.group(0) if sym else None
    if out["total"] is None:
        raise RuntimeError("could not find a total")
    return out


# ---------- processing ----------
def load_cache(folder: Path) -> dict:
    p = folder / CACHE_NAME
    return json.loads(p.read_text()) if p.exists() else {}


def process(folder: Path, offline: bool):
    cache = load_cache(folder)
    rows, problems = [], []
    use_claude = bool(os.environ.get("ANTHROPIC_API_KEY")) and not offline
    for f in sorted(p for p in folder.rglob("*") if p.suffix.lower() in SUPPORTED):
        key = f"{f.relative_to(folder)}:{f.stat().st_size}"
        try:
            if key in cache:
                rec = cache[key]
            else:
                rec = extract_with_claude(f) if use_claude else extract_offline(f)
                cache[key] = rec
            rec = dict(rec, file=str(f.relative_to(folder)))
            rows.append(rec)
        except Exception as e:
            problems.append((str(f.relative_to(folder)), str(e)))
    (folder / CACHE_NAME).write_text(json.dumps(cache, indent=1))
    return rows, problems, ("Claude" if use_claude else "offline text reader")


def dec(v):
    try:
        return Decimal(str(v)).quantize(Decimal("0.01"))
    except (InvalidOperation, TypeError):
        return None


def flag_duplicates(rows):
    seen = {}
    for r in rows:
        r["flag"] = ""
        if not r.get("invoice_date"):
            r["flag"] = "NO DATE - check"
        k = (str(r.get("supplier", "")).strip().lower(), str(r.get("invoice_number") or "").strip().lower())
        if k[1]:
            if k in seen:
                r["flag"] = f"DUPLICATE of {seen[k]} - not counted"
            else:
                seen[k] = r["file"]
        # sanity check: subtotal + tax should equal total
        s, t, tot = dec(r.get("subtotal")), dec(r.get("tax")), dec(r.get("total"))
        if s is not None and t is not None and tot is not None and abs(s + t - tot) > Decimal("0.02"):
            r["flag"] = (r["flag"] + "; " if r["flag"] else "") + "AMOUNTS DON'T ADD UP - check"


# ---------- report ----------
def write_report(rows, problems, month, out_path, engine):
    wb = Workbook()
    bold = Font(bold=True)
    head = PatternFill("solid", fgColor="DDEBF7")
    bad = PatternFill("solid", fgColor="FFE0E0")

    def header(ws, cols):
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill = bold, head
        ws.freeze_panes = "A2"

    counted = [r for r in rows if "DUPLICATE" not in r["flag"]
               and (not month or (r.get("invoice_date") or "").startswith(month))]
    outside = [r for r in rows if month and not (r.get("invoice_date") or "").startswith(month)]

    ws = wb.active
    ws.title = "Invoices"
    header(ws, ["Date", "Supplier", "Invoice #", "Subtotal", "Tax", "Total", "Description", "File", "Flag"])
    for r in sorted(rows, key=lambda r: (r.get("invoice_date") or "9999", r["supplier"])):
        ws.append([r.get("invoice_date"), r["supplier"], r.get("invoice_number"),
                   r.get("subtotal"), r.get("tax"), r.get("total"), r.get("description"), r["file"], r["flag"]])
        if r["flag"]:
            for c in ws[ws.max_row]:
                c.fill = bad
    for col in "DEF":
        for c in ws[col][1:]:
            c.number_format = "#,##0.00"

    # Daily
    wd = wb.create_sheet("Daily")
    header(wd, ["Date", "Invoices", "Total"])
    daily = defaultdict(lambda: [0, Decimal(0)])
    for r in counted:
        d = daily[r.get("invoice_date") or "NO DATE"]
        d[0] += 1
        d[1] += dec(r["total"])
    for d in sorted(daily):
        wd.append([d, daily[d][0], float(daily[d][1])])
    # Suppliers
    wsup = wb.create_sheet("Suppliers")
    header(wsup, ["Supplier", "Invoices", "Total"])
    sup = defaultdict(lambda: [0, Decimal(0)])
    for r in counted:
        s = sup[r["supplier"].strip()]
        s[0] += 1
        s[1] += dec(r["total"])
    for s in sorted(sup, key=lambda k: -sup[k][1]):
        wsup.append([s, sup[s][0], float(sup[s][1])])
    for w in (wd, wsup):
        for c in w["C"][1:]:
            c.number_format = "#,##0.00"

    grand = sum((dec(r["total"]) for r in counted), Decimal(0))
    tax = sum((dec(r.get("tax")) or Decimal(0) for r in counted), Decimal(0))
    sm = wb.create_sheet("Summary", 0)
    sm.append(["Month", month or "all"])
    sm.append(["Invoices counted", len(counted)])
    sm.append(["TOTAL FOR MONTH", float(grand)])
    sm.append(["of which tax", float(tax)])
    sm.append(["Excluded duplicates", sum(1 for r in rows if "DUPLICATE" in r["flag"])])
    sm.append(["Dated outside this month (not counted)", len(outside)])
    sm.append(["Needs a human check", sum(1 for r in counted if r["flag"]) + len(problems)])
    sm.append(["Read by", engine])
    sm["A3"].font = sm["B3"].font = Font(bold=True, size=14)
    sm["B3"].number_format = "#,##0.00"
    sm["B4"].number_format = "#,##0.00"
    if problems:
        sm.append([])
        sm.append(["FILES THAT COULD NOT BE READ"])
        sm.cell(sm.max_row, 1).font = bold
        for f, e in problems:
            sm.append([f, e])
    for w in wb.worksheets:
        for i, col in enumerate(w.columns, 1):
            w.column_dimensions[get_column_letter(i)].width = min(
                45, max(10, max(len(str(c.value or "")) for c in col) + 2))
    wb.save(out_path)
    return counted, grand


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inbox", nargs="?", default="inbox")
    ap.add_argument("--month", help="YYYY-MM (default: month of the newest invoice)")
    ap.add_argument("--offline", action="store_true")
    ap.add_argument("--out", default="monthly_report.xlsx")
    a = ap.parse_args()

    folder = Path(a.inbox)
    rows, problems, engine = process(folder, a.offline)
    if not rows and not problems:
        sys.exit(f"No invoices found in {folder}/")
    flag_duplicates(rows)
    month = a.month or max((r["invoice_date"] for r in rows if r.get("invoice_date")), default="")[:7]
    counted, grand = write_report(rows, problems, month, a.out, engine)

    print(f"Read {len(rows)} files with {engine}; {len(problems)} unreadable")
    for r in sorted(counted, key=lambda r: r.get("invoice_date") or ""):
        print(f"  {r.get('invoice_date')}  {r['supplier'][:28]:28} {str(r.get('invoice_number') or ''):12} {dec(r['total']):>10}"
              + (f"   <-- {r['flag']}" if r["flag"] else ""))
    for r in rows:
        if "DUPLICATE" in r["flag"]:
            print(f"  (excluded) {r['supplier']} {r.get('invoice_number')}  {r['flag']}")
    for f, e in problems:
        print(f"  !! {f}: {e}")
    print(f"\nMONTH {month}: {len(counted)} invoices, TOTAL = {grand:,.2f}\nReport: {a.out}")


if __name__ == "__main__":
    main()
