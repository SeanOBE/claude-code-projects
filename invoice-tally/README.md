# Invoice Tally

Drop supplier invoices (PDF / jpg / png) into `inbox/`, then run:

    pip install anthropic openpyxl pypdf
    export ANTHROPIC_API_KEY=...        # lets Claude read scans, photos, any layout
    python tally.py inbox --month 2026-11

Produces `monthly_report.xlsx`: Summary (month total), Invoices, Daily, Suppliers.
Duplicates (same supplier + invoice number) are excluded and anything odd
(no date, subtotal+tax != total, unreadable file) is flagged for a human.
Each file is read once and cached in `inbox/.extracted.json`.

`python make_samples.py` creates fake test invoices. `--offline` uses a basic text reader (clean PDFs only).
