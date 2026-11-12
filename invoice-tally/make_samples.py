#!/usr/bin/env python3
"""Generate fake supplier invoices into inbox/ so the system can be tested."""
import random, sys
from decimal import Decimal
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

random.seed(7)
out = Path(sys.argv[1] if len(sys.argv) > 1 else "inbox")
out.mkdir(exist_ok=True)

SUPPLIERS = [("Northern Fuel Distributors Ltd", "FD", 2500, 9000, "Unleaded & diesel delivery"),
             ("Fresh Daily Foods", "FDF", 120, 600, "Shop stock - sandwiches & dairy"),
             ("Castle Drinks Wholesale", "CDW", 200, 900, "Soft drinks & energy drinks"),
             ("Lube & Oil Supplies", "LOS", 80, 400, "Engine oil, screen wash"),
             ("Clean Sweep Services", "CSS", 60, 150, "Forecourt cleaning")]
VAT = Decimal("0.20")
made = []


def invoice(name, prefix, num, day, sub, desc, fname):
    sub = Decimal(sub).quantize(Decimal("0.01"))
    vat = (sub * VAT).quantize(Decimal("0.01"))
    c = canvas.Canvas(str(out / fname), pagesize=A4)
    c.setFont("Helvetica-Bold", 16); c.drawString(50, 780, name)
    c.setFont("Helvetica", 11)
    c.drawString(50, 750, f"Invoice No: {prefix}-{num}")
    c.drawString(50, 735, f"Date: {day:02d}/11/2026")
    c.drawString(50, 710, f"Description: {desc}")
    c.drawString(50, 660, f"Subtotal: £{sub:,}")
    c.drawString(50, 645, f"VAT (20%): £{vat:,}")
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 625, f"TOTAL DUE: £{sub + vat:,}")
    c.save()
    return sub + vat


n = 1000
for day in range(2, 29):
    if day % 7 in (0, 1) and random.random() < .6:
        continue
    for _ in range(random.randint(1, 3)):
        name, pre, lo, hi, desc = random.choice(SUPPLIERS)
        n += 1
        sub = random.uniform(lo, hi)
        t = invoice(name, pre, n, day, sub, desc, f"{pre}-{n}.pdf")
        made.append(t)

# a supplier re-sending an invoice they already sent (should NOT be counted twice)
first = sorted(out.glob("FD-*.pdf"))[0]
(out / ("RESENT_" + first.name)).write_bytes(first.read_bytes())
# duplicate re-send is detected by invoice number, so make bytes differ slightly
(out / ("RESENT_" + first.name)).write_bytes(first.read_bytes() + b"\n%resent")
print(f"Made {len(made)} invoices (+1 resent duplicate). True total of unique invoices: {sum(made):,.2f}")
