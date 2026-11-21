"""Invoice tracker + live Excel dashboard generator (XLSX + PNG preview).

Data is fictional and generated with a fixed seed so the preview matches the workbook.
Usage: python build_dashboard.py
       -> output/invoice-dashboard.xlsx and docs/preview.png
"""
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
REG = ["segoeui.ttf", "Arial.ttf", "DejaVuSans.ttf"]
BOLD = ["segoeuib.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf"]


def _font(names, size):
    """First available TrueType font (Windows, macOS, Linux), else Pillow's default."""
    for n in names:
        try:
            return ImageFont.truetype(n, size)
        except OSError:
            continue
    return ImageFont.load_default()


AS_OF = date(2026, 6, 30)
REGIONS = ["UK", "Ireland", "France", "Germany"]
PRODUCTS = {"Starter": 49, "Pro": 129, "Enterprise": 390}
REPS = ["Alice", "Ben", "Chloe", "David"]
MONTHS = [date(2026, m, 1) for m in range(1, 7)]
NAVY, GREEN, RED, GREY = "1F3A5F", "1F7A4D", "B23B3B", "F2F2F2"


def make_rows(n=150, seed=7):
    rnd = random.Random(seed)
    rows = []
    for i in range(n):
        d = date(2026, 1, 1) + timedelta(days=rnd.randrange(0, 181))
        prod = rnd.choices(list(PRODUCTS), weights=[5, 4, 1])[0]
        qty = rnd.randint(1, 12)
        paid = rnd.random() < (0.9 if d < date(2026, 5, 1) else 0.55)
        rows.append({
            "inv": f"INV-{1001 + i}", "date": d, "region": rnd.choice(REGIONS),
            "product": prod, "rep": rnd.choice(REPS), "qty": qty,
            "price": PRODUCTS[prod], "due": d + timedelta(days=30),
            "status": "Paid" if paid else "Unpaid",
        })
    rows.sort(key=lambda r: r["date"])
    return rows


def aggregates(rows):
    amt = lambda r: r["qty"] * r["price"]
    total = sum(amt(r) for r in rows)
    unpaid = [r for r in rows if r["status"] == "Unpaid"]
    outstanding = sum(amt(r) for r in unpaid)
    overdue = sum(amt(r) for r in unpaid if (AS_OF - r["due"]).days > 0)
    monthly = [sum(amt(r) for r in rows if r["date"].month == m.month) for m in MONTHS]
    by_region = [sum(amt(r) for r in rows if r["region"] == g) for g in REGIONS]
    return {"total": total, "paid": total - outstanding, "outstanding": outstanding,
            "overdue": overdue, "count": len(rows), "monthly": monthly, "region": by_region}


def build_xlsx(rows, path):
    wb = Workbook()
    thin = Side(style="thin", color="D0D0D0")
    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor=NAVY)

    # --- Data sheet
    ws = wb.active
    ws.title = "Invoices"
    heads = ["Invoice", "Date", "Region", "Product", "Sales rep", "Qty", "Unit price",
             "Amount", "Due date", "Status", "Days overdue"]
    ws.append(heads)
    for i, r in enumerate(rows, start=2):
        ws.append([r["inv"], r["date"], r["region"], r["product"], r["rep"], r["qty"],
                   r["price"], f"=F{i}*G{i}", r["due"], r["status"],
                   f'=IF(J{i}="Unpaid",MAX(0,Dashboard!$C$3-I{i}),0)'])
    last = len(rows) + 1
    tab = Table(displayName="Invoices", ref=f"A1:K{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    for col, w in zip("ABCDEFGHIJK", (12, 12, 10, 12, 11, 6, 11, 12, 12, 9, 13)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2, max_row=last):
        row[1].number_format = row[8].number_format = "dd/mm/yyyy"
        row[6].number_format = row[7].number_format = 'Â£#,##0'
    ws.conditional_formatting.add(
        f"K2:K{last}", CellIsRule(operator="greaterThan", formula=["0"],
                                  font=Font(color=RED, bold=True),
                                  fill=PatternFill("solid", fgColor="FBEDED")))
    ws.freeze_panes = "A2"

    # --- Dashboard sheet
    db = wb.create_sheet("Dashboard", 0)
    db.sheet_view.showGridLines = False
    db["B1"] = "Sales & invoices dashboard"
    db["B1"].font = Font(size=18, bold=True, color=NAVY)
    db["B3"], db["C3"] = "Report date", AS_OF
    db["C3"].number_format = "dd/mm/yyyy"
    db["B3"].font = Font(italic=True, color="666666")

    rng = lambda c: f"Invoices!${c}$2:${c}${last}"
    kpis = [
        ("Total invoiced", f"=SUM({rng('H')})"),
        ("Paid", f'=SUMIFS({rng("H")},{rng("J")},"Paid")'),
        ("Outstanding", f'=SUMIFS({rng("H")},{rng("J")},"Unpaid")'),
        ("Overdue", f'=SUMIFS({rng("H")},{rng("K")},">0")'),
        ("Invoices", f"=COUNTA({rng('A')})"),
    ]
    for j, (label, formula) in enumerate(kpis):
        col = 2 + j * 2
        lc, vc = db.cell(row=5, column=col), db.cell(row=6, column=col)
        lc.value, vc.value = label, formula
        lc.font = Font(color="666666", size=10)
        vc.font = Font(bold=True, size=16, color=RED if label == "Overdue" else NAVY)
        vc.number_format = "0" if label == "Invoices" else 'Â£#,##0'
        for c in (lc, vc):
            c.fill = PatternFill("solid", fgColor=GREY)
            c.alignment = Alignment(horizontal="center")
        db.column_dimensions[lc.column_letter].width = 16

    db["B9"], db["C9"] = "Month", "Revenue"
    db["F9"], db["G9"] = "Region", "Revenue"
    for c in ("B9", "C9", "F9", "G9"):
        db[c].font, db[c].fill = hdr_font, hdr_fill
    for k, m in enumerate(MONTHS):
        r = 10 + k
        db.cell(row=r, column=2, value=m).number_format = "mmm yyyy"
        db.cell(row=r, column=3, value=(f'=SUMIFS({rng("H")},{rng("B")},">="&B{r},'
                                        f'{rng("B")},"<"&EDATE(B{r},1))')).number_format = 'Â£#,##0'
    for k, g in enumerate(REGIONS):
        r = 10 + k
        db.cell(row=r, column=6, value=g)
        db.cell(row=r, column=7, value=f'=SUMIFS({rng("H")},{rng("C")},F{r})').number_format = 'Â£#,##0'
    for row in db.iter_rows(min_row=9, max_row=15, min_col=2, max_col=7):
        for c in row:
            if c.column in (2, 3, 6, 7) and (c.row <= 15 if c.column < 4 else c.row <= 13):
                c.border = Border(bottom=thin)

    ch = BarChart()
    ch.title, ch.y_axis.title, ch.legend = "Monthly revenue", "Â£", None
    ch.add_data(Reference(db, min_col=3, min_row=9, max_row=15), titles_from_data=True)
    ch.set_categories(Reference(db, min_col=2, min_row=10, max_row=15))
    ch.height, ch.width = 7.5, 15
    db.add_chart(ch, "B18")
    ch2 = BarChart()
    ch2.type, ch2.title, ch2.legend = "bar", "Revenue by region", None
    ch2.add_data(Reference(db, min_col=7, min_row=9, max_row=13), titles_from_data=True)
    ch2.set_categories(Reference(db, min_col=6, min_row=10, max_row=13))
    ch2.height, ch2.width = 7.5, 12
    db.add_chart(ch2, "I18")

    # --- How it works sheet
    hw = wb.create_sheet("How it works")
    lines = [
        "How this workbook works",
        "1. Add or paste new invoices at the bottom of the 'Invoices' table: it extends automatically.",
        "2. Amount and Days overdue are calculated for you (Qty x Unit price; days past the due date if unpaid).",
        "3. Change the Report date on the Dashboard to see the situation at any date.",
        "4. All dashboard figures and both charts update instantly: no macros, standard Excel formulas only.",
        "Sample data: fictional company, generated for demonstration.",
    ]
    for k, t in enumerate(lines, start=1):
        hw.cell(row=k, column=1, value=t).font = Font(bold=(k == 1), size=14 if k == 1 else 11)
    hw.column_dimensions["A"].width = 110

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def gbp(v):
    return f"Â£{v:,.0f}"


def build_png(agg, path):
    W, H = 1600, 900
    img = Image.new("RGB", (W, H), "#F7F7F5")
    d = ImageDraw.Draw(img)
    f_t, f_s = _font(BOLD, 42), _font(REG, 20)
    f_k, f_kv = _font(REG, 20), _font(BOLD, 38)
    d.text((60, 40), "Excel dashboard: invoices & cash collection", font=f_t, fill="#1B1B1B")
    d.text((60, 100), f"Live formulas, no macros Â· report date {AS_OF:%d/%m/%Y} Â· fictional data",
           font=f_s, fill="#666666")

    tiles = [("Total invoiced", gbp(agg["total"]), "#1F3A5F"), ("Paid", gbp(agg["paid"]), "#1F7A4D"),
             ("Outstanding", gbp(agg["outstanding"]), "#1F3A5F"), ("Overdue", gbp(agg["overdue"]), "#B23B3B"),
             ("Invoices", str(agg["count"]), "#1F3A5F")]
    tw = (W - 120 - 4 * 24) // 5
    for i, (lab, val, ink) in enumerate(tiles):
        x = 60 + i * (tw + 24)
        d.rounded_rectangle((x, 150, x + tw, 270), radius=14, fill="#FFFFFF", outline="#E2E2E2")
        d.text((x + 22, 168), lab, font=f_k, fill="#666666")
        d.text((x + 22, 200), val, font=f_kv, fill=ink)

    # Monthly bars
    x0, y0, cw, chh = 60, 310, 900, 520
    d.rounded_rectangle((x0, y0, x0 + cw, y0 + chh), radius=14, fill="#FFFFFF", outline="#E2E2E2")
    d.text((x0 + 24, y0 + 18), "Monthly revenue", font=_font(BOLD, 24), fill="#1B1B1B")
    mx = max(agg["monthly"])
    bw, base = 90, y0 + chh - 60
    for k, (m, v) in enumerate(zip(MONTHS, agg["monthly"])):
        bx = x0 + 60 + k * 135
        top = base - int((v / mx) * 340)
        d.rectangle((bx, top, bx + bw, base), fill="#1F3A5F")
        d.text((bx + bw / 2, top - 8), gbp(v), font=f_k, fill="#333333", anchor="mb")
        d.text((bx + bw / 2, base + 12), f"{m:%b}", font=f_k, fill="#666666", anchor="mt")

    # Region bars
    rx0, rw = 990, W - 60 - 990
    d.rounded_rectangle((rx0, y0, rx0 + rw, y0 + chh), radius=14, fill="#FFFFFF", outline="#E2E2E2")
    d.text((rx0 + 24, y0 + 18), "Revenue by region", font=_font(BOLD, 24), fill="#1B1B1B")
    rmx = max(agg["region"])
    for k, (g, v) in enumerate(zip(REGIONS, agg["region"])):
        by = y0 + 100 + k * 95
        d.text((rx0 + 24, by + 8), g, font=f_k, fill="#333333")
        length = int((v / rmx) * (rw - 260))
        d.rectangle((rx0 + 130, by, rx0 + 130 + length, by + 44), fill="#1F7A4D")
        d.text((rx0 + 140 + length, by + 22), gbp(v), font=f_k, fill="#333333", anchor="lm")
    img.save(path)


if __name__ == "__main__":
    data = make_rows()
    agg = aggregates(data)
    (ROOT / "output").mkdir(exist_ok=True)
    (ROOT / "docs").mkdir(exist_ok=True)
    build_xlsx(data, ROOT / "output" / "invoice-dashboard.xlsx")
    build_png(agg, ROOT / "docs" / "preview.png")
    print(f"Total {agg['total']:,} | Paid {agg['paid']:,} | Outstanding {agg['outstanding']:,}"
          f" | Overdue {agg['overdue']:,} | {agg['count']} invoices")

