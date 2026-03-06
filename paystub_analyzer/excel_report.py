import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from paystub_analyzer.config import OUTPUT_EXCEL
from paystub_analyzer.logger import get_logger

logger = get_logger(__name__)

# ── Colors ─────────────────────────────────────────────────────────────────────
DARK_BLUE   = "1F3864"
MID_BLUE    = "2E75B6"
LIGHT_BLUE  = "D6E4F0"
LIGHT_GRAY  = "F2F2F2"
WHITE       = "FFFFFF"
GREEN       = "1E8449"
LIGHT_GREEN = "D5F5E3"

HEADERS = ["Company", "Pay Period Start", "Pay Period End", "Gross Pay",
           "Net Pay", "Federal Tax", "Provincial Tax", "CPP", "EI",
           "Vacation Pay", "Hours Worked"]

# ── Styles ─────────────────────────────────────────────────────────────────────
def title_style(ws, cell_ref, text, size=14):
    c = ws[cell_ref]
    c.value = text
    c.font = Font(name="Arial", bold=True, size=size, color=WHITE)
    c.fill = PatternFill("solid", fgColor=DARK_BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")

def header_cell(cell, text):
    cell.value = text
    cell.font = Font(name="Arial", bold=True, size=10, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=MID_BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border()

def data_cell(cell, value, fmt=None, bold=False, bg=None):
    cell.value = value
    cell.font = Font(name="Arial", size=10, bold=bold)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border()
    if fmt:
        cell.number_format = fmt
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

# ── Load existing data ─────────────────────────────────────────────────────────
def load_existing_data() -> list:
    if not os.path.exists(OUTPUT_EXCEL):
        return []
    wb = openpyxl.load_workbook(OUTPUT_EXCEL)
    if "📋 Raw Data" not in wb.sheetnames:
        # Legacy single-sheet format
        ws = wb.active
        existing = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(row):
                existing.append(list(row))
        logger.info(f"📂 Loaded {len(existing)} existing records (legacy format)")
        return existing
    ws = wb["📋 Raw Data"]
    existing = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        if any(row):
            existing.append(list(row))
    logger.info(f"📂 Loaded {len(existing)} existing records")
    return existing

# ── Deduplicate ────────────────────────────────────────────────────────────────
def deduplicate(existing: list, new_data: list) -> list:
    keys = set()
    rows = []
    for row in existing:
        k = (row[0], row[2])
        if k not in keys:
            keys.add(k)
            rows.append(list(row))

    added = 0
    for d in new_data:
        k = (d.get("company"), d.get("pay_period_end"))
        if k not in keys:
            keys.add(k)
            rows.append([
                d.get("company"), d.get("pay_period_start"), d.get("pay_period_end"),
                d.get("gross_pay"), d.get("net_pay"), d.get("federal_tax"),
                d.get("provincial_tax"), d.get("cpp"), d.get("ei"),
                d.get("vacation_pay"), d.get("hours_worked")
            ])
            added += 1

    rows.sort(key=lambda x: x[1] or "")
    logger.info(f"✅ {added} new records added — {len(rows)} total")
    return rows

# ── Sheet 1: Raw Data ──────────────────────────────────────────────────────────
def build_raw_data(wb, rows):
    ws = wb.create_sheet("📋 Raw Data")
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A1:K1")
    title_style(ws, "A1", "📋  RAW DATA — ALL PAY PERIODS", size=13)

    for col, h in enumerate(HEADERS, 1):
        header_cell(ws.cell(row=3, column=col), h)

    for r, row in enumerate(rows, 4):
        bg = LIGHT_GRAY if r % 2 == 0 else WHITE
        for c, val in enumerate(row, 1):
            fmt = "$#,##0.00" if c in (4, 5, 6, 7, 8, 9, 10) else None
            data_cell(ws.cell(row=r, column=c), val, fmt=fmt, bg=bg)

    set_col_widths(ws, {"A": 36, "B": 18, "C": 18, "D": 14, "E": 12,
                        "F": 14, "G": 16, "H": 10, "I": 10, "J": 14, "K": 14})
    ws.freeze_panes = "A4"

# ── Sheet 2: Annual Summary ────────────────────────────────────────────────────
def build_annual_summary(wb, rows):
    ws = wb.create_sheet("📊 Annual Summary")
    ws.merge_cells("A1:H1")
    ws.row_dimensions[1].height = 30
    title_style(ws, "A1", "📊  ANNUAL SUMMARY — EARNINGS BY YEAR", size=13)

    hdrs = ["Year", "Pay Periods", "Gross Pay", "Net Pay",
            "Federal Tax", "Provincial Tax", "CPP", "EI"]
    for c, h in enumerate(hdrs, 1):
        header_cell(ws.cell(row=3, column=c), h)

    # Aggregate by year
    years = {}
    for row in rows:
        y = str(row[1])[:4] if row[1] else "Unknown"
        if y not in years:
            years[y] = [0] * 7
        years[y][0] += 1
        for i, idx in enumerate([3, 4, 5, 6, 7, 8]):
            try:
                years[y][i + 1] += float(row[idx] or 0)
            except (TypeError, ValueError):
                pass

    for r, (y, vals) in enumerate(sorted(years.items()), 4):
        bg = LIGHT_BLUE if r % 2 == 0 else WHITE
        data_cell(ws.cell(row=r, column=1), y, bg=bg)
        data_cell(ws.cell(row=r, column=2), vals[0], bg=bg)
        for c, v in enumerate(vals[1:], 3):
            data_cell(ws.cell(row=r, column=c), v, fmt="$#,##0.00", bg=bg)

    # Totals row
    tr = 4 + len(years)
    data_cell(ws.cell(row=tr, column=1), "TOTAL", bold=True, bg=LIGHT_GREEN)
    data_cell(ws.cell(row=tr, column=2), f"=SUM(B4:B{tr-1})", bold=True, bg=LIGHT_GREEN)
    for c in range(3, 9):
        col = get_column_letter(c)
        data_cell(ws.cell(row=tr, column=c),
                  f"=SUM({col}4:{col}{tr-1})", fmt="$#,##0.00", bold=True, bg=LIGHT_GREEN)

    set_col_widths(ws, {"A": 12, "B": 14, "C": 14, "D": 14,
                        "E": 14, "F": 16, "G": 12, "H": 12})

# ── Sheet 3: Monthly Summary ───────────────────────────────────────────────────
def build_monthly_summary(wb, rows):
    ws = wb.create_sheet("📅 Monthly Summary")
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 30
    title_style(ws, "A1", "📅  MONTHLY SUMMARY — EARNINGS OVER TIME", size=13)

    hdrs = ["Year", "Month", "Pay Periods", "Gross Pay", "Net Pay", "Total Tax", "Net Rate %"]
    for c, h in enumerate(hdrs, 1):
        header_cell(ws.cell(row=3, column=c), h)

    months_map = {"01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr",
                  "05": "May", "06": "Jun", "07": "Jul", "08": "Aug",
                  "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"}
    monthly = {}
    for row in rows:
        if not row[1]:
            continue
        key = str(row[1])[:7]
        y, m = key[:4], key[5:7]
        if key not in monthly:
            monthly[key] = {"year": y, "month": months_map.get(m, m), "periods": 0,
                            "gross": 0, "net": 0, "tax": 0}
        monthly[key]["periods"] += 1
        try:
            monthly[key]["gross"] += float(row[3] or 0)
            monthly[key]["net"] += float(row[4] or 0)
            monthly[key]["tax"] += float(row[5] or 0) + float(row[6] or 0)
        except (TypeError, ValueError):
            pass

    for r, (_, v) in enumerate(sorted(monthly.items()), 4):
        bg = LIGHT_GRAY if r % 2 == 0 else WHITE
        data_cell(ws.cell(row=r, column=1), v["year"], bg=bg)
        data_cell(ws.cell(row=r, column=2), v["month"], bg=bg)
        data_cell(ws.cell(row=r, column=3), v["periods"], bg=bg)
        data_cell(ws.cell(row=r, column=4), v["gross"], fmt="$#,##0.00", bg=bg)
        data_cell(ws.cell(row=r, column=5), v["net"], fmt="$#,##0.00", bg=bg)
        data_cell(ws.cell(row=r, column=6), v["tax"], fmt="$#,##0.00", bg=bg)
        rate = (v["net"] / v["gross"] * 100) if v["gross"] else 0
        data_cell(ws.cell(row=r, column=7), round(rate, 1), fmt="0.0%", bg=bg)

    set_col_widths(ws, {"A": 10, "B": 10, "C": 13, "D": 14, "E": 14, "F": 14, "G": 12})

# ── Sheet 4: By Company ────────────────────────────────────────────────────────
def build_by_company(wb, rows):
    ws = wb.create_sheet("🏢 By Company")
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 30
    title_style(ws, "A1", "🏢  EARNINGS BY COMPANY", size=13)

    hdrs = ["Company", "Pay Periods", "Gross Pay", "Net Pay",
            "Avg Gross/Period", "Avg Net/Period", "Avg Hours"]
    for c, h in enumerate(hdrs, 1):
        header_cell(ws.cell(row=3, column=c), h)

    companies = {}
    for row in rows:
        co = row[0] or "Unknown"
        if co not in companies:
            companies[co] = {"periods": 0, "gross": 0, "net": 0, "hours": 0}
        companies[co]["periods"] += 1
        try:
            companies[co]["gross"] += float(row[3] or 0)
            companies[co]["net"] += float(row[4] or 0)
            companies[co]["hours"] += float(row[10] or 0)
        except (TypeError, ValueError):
            pass

    for r, (co, v) in enumerate(sorted(companies.items()), 4):
        bg = LIGHT_BLUE if r % 2 == 0 else WHITE
        p = v["periods"]
        data_cell(ws.cell(row=r, column=1), co, bg=bg)
        data_cell(ws.cell(row=r, column=2), p, bg=bg)
        data_cell(ws.cell(row=r, column=3), v["gross"], fmt="$#,##0.00", bg=bg)
        data_cell(ws.cell(row=r, column=4), v["net"], fmt="$#,##0.00", bg=bg)
        data_cell(ws.cell(row=r, column=5), round(v["gross"] / p, 2) if p else 0, fmt="$#,##0.00", bg=bg)
        data_cell(ws.cell(row=r, column=6), round(v["net"] / p, 2) if p else 0, fmt="$#,##0.00", bg=bg)
        data_cell(ws.cell(row=r, column=7), round(v["hours"] / p, 1) if p else 0, bg=bg)

    set_col_widths(ws, {"A": 38, "B": 13, "C": 14, "D": 14, "E": 18, "F": 16, "G": 13})

# ── Sheet 5: Deductions ────────────────────────────────────────────────────────
def build_deductions(wb, rows):
    ws = wb.create_sheet("💰 Deductions")
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 30
    title_style(ws, "A1", "💰  DEDUCTIONS BREAKDOWN", size=13)

    hdrs = ["Pay Period End", "Company", "Gross Pay", "Federal Tax",
            "Provincial Tax", "CPP", "EI"]
    for c, h in enumerate(hdrs, 1):
        header_cell(ws.cell(row=3, column=c), h)

    for r, row in enumerate(rows, 4):
        bg = LIGHT_GRAY if r % 2 == 0 else WHITE
        data_cell(ws.cell(row=r, column=1), row[2], bg=bg)
        data_cell(ws.cell(row=r, column=2), row[0], bg=bg)
        for c, idx in enumerate([3, 5, 6, 7, 8], 3):
            data_cell(ws.cell(row=r, column=c), row[idx], fmt="$#,##0.00", bg=bg)

    # Totals
    tr = 4 + len(rows)
    data_cell(ws.cell(row=tr, column=1), "TOTAL", bold=True, bg=LIGHT_GREEN)
    data_cell(ws.cell(row=tr, column=2), "", bg=LIGHT_GREEN)
    for c in range(3, 8):
        col = get_column_letter(c)
        data_cell(ws.cell(row=tr, column=c),
                  f"=SUM({col}4:{col}{tr-1})", fmt="$#,##0.00", bold=True, bg=LIGHT_GREEN)

    set_col_widths(ws, {"A": 16, "B": 36, "C": 14, "D": 14, "E": 16, "F": 10, "G": 10})
    ws.freeze_panes = "A4"

# ── Sheet 6: Dashboard ─────────────────────────────────────────────────────────
def build_dashboard(wb, rows):
    ws = wb.create_sheet("🏠 Dashboard")
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 40
    title_style(ws, "A1", "🏠  PAYSTUB ANALYZER — EARNINGS DASHBOARD", size=14)

    total_gross = sum(float(r[3] or 0) for r in rows)
    total_net = sum(float(r[4] or 0) for r in rows)
    total_tax = sum(float(r[5] or 0) + float(r[6] or 0) for r in rows)
    total_cpp = sum(float(r[7] or 0) for r in rows)
    total_ei = sum(float(r[8] or 0) for r in rows)
    periods = len(rows)
    avg_gross = total_gross / periods if periods else 0
    avg_net = total_net / periods if periods else 0

    metrics = [
        ("Total Pay Periods", periods, None),
        ("Total Gross Pay", total_gross, "$#,##0.00"),
        ("Total Net Pay", total_net, "$#,##0.00"),
        ("Total Income Tax", total_tax, "$#,##0.00"),
        ("Total CPP", total_cpp, "$#,##0.00"),
        ("Total EI", total_ei, "$#,##0.00"),
        ("Avg Gross / Period", avg_gross, "$#,##0.00"),
        ("Avg Net / Period", avg_net, "$#,##0.00"),
    ]

    ws.cell(row=3, column=1).value = "METRIC"
    ws.cell(row=3, column=1).font = Font(name="Arial", bold=True, size=11)
    ws.cell(row=3, column=2).value = "VALUE"
    ws.cell(row=3, column=2).font = Font(name="Arial", bold=True, size=11)

    for r, (label, val, fmt) in enumerate(metrics, 4):
        bg = LIGHT_BLUE if r % 2 == 0 else WHITE
        lc = ws.cell(row=r, column=1)
        vc = ws.cell(row=r, column=2)
        lc.value = label
        lc.font = Font(name="Arial", size=11)
        lc.fill = PatternFill("solid", fgColor=bg)
        lc.border = thin_border()
        vc.value = val
        vc.font = Font(name="Arial", size=11, bold=True, color=GREEN)
        vc.fill = PatternFill("solid", fgColor=bg)
        vc.border = thin_border()
        if fmt:
            vc.number_format = fmt

    set_col_widths(ws, {"A": 25, "B": 18})

# ── Sheet 7: Glossary ──────────────────────────────────────────────────────────
def build_glossary(wb):
    ws = wb.create_sheet("📖 Glossary")
    ws.merge_cells("A1:D1")
    ws.row_dimensions[1].height = 30
    title_style(ws, "A1", "📖  GLOSSARY — CANADIAN PAYSTUB TERMS", size=13)

    hdrs = ["Term", "Full Name", "Description", "Who Pays"]
    for c, h in enumerate(hdrs, 1):
        header_cell(ws.cell(row=3, column=c), h)

    terms = [
        ("Gross Pay", "Gross Earnings", "Total earnings before any deductions", "N/A"),
        ("Net Pay", "Net Earnings / Take-Home Pay", "Amount received after all deductions", "N/A"),
        ("Federal Tax", "Federal Income Tax", "Tax paid to the Government of Canada", "Employee"),
        ("Provincial Tax", "Provincial Income Tax", "Tax paid to the provincial government", "Employee"),
        ("CPP", "Canada Pension Plan", "Retirement savings contribution", "Employee & Employer"),
        ("EI", "Employment Insurance", "Insurance for job loss or leave", "Employee & Employer"),
        ("Vacation Pay", "Vacation Pay Accrual", "Usually 4% of gross, paid out or accrued", "Employer"),
        ("Hours Worked", "Regular Hours", "Total hours worked in the pay period", "N/A"),
    ]

    for r, (term, full, desc, who) in enumerate(terms, 4):
        bg = LIGHT_GRAY if r % 2 == 0 else WHITE
        for c, val in enumerate([term, full, desc, who], 1):
            cell = ws.cell(row=r, column=c)
            cell.value = val
            cell.font = Font(name="Arial", size=10, bold=(c == 1))
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()
            cell.alignment = Alignment(wrap_text=True, vertical="center")

    set_col_widths(ws, {"A": 16, "B": 30, "C": 50, "D": 22})

# ── Main entry point ───────────────────────────────────────────────────────────
def create_excel(data_list: list) -> None:
    try:
        existing = load_existing_data()
        all_rows = deduplicate(existing, data_list)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # remove default sheet

        build_dashboard(wb, all_rows)
        build_raw_data(wb, all_rows)
        build_annual_summary(wb, all_rows)
        build_monthly_summary(wb, all_rows)
        build_by_company(wb, all_rows)
        build_deductions(wb, all_rows)
        build_glossary(wb)

        wb.save(OUTPUT_EXCEL)
        logger.info(f"💾 Excel saved at: {OUTPUT_EXCEL}")

    except Exception as e:
        logger.error(f"❌ Failed to create Excel: {e}")
        raise