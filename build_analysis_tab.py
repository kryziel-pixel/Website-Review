"""
Build / refresh the "Analysis" tab in the Creekside Unit Turn spreadsheet.

Layout:
  A1   — Title + timestamp
  A4   — Key Metrics KPIs
  A11  — Monthly Spend table  + chart (cols F–K)
  A26  — Quarterly Spend table + chart (cols F–K)
  A36  — Spend by Category table + pie chart (cols F–K)
  A58  — Top 10 Vendors table  + bar chart (cols F–K)
  A72  — Top 15 Units by Cost table + bar chart (cols F–K)
  A91  — Flagged Units (>$5k)
  A110 — Watch List ($3k–$5k)
  A125 — CFO Executive Summary (paragraph + bullets, one line per row)
"""

import os, sys, re, json, time
from datetime import datetime
from collections import defaultdict

sys.path.insert(0, '/home/user/Website-Review')
os.chdir('/home/user/Website-Review')

from creekside_automation import (
    get_gspread_client, SPREADSHEET_ID,
    UT_SHEET, PERQ_SHEET,
    norm_unit, norm_amt, to_date, quarter_of,
)

ANALYSIS_SHEET      = "Analysis"
SRC_SHEET           = "DATA Invoice By Location"
HIGH_COST_THRESHOLD = 5000
WATCH_THRESHOLD     = 3000

# GL Account → spend category mapping
GL_CATEGORY = {
    "5566": "Paint",
    "5217": "Paint",
    "5555": "Resurfacing",
    "7559": "Flooring",
    "5552": "Flooring",
    "5556": "Flooring",
    "7565": "Make Ready",
    "5214": "Make Ready",
    "5558": "HVAC",
    "5551": "Appliances",
    "7562": "Cabinets/Counters",
    "5559": "Cabinets/Counters",
    "6514": "Plumbing",
    "5564": "Plumbing",
    "5560": "Plumbing",
    "5316": "Water Mitigation",
    "5213": "Cleaning",
    "7570": "Cleaning",
    "5314": "Electrical",
    "5557": "Electrical",
    "5553": "Blinds/Window Treatments",
    "5563": "Doors",
    "5223": "Structural/Exterior",
    "6540": "Structural/Exterior",
    "6539": "Biohazard/Specialty",
    "5314": "Electrical",
    "6512": "Life Safety",
    "6520": "Insurance/Admin",
    "6522": "Insurance/Admin",
    "6505": "Hardware/Supplies",
    "5313": "Hardware/Supplies",
    "6503": "Hardware/Supplies",
    "6541": "Pest Control",
}

# ── colors ─────────────────────────────────────────────────────────────────
DARK_BLUE  = {"red": 0.122, "green": 0.220, "blue": 0.408}
MED_BLUE   = {"red": 0.235, "green": 0.420, "blue": 0.643}
LIGHT_BLUE = {"red": 0.812, "green": 0.886, "blue": 0.953}
AMBER      = {"red": 1.000, "green": 0.702, "blue": 0.000}
RED        = {"red": 0.800, "green": 0.000, "blue": 0.000}
GREEN      = {"red": 0.204, "green": 0.600, "blue": 0.200}
WHITE      = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
BLACK      = {"red": 0.0,   "green": 0.0,   "blue": 0.0}
LIGHT_GRAY = {"red": 0.950, "green": 0.950, "blue": 0.950}


def col_letter(n):
    r = ""
    while n:
        n, rem = divmod(n - 1, 26)
        r = chr(65 + rem) + r
    return r


# ── formatting request helpers ──────────────────────────────────────────────

def cell_fmt(sid, r, c, rows=1, cols=1, **kw):
    fmt = {}
    if "bg" in kw:
        fmt["backgroundColor"] = kw["bg"]
    tf = {}
    if kw.get("bold"):    tf["bold"] = True
    if "fg" in kw:        tf["foregroundColor"] = kw["fg"]
    if "size" in kw:      tf["fontSize"] = kw["size"]
    if kw.get("italic"):  tf["italic"] = True
    if tf:
        fmt["textFormat"] = tf
    if "halign" in kw:
        fmt["horizontalAlignment"] = kw["halign"]
    if "valign" in kw:
        fmt["verticalAlignment"] = kw["valign"]
    if "wrap" in kw:
        fmt["wrapStrategy"] = "WRAP" if kw["wrap"] else "OVERFLOW_CELL"
    fields_list = []
    if "bg" in kw:    fields_list.append("backgroundColor")
    if tf:            fields_list.append("textFormat")
    if "halign" in kw: fields_list.append("horizontalAlignment")
    if "valign" in kw:  fields_list.append("verticalAlignment")
    if "wrap" in kw:    fields_list.append("wrapStrategy")
    return {
        "repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r + rows,
                      "startColumnIndex": c, "endColumnIndex": c + cols},
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(" + ",".join(fields_list) + ")",
        }
    }


def merge_req(sid, r, c, rows=1, cols=1):
    return {"mergeCells": {
        "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r + rows,
                  "startColumnIndex": c, "endColumnIndex": c + cols},
        "mergeType": "MERGE_ALL",
    }}


def border_req(sid, r, c, rows=1, cols=1, color=None):
    color = color or {"red": 0.75, "green": 0.75, "blue": 0.75}
    b = {"style": "SOLID", "width": 1, "color": color}
    return {
        "updateBorders": {
            "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r + rows,
                      "startColumnIndex": c, "endColumnIndex": c + cols},
            "top": b, "bottom": b, "left": b, "right": b,
            "innerHorizontal": b, "innerVertical": b,
        }
    }


def number_fmt_req(sid, r, c, rows, cols, pattern):
    return {
        "repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r + rows,
                      "startColumnIndex": c, "endColumnIndex": c + cols},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": pattern}}},
            "fields": "userEnteredFormat.numberFormat",
        }
    }


def col_width_req(sid, start_col, end_col, pixels):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS",
                  "startIndex": start_col, "endIndex": end_col},
        "properties": {"pixelSize": pixels}, "fields": "pixelSize"}}


def row_height_req(sid, start_row, end_row, pixels):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "ROWS",
                  "startIndex": start_row, "endIndex": end_row},
        "properties": {"pixelSize": pixels}, "fields": "pixelSize"}}


def section_header(sid, row, label="", ncols=5):
    return [
        merge_req(sid, row, 0, 1, ncols),
        cell_fmt(sid, row, 0, 1, ncols, bg=MED_BLUE, fg=WHITE, bold=True,
                 size=10, halign="LEFT"),
    ]


# ── chart builders ──────────────────────────────────────────────────────────

def _source(sid, r_start, r_end, c_start, c_end):
    return {"sourceRange": {"sources": [{
        "sheetId": sid,
        "startRowIndex": r_start, "endRowIndex": r_end,
        "startColumnIndex": c_start, "endColumnIndex": c_end,
    }]}}


def column_chart(sid, title, data_start, n_rows, anchor_row, anchor_col,
                 w=460, h=260, x_title="", y_title="Spend ($)"):
    return {"addChart": {"chart": {
        "spec": {
            "title": title,
            "basicChart": {
                "chartType": "COLUMN",
                "legendPosition": "NO_LEGEND",
                "axis": [
                    {"position": "BOTTOM_AXIS", "title": x_title},
                    {"position": "LEFT_AXIS",   "title": y_title},
                ],
                "domains": [{"domain": _source(sid, data_start, data_start + n_rows, 0, 1)}],
                "series":  [{"series": _source(sid, data_start, data_start + n_rows, 1, 2),
                              "targetAxis": "LEFT_AXIS"}],
                "headerCount": 0,
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": sid, "rowIndex": anchor_row, "columnIndex": anchor_col},
            "widthPixels": w, "heightPixels": h,
        }},
    }}}


def bar_chart(sid, title, data_start, n_rows, anchor_row, anchor_col, w=460, h=300):
    return {"addChart": {"chart": {
        "spec": {
            "title": title,
            "basicChart": {
                "chartType": "BAR",
                "legendPosition": "NO_LEGEND",
                "axis": [
                    {"position": "BOTTOM_AXIS", "title": "Spend ($)"},
                    {"position": "LEFT_AXIS",   "title": ""},
                ],
                "domains": [{"domain": _source(sid, data_start, data_start + n_rows, 0, 1)}],
                "series":  [{"series": _source(sid, data_start, data_start + n_rows, 1, 2),
                              "targetAxis": "BOTTOM_AXIS"}],
                "headerCount": 0,
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": sid, "rowIndex": anchor_row, "columnIndex": anchor_col},
            "widthPixels": w, "heightPixels": h,
        }},
    }}}


def pie_chart(sid, title, data_start, n_rows, anchor_row, anchor_col, w=420, h=320):
    return {"addChart": {"chart": {
        "spec": {
            "title": title,
            "pieChart": {
                "legendPosition": "RIGHT_LEGEND",
                "domain": _source(sid, data_start, data_start + n_rows, 0, 1),
                "series": _source(sid, data_start, data_start + n_rows, 1, 2),
                "pieHole": 0.4,
            },
        },
        "position": {"overlayPosition": {
            "anchorCell": {"sheetId": sid, "rowIndex": anchor_row, "columnIndex": anchor_col},
            "widthPixels": w, "heightPixels": h,
        }},
    }}}


def delete_existing_charts(wb_gs, sid):
    sheet_data = wb_gs.fetch_sheet_metadata()
    for s in sheet_data["sheets"]:
        if s["properties"]["sheetId"] == sid:
            charts = s.get("charts", [])
            reqs = [{"deleteEmbeddedObject": {"objectId": c["chartId"]}} for c in charts]
            if reqs:
                wb_gs.batch_update({"requests": reqs})
            return


# ── data aggregation ────────────────────────────────────────────────────────

def aggregate_ut(ut_vals):
    """Aggregate from Unit Turn Cost by Unit sheet."""
    unit_totals  = defaultdict(float)
    unit_inv_cnt = defaultdict(int)
    monthly      = defaultdict(float)
    monthly_units = defaultdict(set)
    qtr          = defaultdict(float)
    qtr_units    = defaultdict(set)

    for row in ut_vals[1:]:
        unit  = norm_unit(row[0]) if len(row) > 0 else ""
        amt   = norm_amt(row[5]) if len(row) > 5 else 0
        inv_d = to_date(row[2]) if len(row) > 2 else None
        ins_d = to_date(row[3]) if len(row) > 3 else None
        d     = ins_d or inv_d
        if not unit:
            continue
        unit_totals[unit]  += amt
        unit_inv_cnt[unit] += 1
        if d:
            mk = f"{d.year}-{d.month:02d}"
            monthly[mk]        += amt
            monthly_units[mk].add(unit)
            q = quarter_of(d)
            qtr[q]             += amt
            qtr_units[q].add(unit)

    return unit_totals, unit_inv_cnt, monthly, monthly_units, qtr, qtr_units


def aggregate_src(src_vals):
    """Aggregate by category and vendor from DATA Invoice By Location."""
    cat_spend   = defaultdict(float)
    cat_cnt     = defaultdict(int)
    vendor_spend = defaultdict(float)
    vendor_cnt  = defaultdict(int)

    for row in src_vals[1:]:
        gl     = row[5].strip() if len(row) > 5 else ""
        amt    = norm_amt(row[8]) if len(row) > 8 else 0
        vendor = row[10].strip() if len(row) > 10 else ""
        cat    = GL_CATEGORY.get(gl, "Other")
        if amt:
            cat_spend[cat]   += amt
            cat_cnt[cat]     += 1
        if vendor and amt:
            vendor_spend[vendor] += amt
            vendor_cnt[vendor]   += 1

    return cat_spend, cat_cnt, vendor_spend, vendor_cnt


# ── CFO summary builder ─────────────────────────────────────────────────────

def build_cfo_summary(unit_totals, monthly, qtr, cat_spend, vendor_spend, now_label):
    total_spend  = sum(unit_totals.values())
    total_units  = len(unit_totals)
    avg_per_unit = total_spend / total_units if total_units else 0
    high_units   = sorted([(u, t) for u, t in unit_totals.items() if t >= HIGH_COST_THRESHOLD],
                          key=lambda x: -x[1])
    watch_units  = [(u, t) for u, t in unit_totals.items()
                    if WATCH_THRESHOLD <= t < HIGH_COST_THRESHOLD]

    sorted_months = sorted(monthly)
    last_mk       = sorted_months[-1] if sorted_months else ""
    last_spend    = monthly.get(last_mk, 0)
    prev_mk       = sorted_months[-2] if len(sorted_months) >= 2 else ""
    prev_spend    = monthly.get(prev_mk, 0)
    mom_pct       = ((last_spend - prev_spend) / prev_spend * 100) if prev_spend else 0
    mom_dir       = "higher" if mom_pct >= 0 else "lower"

    sorted_qtrs   = sorted(qtr)
    last_qtr      = sorted_qtrs[-1] if sorted_qtrs else ""
    last_qtr_spend = qtr.get(last_qtr, 0)
    prev_qtr      = sorted_qtrs[-2] if len(sorted_qtrs) >= 2 else ""
    prev_qtr_spend = qtr.get(prev_qtr, 0)
    qtr_pct       = ((last_qtr_spend - prev_qtr_spend) / prev_qtr_spend * 100) if prev_qtr_spend else 0
    qtr_dir       = "higher" if qtr_pct >= 0 else "lower"

    top3_units    = sorted(unit_totals.items(), key=lambda x: -x[1])[:3]
    top3_cats     = sorted(cat_spend.items(), key=lambda x: -x[1])[:3]
    top3_vendors  = sorted(vendor_spend.items(), key=lambda x: -x[1])[:3]
    top_cat       = top3_cats[0] if top3_cats else ("—", 0)
    top_cat_pct   = (top_cat[1] / total_spend * 100) if total_spend else 0

    # Try to determine month name
    try:
        last_month_label = datetime.strptime(last_mk, "%Y-%m").strftime("%B %Y")
        prev_month_label = datetime.strptime(prev_mk, "%Y-%m").strftime("%B %Y") if prev_mk else "prior month"
    except Exception:
        last_month_label = last_mk
        prev_month_label = prev_mk

    lines = []

    lines += [
        "OAKS AT CREEKSIDE — UNIT TURN COST ANALYSIS",
        f"Prepared for CFO Review  |  {now_label}",
        "",
        "──────────────────────────────────────────────────────────",
        "OVERVIEW",
        "──────────────────────────────────────────────────────────",
        (f"Since tracking began, Oaks at Creekside has recorded ${total_spend:,.0f} in unit-turn "
         f"expenses across {total_units} units, averaging ${avg_per_unit:,.0f} per unit. "
         f"The largest single spend category is {top_cat[0]} (${top_cat[1]:,.0f}, "
         f"{top_cat_pct:.0f}% of total spend), and the top vendor is "
         f"{top3_vendors[0][0]} (${top3_vendors[0][1]:,.0f})."),
        "",
        "──────────────────────────────────────────────────────────",
        "MOST RECENT MONTH — " + last_month_label.upper(),
        "──────────────────────────────────────────────────────────",
        (f"{last_month_label} spend was ${last_spend:,.0f}, which is {abs(mom_pct):.0f}% "
         f"{mom_dir} than {prev_month_label} (${prev_spend:,.0f}). "
         f"{'This increase reflects elevated unit-turn activity and may indicate a seasonal uptick or portfolio-wide vacancy.' if mom_pct > 20 else ('The decrease is a positive trend; monitor whether it is sustained next month.' if mom_pct < -10 else 'Month-over-month spend is relatively stable.')}"),
        "",
        "──────────────────────────────────────────────────────────",
        "QUARTERLY PERFORMANCE",
        "──────────────────────────────────────────────────────────",
        (f"{last_qtr} total spend was ${last_qtr_spend:,.0f}, "
         f"{abs(qtr_pct):.0f}% {qtr_dir} than {prev_qtr} (${prev_qtr_spend:,.0f}). "
         f"{'Quarter-over-quarter growth in spend warrants a review of turn volume and vendor pricing.' if qtr_pct > 15 else ('The reduction in quarterly spend is a favorable outcome.' if qtr_pct < -10 else 'Quarterly spend is trending within a normal range.')}"),
        "",
        "──────────────────────────────────────────────────────────",
        "SPEND BY CATEGORY (TOP 3)",
        "──────────────────────────────────────────────────────────",
    ]
    for cat, amt in top3_cats:
        pct = (amt / total_spend * 100) if total_spend else 0
        lines.append(f"  • {cat}: ${amt:,.0f}  ({pct:.1f}% of total)")

    lines += [
        "",
        "──────────────────────────────────────────────────────────",
        "TOP VENDORS",
        "──────────────────────────────────────────────────────────",
    ]
    for vendor, amt in top3_vendors:
        pct = (amt / total_spend * 100) if total_spend else 0
        lines.append(f"  • {vendor}: ${amt:,.0f}  ({pct:.1f}% of total spend)")
    lines.append(f"  The top 3 vendors together account for "
                 f"${sum(v[1] for v in top3_vendors):,.0f} "
                 f"({sum(v[1] for v in top3_vendors)/total_spend*100:.0f}% of all spend). "
                 f"Consider a vendor performance and pricing review annually.")

    lines += [
        "",
        "──────────────────────────────────────────────────────────",
        "FLAGS & ITEMS REQUIRING ATTENTION",
        "──────────────────────────────────────────────────────────",
    ]
    if high_units:
        lines.append(f"  🚨  {len(high_units)} units have exceeded ${HIGH_COST_THRESHOLD:,} in cumulative turn costs.")
        lines.append(f"      These units may be better candidates for capital improvements rather than")
        lines.append(f"      repeat repair-and-turn cycles. Recommend CapEx assessment for top units:")
        for u, t in high_units[:5]:
            lines.append(f"        - Unit {u}: ${t:,.0f} cumulative")
        if len(high_units) > 5:
            lines.append(f"        - (+ {len(high_units)-5} more units above ${HIGH_COST_THRESHOLD:,})")
    else:
        lines.append(f"  ✅  No units currently exceed ${HIGH_COST_THRESHOLD:,} cumulative threshold.")

    if watch_units:
        lines.append(f"")
        lines.append(f"  ⚠   {len(watch_units)} units are approaching the ${HIGH_COST_THRESHOLD:,} threshold (between ${WATCH_THRESHOLD:,}–${HIGH_COST_THRESHOLD:,}).")
        lines.append(f"      These should be monitored monthly. If turn costs continue to rise, escalate")
        lines.append(f"      to a capital plan discussion before the next lease renewal.")

    lines += [
        "",
        "──────────────────────────────────────────────────────────",
        "RECOMMENDATIONS",
        "──────────────────────────────────────────────────────────",
        f"  1. Schedule CapEx assessments for the top {min(5, len(high_units))} highest-cost units prior",
        f"     to their next lease renewal to evaluate full renovation vs. continued turns.",
        f"  2. Review vendor contracts — the top 3 vendors represent a significant concentration",
        f"     of spend. Verify competitive pricing on Paint and Flooring categories annually.",
        f"  3. Monitor the {len(watch_units)} watch-list units monthly; flag any that cross ${HIGH_COST_THRESHOLD:,}.",
        f"  4. Consider a unit condition audit for units with 10+ invoices to identify",
        f"     systemic issues (plumbing, HVAC, structural) driving recurring costs.",
        "",
    ]

    return lines


# ── main ────────────────────────────────────────────────────────────────────

def main():
    gc = get_gspread_client()
    wb = gc.open_by_key(SPREADSHEET_ID)

    ut_vals  = wb.worksheet(UT_SHEET).get_all_values()
    src_vals = wb.worksheet(SRC_SHEET).get_all_values()
    print(f"Read {len(ut_vals)-1} rows from {UT_SHEET}")
    print(f"Read {len(src_vals)-1} rows from {SRC_SHEET}")

    unit_totals, unit_inv_cnt, monthly, monthly_units, qtr, qtr_units = aggregate_ut(ut_vals)
    cat_spend, cat_cnt, vendor_spend, vendor_cnt = aggregate_src(src_vals)

    now_label = datetime.now().strftime("%B %d, %Y")

    # Get or create Analysis sheet
    existing = [s.title for s in wb.worksheets()]
    if ANALYSIS_SHEET in existing:
        ws = wb.worksheet(ANALYSIS_SHEET)
        print(f"Refreshing '{ANALYSIS_SHEET}' tab.")
    else:
        ws = wb.add_worksheet(title=ANALYSIS_SHEET, rows=300, cols=20)
        print(f"Created '{ANALYSIS_SHEET}' tab.")
    sid = ws.id
    ws.clear()
    time.sleep(1)
    # Unmerge all cells before rebuilding (avoids "must select merged range" errors)
    wb.batch_update({"requests": [{"unmergeCells": {
        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 400,
                  "startColumnIndex": 0, "endColumnIndex": 10}
    }}]})
    time.sleep(1)

    # ── Derived data ─────────────────────────────────────────────────────────
    total_spend   = sum(unit_totals.values())
    total_units   = len(unit_totals)
    avg_per_unit  = total_spend / total_units if total_units else 0
    high_units    = sorted([(u, t) for u, t in unit_totals.items() if t >= HIGH_COST_THRESHOLD],
                           key=lambda x: -x[1])
    watch_units   = sorted([(u, t) for u, t in unit_totals.items()
                             if WATCH_THRESHOLD <= t < HIGH_COST_THRESHOLD],
                           key=lambda x: -x[1])

    sorted_months = sorted(monthly)[-12:]
    sorted_qtrs   = sorted(qtr)
    top15_units   = sorted(unit_totals.items(), key=lambda x: -x[1])[:15]
    top10_vendors = sorted(vendor_spend.items(), key=lambda x: -x[1])[:10]
    cats_ranked   = sorted(cat_spend.items(), key=lambda x: -x[1])

    last_mk    = sorted_months[-1] if sorted_months else ""
    last_spend = monthly.get(last_mk, 0)
    prev_mk    = sorted_months[-2] if len(sorted_months) >= 2 else ""
    prev_spend = monthly.get(prev_mk, 0)
    last_qtr   = sorted_qtrs[-1] if sorted_qtrs else ""
    prev_qtr   = sorted_qtrs[-2] if len(sorted_qtrs) >= 2 else ""

    cfo_lines = build_cfo_summary(unit_totals, monthly, qtr, cat_spend, vendor_spend, now_label)

    # ── Row layout (0-based) ─────────────────────────────────────────────────
    R_TITLE      = 0
    R_UPDATED    = 1
    R_KPI_HEAD   = 3
    R_KPI_DATA   = 4    # 3 rows of 2-wide KPI pairs

    R_MON_HEAD   = 9
    R_MON_COL    = 10   # sub-header row for Monthly table
    R_MON_DATA   = 11
    n_mon        = len(sorted_months)

    R_QTR_HEAD   = R_MON_DATA + n_mon + 1
    R_QTR_COL    = R_QTR_HEAD + 1
    R_QTR_DATA   = R_QTR_HEAD + 2
    n_qtr        = len(sorted_qtrs)

    R_CAT_HEAD   = R_QTR_DATA + n_qtr + 1
    R_CAT_COL    = R_CAT_HEAD + 1
    R_CAT_DATA   = R_CAT_HEAD + 2
    n_cat        = len(cats_ranked)

    R_VEND_HEAD  = R_CAT_DATA + n_cat + 1
    R_VEND_COL   = R_VEND_HEAD + 1
    R_VEND_DATA  = R_VEND_HEAD + 2
    n_vend       = len(top10_vendors)

    R_TOP15_HEAD = R_VEND_DATA + n_vend + 1
    R_TOP15_COL  = R_TOP15_HEAD + 1
    R_TOP15_DATA = R_TOP15_HEAD + 2
    n_top15      = len(top15_units)

    R_FLAG_HEAD  = R_TOP15_DATA + n_top15 + 1
    R_FLAG_COL   = R_FLAG_HEAD + 1
    R_FLAG_DATA  = R_FLAG_HEAD + 2
    n_flag       = max(len(high_units), 1)

    R_WATCH_HEAD = R_FLAG_DATA + n_flag + 1
    R_WATCH_COL  = R_WATCH_HEAD + 1
    R_WATCH_DATA = R_WATCH_HEAD + 2
    n_watch      = max(len(watch_units), 1)

    R_SUM_HEAD   = R_WATCH_DATA + n_watch + 2
    R_SUM_DATA   = R_SUM_HEAD + 1
    n_sum        = len(cfo_lines)

    total_rows   = R_SUM_DATA + n_sum + 5

    # ── Build cell data ───────────────────────────────────────────────────────
    data = [[""] * 5 for _ in range(total_rows)]

    # Title
    data[R_TITLE][0] = "OAKS AT CREEKSIDE — UNIT TURN COST ANALYSIS"

    # Updated
    data[R_UPDATED][0] = f"Last updated: {now_label}    |    Data source: Resman → Google Sheets"

    # KPI header
    data[R_KPI_HEAD][0] = "KEY METRICS"

    kpi_pairs = [
        ("Total Spend (All-Time)",      f"${total_spend:,.0f}"),
        ("Units with Recorded Spend",   str(total_units)),
        ("Average Cost per Unit",       f"${avg_per_unit:,.0f}"),
        ("Units Exceeding $5,000",      f"{len(high_units)} units"),
        ("Units on Watch List ($3k–$5k)",f"{len(watch_units)} units"),
        (f"Most Recent Month ({last_mk})", f"${last_spend:,.0f}"),
    ]
    for i, (lbl, val) in enumerate(kpi_pairs):
        row = R_KPI_DATA + (i // 2)
        col = (i % 2) * 3   # pairs at col 0 and col 3
        data[row][col]     = lbl
        data[row][col + 1] = val

    # Monthly table
    data[R_MON_HEAD][0] = "MONTHLY SPEND (Last 12 Months)"
    data[R_MON_COL][0]  = "Month"
    data[R_MON_COL][1]  = "Total Spend"
    data[R_MON_COL][2]  = "# Units"
    for i, mk in enumerate(sorted_months):
        r = R_MON_DATA + i
        data[r][0] = mk
        data[r][1] = monthly[mk]
        data[r][2] = len(monthly_units[mk])

    # Quarterly table
    data[R_QTR_HEAD][0] = "QUARTERLY SPEND"
    data[R_QTR_COL][0]  = "Quarter"
    data[R_QTR_COL][1]  = "Total Spend"
    data[R_QTR_COL][2]  = "# Units"
    for i, q in enumerate(sorted_qtrs):
        r = R_QTR_DATA + i
        data[r][0] = q
        data[r][1] = qtr[q]
        data[r][2] = len(qtr_units[q])

    # Category table
    data[R_CAT_HEAD][0] = "SPEND BY CATEGORY"
    data[R_CAT_COL][0]  = "Category"
    data[R_CAT_COL][1]  = "Total Spend"
    data[R_CAT_COL][2]  = "% of Total"
    data[R_CAT_COL][3]  = "# Invoices"
    for i, (cat, amt) in enumerate(cats_ranked):
        r = R_CAT_DATA + i
        pct = (amt / total_spend * 100) if total_spend else 0
        data[r][0] = cat
        data[r][1] = amt
        data[r][2] = round(pct / 100, 4)   # as decimal for percentage format
        data[r][3] = cat_cnt[cat]
    # Total row
    r_cat_total = R_CAT_DATA + n_cat
    data[r_cat_total][0] = "Total"
    data[r_cat_total][1] = total_spend
    data[r_cat_total][2] = 1.0

    # Vendor table
    data[R_VEND_HEAD][0] = "TOP 10 VENDORS BY SPEND"
    data[R_VEND_COL][0]  = "Vendor"
    data[R_VEND_COL][1]  = "Total Spend"
    data[R_VEND_COL][2]  = "# Invoices"
    data[R_VEND_COL][3]  = "% of Total"
    for i, (v, amt) in enumerate(top10_vendors):
        r = R_VEND_DATA + i
        pct = (amt / total_spend * 100) if total_spend else 0
        data[r][0] = v
        data[r][1] = amt
        data[r][2] = vendor_cnt[v]
        data[r][3] = round(pct / 100, 4)

    # Top 15 units
    data[R_TOP15_HEAD][0] = "TOP 15 UNITS BY CUMULATIVE COST"
    data[R_TOP15_COL][0]  = "Unit"
    data[R_TOP15_COL][1]  = "Total Cost"
    data[R_TOP15_COL][2]  = "# Invoices"
    data[R_TOP15_COL][3]  = "Status"
    for i, (u, t) in enumerate(top15_units):
        r = R_TOP15_DATA + i
        flag = "🚨 HIGH" if t >= HIGH_COST_THRESHOLD else ("⚠ WATCH" if t >= WATCH_THRESHOLD else "OK")
        data[r][0] = f"Unit {u}"
        data[r][1] = t
        data[r][2] = unit_inv_cnt[u]
        data[r][3] = flag

    # Flagged units
    data[R_FLAG_HEAD][0] = "🚨  FLAGGED UNITS — CUMULATIVE COST > $5,000"
    data[R_FLAG_COL][0]  = "Unit"
    data[R_FLAG_COL][1]  = "Cumulative Cost"
    data[R_FLAG_COL][2]  = "# Invoices"
    data[R_FLAG_COL][3]  = "Action"
    if high_units:
        for i, (u, t) in enumerate(high_units):
            r = R_FLAG_DATA + i
            data[r][0] = f"Unit {u}"
            data[r][1] = t
            data[r][2] = unit_inv_cnt[u]
            data[r][3] = "CapEx Review"
    else:
        data[R_FLAG_DATA][0] = "No units flagged at this time."

    # Watch list
    data[R_WATCH_HEAD][0] = "⚠  WATCH LIST — APPROACHING $5,000 THRESHOLD ($3,000 – $4,999)"
    data[R_WATCH_COL][0]  = "Unit"
    data[R_WATCH_COL][1]  = "Cumulative Cost"
    data[R_WATCH_COL][2]  = "# Invoices"
    data[R_WATCH_COL][3]  = "Action"
    if watch_units:
        for i, (u, t) in enumerate(watch_units):
            r = R_WATCH_DATA + i
            data[r][0] = f"Unit {u}"
            data[r][1] = t
            data[r][2] = unit_inv_cnt[u]
            data[r][3] = "Monitor"
    else:
        data[R_WATCH_DATA][0] = "No units in watch range at this time."

    # CFO summary — one line per row, no merging
    data[R_SUM_HEAD][0] = "EXECUTIVE SUMMARY FOR CFO"
    for i, line in enumerate(cfo_lines):
        data[R_SUM_DATA + i][0] = line

    # ── Write to sheet ────────────────────────────────────────────────────────
    print(f"Writing {total_rows} rows to Analysis tab…")
    ws.update("A1", data, value_input_option="USER_ENTERED")
    time.sleep(3)

    # ── Formatting ────────────────────────────────────────────────────────────
    print("Applying formatting…")
    reqs = []

    # Title
    reqs.append(merge_req(sid, R_TITLE, 0, 1, 5))
    reqs.append(cell_fmt(sid, R_TITLE, 0, 1, 5, bg=DARK_BLUE, fg=WHITE, bold=True, size=14, halign="CENTER"))
    reqs.append(row_height_req(sid, R_TITLE, R_TITLE+1, 36))

    # Updated
    reqs.append(merge_req(sid, R_UPDATED, 0, 1, 5))
    reqs.append(cell_fmt(sid, R_UPDATED, 0, 1, 5, fg={"red":0.5,"green":0.5,"blue":0.5}, italic=True))

    # KPI header
    reqs.append(merge_req(sid, R_KPI_HEAD, 0, 1, 5))
    reqs.append(cell_fmt(sid, R_KPI_HEAD, 0, 1, 5, bg=MED_BLUE, fg=WHITE, bold=True, size=11, halign="CENTER"))

    # KPI rows
    for i in range(3):
        r = R_KPI_DATA + i
        reqs.append(cell_fmt(sid, r, 0, 1, 1, fg={"red":0.3,"green":0.3,"blue":0.3}))
        reqs.append(cell_fmt(sid, r, 1, 1, 1, bold=True, fg=DARK_BLUE, halign="RIGHT"))
        reqs.append(cell_fmt(sid, r, 3, 1, 1, fg={"red":0.3,"green":0.3,"blue":0.3}))
        reqs.append(cell_fmt(sid, r, 4, 1, 1, bold=True, fg=DARK_BLUE, halign="RIGHT"))

    def table_col_header(row, ncols=4):
        return [cell_fmt(sid, row, 0, 1, ncols, bg=DARK_BLUE, fg=WHITE, bold=True)]

    def zebra(row, n_data_rows, ncols=4, amt_col=1):
        rows_reqs = []
        for i in range(n_data_rows):
            r = row + i
            bg = LIGHT_BLUE if i % 2 == 0 else WHITE
            rows_reqs.append(cell_fmt(sid, r, 0, 1, ncols, bg=bg))
            rows_reqs.append(cell_fmt(sid, r, amt_col, 1, 1, halign="RIGHT"))
        return rows_reqs

    # Monthly
    reqs += section_header(sid, R_MON_HEAD)
    reqs += table_col_header(R_MON_COL, 3)
    reqs += zebra(R_MON_DATA, n_mon, ncols=3)
    reqs.append(border_req(sid, R_MON_COL, 0, n_mon + 1, 3))
    reqs.append(number_fmt_req(sid, R_MON_DATA, 1, n_mon, 1, '"$"#,##0'))

    # Quarterly
    reqs += section_header(sid, R_QTR_HEAD)
    reqs += table_col_header(R_QTR_COL, 3)
    reqs += zebra(R_QTR_DATA, n_qtr, ncols=3)
    reqs.append(border_req(sid, R_QTR_COL, 0, n_qtr + 1, 3))
    reqs.append(number_fmt_req(sid, R_QTR_DATA, 1, n_qtr, 1, '"$"#,##0'))

    # Category table
    reqs += section_header(sid, R_CAT_HEAD)
    reqs += table_col_header(R_CAT_COL, 4)
    reqs += zebra(R_CAT_DATA, n_cat, ncols=4)
    # Total row
    r_total = R_CAT_DATA + n_cat
    reqs.append(cell_fmt(sid, r_total, 0, 1, 4, bg=DARK_BLUE, fg=WHITE, bold=True))
    reqs.append(cell_fmt(sid, r_total, 1, 1, 1, halign="RIGHT"))
    reqs.append(border_req(sid, R_CAT_COL, 0, n_cat + 2, 4))
    reqs.append(number_fmt_req(sid, R_CAT_DATA, 1, n_cat + 1, 1, '"$"#,##0'))
    reqs.append(number_fmt_req(sid, R_CAT_DATA, 2, n_cat, 1, '0.0%'))

    # Vendor table
    reqs += section_header(sid, R_VEND_HEAD)
    reqs += table_col_header(R_VEND_COL, 4)
    reqs += zebra(R_VEND_DATA, n_vend, ncols=4)
    reqs.append(border_req(sid, R_VEND_COL, 0, n_vend + 1, 4))
    reqs.append(number_fmt_req(sid, R_VEND_DATA, 1, n_vend, 1, '"$"#,##0'))
    reqs.append(number_fmt_req(sid, R_VEND_DATA, 3, n_vend, 1, '0.0%'))

    # Top 15 units
    reqs += section_header(sid, R_TOP15_HEAD)
    reqs += table_col_header(R_TOP15_COL, 4)
    for i, (u, t) in enumerate(top15_units):
        r = R_TOP15_DATA + i
        is_high = t >= HIGH_COST_THRESHOLD
        is_watch = t >= WATCH_THRESHOLD
        bg = {"red": 1.0, "green": 0.88, "blue": 0.88} if is_high else (
             {"red": 1.0, "green": 0.97, "blue": 0.80} if is_watch else
             (LIGHT_BLUE if i % 2 == 0 else WHITE))
        reqs.append(cell_fmt(sid, r, 0, 1, 4, bg=bg))
        reqs.append(cell_fmt(sid, r, 1, 1, 1, halign="RIGHT"))
        reqs.append(cell_fmt(sid, r, 2, 1, 1, halign="CENTER"))
        if is_high:
            reqs.append(cell_fmt(sid, r, 3, 1, 1, fg=RED, bold=True))
        elif is_watch:
            reqs.append(cell_fmt(sid, r, 3, 1, 1, fg=AMBER, bold=True))
    reqs.append(border_req(sid, R_TOP15_COL, 0, n_top15 + 1, 4))
    reqs.append(number_fmt_req(sid, R_TOP15_DATA, 1, n_top15, 1, '"$"#,##0'))

    # Flagged
    reqs += section_header(sid, R_FLAG_HEAD)
    reqs += table_col_header(R_FLAG_COL, 4)
    for i in range(n_flag):
        r = R_FLAG_DATA + i
        reqs.append(cell_fmt(sid, r, 0, 1, 4, bg={"red": 1.0, "green": 0.88, "blue": 0.88}))
        reqs.append(cell_fmt(sid, r, 1, 1, 1, halign="RIGHT"))
        reqs.append(cell_fmt(sid, r, 3, 1, 1, fg=RED, bold=True))
    reqs.append(border_req(sid, R_FLAG_COL, 0, n_flag + 1, 4))
    if high_units:
        reqs.append(number_fmt_req(sid, R_FLAG_DATA, 1, n_flag, 1, '"$"#,##0'))

    # Watch list
    reqs += section_header(sid, R_WATCH_HEAD)
    reqs += table_col_header(R_WATCH_COL, 4)
    for i in range(n_watch):
        r = R_WATCH_DATA + i
        reqs.append(cell_fmt(sid, r, 0, 1, 4, bg={"red": 1.0, "green": 0.97, "blue": 0.80}))
        reqs.append(cell_fmt(sid, r, 1, 1, 1, halign="RIGHT"))
        reqs.append(cell_fmt(sid, r, 3, 1, 1, fg=AMBER, bold=True))
    reqs.append(border_req(sid, R_WATCH_COL, 0, n_watch + 1, 4))
    if watch_units:
        reqs.append(number_fmt_req(sid, R_WATCH_DATA, 1, n_watch, 1, '"$"#,##0'))

    # CFO Summary
    reqs.append(merge_req(sid, R_SUM_HEAD, 0, 1, 5))
    reqs.append(cell_fmt(sid, R_SUM_HEAD, 0, 1, 5, bg=DARK_BLUE, fg=WHITE, bold=True, size=12, halign="LEFT"))
    reqs.append(row_height_req(sid, R_SUM_HEAD, R_SUM_HEAD + 1, 30))

    for i, line in enumerate(cfo_lines):
        r = R_SUM_DATA + i
        # Section dividers
        if line.startswith("──"):
            reqs.append(cell_fmt(sid, r, 0, 1, 5, fg={"red":0.5,"green":0.5,"blue":0.5}, italic=True))
            reqs.append(row_height_req(sid, r, r+1, 8))
        elif line.isupper() and len(line) > 3 and not line.startswith("•") and not line.startswith(" "):
            # Sub-section header
            reqs.append(cell_fmt(sid, r, 0, 1, 5, bg=LIGHT_BLUE, bold=True, fg=DARK_BLUE))
            reqs.append(row_height_req(sid, r, r+1, 22))
        elif line.strip().startswith(("•", "🚨", "⚠", "✅", "1.", "2.", "3.", "4.")):
            reqs.append(cell_fmt(sid, r, 0, 1, 5, wrap=True))
        else:
            reqs.append(cell_fmt(sid, r, 0, 1, 5, wrap=True))

    # Column widths
    reqs += [
        col_width_req(sid, 0, 1, 210),   # A
        col_width_req(sid, 1, 2, 130),   # B
        col_width_req(sid, 2, 3, 100),   # C
        col_width_req(sid, 3, 4, 120),   # D
        col_width_req(sid, 4, 5, 60),    # E (spacer)
        col_width_req(sid, 5, 6, 210),   # F (chart anchor col)
    ]

    # Send formatting in chunks
    CHUNK = 40
    for i in range(0, len(reqs), CHUNK):
        wb.batch_update({"requests": reqs[i:i + CHUNK]})
        print(f"  Formatting {i//CHUNK + 1}/{(len(reqs)-1)//CHUNK + 1}")
        time.sleep(2)

    # ── Charts ────────────────────────────────────────────────────────────────
    print("Rebuilding charts…")
    delete_existing_charts(wb, sid)
    time.sleep(2)

    chart_reqs = [
        column_chart(sid, "Monthly Spend (Last 12 Months)",
                     R_MON_DATA, n_mon, R_MON_HEAD, 5),
        column_chart(sid, "Quarterly Spend",
                     R_QTR_DATA, n_qtr, R_QTR_HEAD, 5),
        pie_chart(sid, "Spend by Category",
                  R_CAT_DATA, n_cat, R_CAT_HEAD, 5, w=440, h=340),
        bar_chart(sid, "Top 10 Vendors by Spend",
                  R_VEND_DATA, n_vend, R_VEND_HEAD, 5, w=440, h=300),
        bar_chart(sid, "Top 15 Units by Cumulative Cost",
                  R_TOP15_DATA, n_top15, R_TOP15_HEAD, 5, w=440, h=360),
    ]
    wb.batch_update({"requests": chart_reqs})
    time.sleep(2)

    # ── Print summary ─────────────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("CFO SUMMARY")
    print("=" * 65)
    for line in cfo_lines:
        print(line)
    print("=" * 65)
    print(f"\n✅  Analysis tab complete  ({total_rows} rows, 5 charts)")


if __name__ == "__main__":
    main()
