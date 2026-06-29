"""
Build / refresh the "Analysis" tab in the Creekside Unit Turn spreadsheet.

Sections written:
  A1  — Header + last-updated timestamp
  A3  — Key Metrics (6 KPI cards)
  A11 — Monthly Spend Table (last 12 months)
  A26 — Quarterly Spend Table
  A36 — Cost by Unit – All Units (ranked)
  A55 — High-Cost Units (>$5k) watchlist
  A70 — CFO Summary (paragraph + bullets)

Charts added (or replaced):
  - Monthly Spend bar chart
  - Quarterly Spend column chart
  - Top 15 Units by Cost horizontal bar
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

ANALYSIS_SHEET = "Analysis"
HIGH_COST_THRESHOLD = 5000
WATCH_THRESHOLD = 3000

# ── colors ─────────────────────────────────────────────────────────────────
DARK_BLUE  = {"red": 0.122, "green": 0.220, "blue": 0.408}   # #1F3868 header
MED_BLUE   = {"red": 0.235, "green": 0.420, "blue": 0.643}   # #3C6BA4 sub-header
LIGHT_BLUE = {"red": 0.812, "green": 0.886, "blue": 0.953}   # #CFE2F3 row band
AMBER      = {"red": 1.000, "green": 0.702, "blue": 0.000}   # #FFB300 warning
RED        = {"red": 0.800, "green": 0.000, "blue": 0.000}   # #CC0000 alert
GREEN      = {"red": 0.204, "green": 0.600, "blue": 0.200}   # #339933 ok
WHITE      = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
BLACK      = {"red": 0.0,   "green": 0.0,   "blue": 0.0}
LIGHT_GRAY = {"red": 0.950, "green": 0.950, "blue": 0.950}


def col_letter(n):
    r = ""
    while n:
        n, rem = divmod(n - 1, 26)
        r = chr(65 + rem) + r
    return r


def cell_fmt(sid, r, c, rows=1, cols=1, **kw):
    """Helper — build a repeatCell request."""
    fmt = {}
    if "bg" in kw:
        fmt["backgroundColor"] = kw["bg"]
    if "bold" in kw or "fg" in kw or "size" in kw or "italic" in kw:
        tf = {}
        if kw.get("bold"):   tf["bold"] = True
        if "fg" in kw:       tf["foregroundColor"] = kw["fg"]
        if "size" in kw:     tf["fontSize"] = kw["size"]
        if kw.get("italic"): tf["italic"] = True
        fmt["textFormat"] = tf
    if "halign" in kw:
        fmt["horizontalAlignment"] = kw["halign"]
    if "valign" in kw:
        fmt["verticalAlignment"] = kw["valign"]
    if "wrap" in kw:
        fmt["wrapStrategy"] = "WRAP" if kw["wrap"] else "OVERFLOW_CELL"
    fields = ",".join(
        (["backgroundColor"] if "bg" in kw else []) +
        (["textFormat"] if any(k in kw for k in ("bold","fg","size","italic")) else []) +
        (["horizontalAlignment"] if "halign" in kw else []) +
        (["verticalAlignment"] if "valign" in kw else []) +
        (["wrapStrategy"] if "wrap" in kw else [])
    )
    return {
        "repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+rows,
                      "startColumnIndex": c, "endColumnIndex": c+cols},
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(" + fields + ")",
        }
    }


def merge_req(sid, r, c, rows=1, cols=1):
    return {"mergeCells": {
        "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+rows,
                  "startColumnIndex": c, "endColumnIndex": c+cols},
        "mergeType": "MERGE_ALL",
    }}


def border_req(sid, r, c, rows=1, cols=1, style="SOLID", width=1, color=None):
    color = color or {"red": 0.7, "green": 0.7, "blue": 0.7}
    b = {"style": style, "width": width, "color": color}
    return {
        "updateBorders": {
            "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+rows,
                      "startColumnIndex": c, "endColumnIndex": c+cols},
            "top": b, "bottom": b, "left": b, "right": b,
            "innerHorizontal": b, "innerVertical": b,
        }
    }


def number_fmt_req(sid, r, c, rows, cols, pattern):
    return {
        "repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": r, "endRowIndex": r+rows,
                      "startColumnIndex": c, "endColumnIndex": c+cols},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": pattern}}},
            "fields": "userEnteredFormat.numberFormat",
        }
    }


# ── data aggregation ────────────────────────────────────────────────────────

def aggregate(ut_vals):
    unit_totals   = defaultdict(float)
    unit_inv_cnt  = defaultdict(int)
    monthly       = defaultdict(float)
    monthly_units = defaultdict(set)
    qtr           = defaultdict(float)
    qtr_units     = defaultdict(set)
    flagged       = {}   # unit -> {"total", "invoices", "last_date"}

    for row in ut_vals[1:]:
        unit = norm_unit(row[0]) if len(row) > 0 else ""
        amt  = norm_amt(row[5]) if len(row) > 5 else 0
        inv_d  = to_date(row[2]) if len(row) > 2 else None
        inst_d = to_date(row[3]) if len(row) > 3 else None
        d = inst_d or inv_d
        if not unit:
            continue
        unit_totals[unit]  += amt
        unit_inv_cnt[unit] += 1
        if d:
            mk = f"{d.year}-{d.month:02d}"
            monthly[mk]       += amt
            monthly_units[mk].add(unit)
            q = quarter_of(d)
            qtr[q]            += amt
            qtr_units[q].add(unit)

    # Flag units
    for u, t in unit_totals.items():
        if t >= WATCH_THRESHOLD:
            flagged[u] = {"total": t, "invoices": unit_inv_cnt[u]}

    return unit_totals, unit_inv_cnt, monthly, monthly_units, qtr, qtr_units, flagged


def build_cfo_summary(unit_totals, monthly, qtr, flagged, now_label):
    total_spend  = sum(unit_totals.values())
    total_units  = len(unit_totals)
    avg_per_unit = total_spend / total_units if total_units else 0
    high_units   = [(u, t) for u, t in unit_totals.items() if t >= HIGH_COST_THRESHOLD]
    high_units.sort(key=lambda x: -x[1])

    sorted_months = sorted(monthly)
    last_month_key = sorted_months[-1] if sorted_months else ""
    last_month_spend = monthly.get(last_month_key, 0)
    prev_month_key = sorted_months[-2] if len(sorted_months) >= 2 else ""
    prev_month_spend = monthly.get(prev_month_key, 0)
    mom_pct = ((last_month_spend - prev_month_spend) / prev_month_spend * 100) if prev_month_spend else 0

    sorted_qtrs = sorted(qtr)
    last_qtr = sorted_qtrs[-1] if sorted_qtrs else ""
    last_qtr_spend = qtr.get(last_qtr, 0)

    top3 = sorted(unit_totals.items(), key=lambda x: -x[1])[:3]

    lines = [
        f"CREEKSIDE UNIT TURN — EXECUTIVE SUMMARY",
        f"Prepared: {now_label}",
        "",
        "OVERVIEW",
        (f"Since tracking began, Oaks at Creekside has recorded ${total_spend:,.0f} in unit-turn "
         f"expenses across {total_units} units, for an average of ${avg_per_unit:,.0f} per unit. "
         f"The most recent month ({last_month_key}) totaled ${last_month_spend:,.0f}, "
         f"{'up' if mom_pct >= 0 else 'down'} {abs(mom_pct):.0f}% vs. the prior month "
         f"(${prev_month_spend:,.0f}). {last_qtr} spend was ${last_qtr_spend:,.0f}."),
        "",
        "KEY METRICS",
        f"• Total spend (all-time): ${total_spend:,.0f}",
        f"• Units with recorded spend: {total_units}",
        f"• Average cost per unit: ${avg_per_unit:,.0f}",
        f"• Most recent month ({last_month_key}): ${last_month_spend:,.0f}",
        f"• Most recent quarter ({last_qtr}): ${last_qtr_spend:,.0f}",
        f"• Units exceeding ${HIGH_COST_THRESHOLD:,} (all-time): {len(high_units)}",
        "",
        "TOP 3 HIGHEST-COST UNITS",
    ]
    for u, t in top3:
        lines.append(f"• Unit {u}: ${t:,.0f} cumulative")

    lines += [
        "",
        "FLAGS & ITEMS TO MONITOR",
    ]
    if high_units:
        lines.append(f"• {len(high_units)} units have exceeded ${HIGH_COST_THRESHOLD:,} in cumulative "
                     f"turn costs. These units may warrant capital-improvement evaluation rather than "
                     f"continued repair-and-turn cycles.")
        lines.append(f"  Highest: Unit {high_units[0][0]} at ${high_units[0][1]:,.0f}")
    watch = [(u, d["total"]) for u, d in flagged.items() if d["total"] < HIGH_COST_THRESHOLD]
    watch.sort(key=lambda x: -x[1])
    if watch:
        lines.append(f"• {len(watch)} additional units are between ${WATCH_THRESHOLD:,}–${HIGH_COST_THRESHOLD:,} "
                     f"and should be monitored for cost escalation.")
    lines += [
        "",
        "RECOMMENDATION",
        ("Units surpassing $5,000 in turn costs should be reviewed for recurring maintenance patterns "
         "or deferred capital items. Consider scheduling CapEx assessments for the top 5 units prior "
         "to the next lease renewal cycle."),
    ]
    return lines


# ── chart builders ──────────────────────────────────────────────────────────

def delete_existing_charts(wb_gs, sid):
    """Remove all charts currently on the Analysis sheet."""
    sheet_data = wb_gs.fetch_sheet_metadata()
    for s in sheet_data["sheets"]:
        if s["properties"]["sheetId"] == sid:
            charts = s.get("charts", [])
            reqs = [{"deleteEmbeddedObject": {"objectId": c["chartId"]}} for c in charts]
            if reqs:
                wb_gs.batch_update({"requests": reqs})
            return


def monthly_bar_chart(sid, data_start_row, num_months):
    """Bar chart: Monthly Spend — data lives in cols A(label), B(spend) starting data_start_row."""
    return {
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Monthly Unit Turn Spend",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Month"},
                            {"position": "LEFT_AXIS",   "title": "Spend ($)"},
                        ],
                        "domains": [{"domain": {"sourceRange": {"sources": [{
                            "sheetId": sid,
                            "startRowIndex": data_start_row,
                            "endRowIndex":   data_start_row + num_months,
                            "startColumnIndex": 0,
                            "endColumnIndex":   1,
                        }]}}}],
                        "series": [{"series": {"sourceRange": {"sources": [{
                            "sheetId": sid,
                            "startRowIndex": data_start_row,
                            "endRowIndex":   data_start_row + num_months,
                            "startColumnIndex": 1,
                            "endColumnIndex":   2,
                        }]}}, "targetAxis": "LEFT_AXIS"}],
                        "headerCount": 0,
                    }
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": sid, "rowIndex": 11, "columnIndex": 5},
                        "widthPixels": 480,
                        "heightPixels": 280,
                    }
                },
            }
        }
    }


def quarterly_chart(sid, data_start_row, num_qtrs):
    return {
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Quarterly Unit Turn Spend",
                    "basicChart": {
                        "chartType": "BAR",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Spend ($)"},
                            {"position": "LEFT_AXIS",   "title": "Quarter"},
                        ],
                        "domains": [{"domain": {"sourceRange": {"sources": [{
                            "sheetId": sid,
                            "startRowIndex": data_start_row,
                            "endRowIndex":   data_start_row + num_qtrs,
                            "startColumnIndex": 0,
                            "endColumnIndex":   1,
                        }]}}}],
                        "series": [{"series": {"sourceRange": {"sources": [{
                            "sheetId": sid,
                            "startRowIndex": data_start_row,
                            "endRowIndex":   data_start_row + num_qtrs,
                            "startColumnIndex": 1,
                            "endColumnIndex":   2,
                        }]}}, "targetAxis": "BOTTOM_AXIS"}],
                        "headerCount": 0,
                    }
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": sid, "rowIndex": 26, "columnIndex": 5},
                        "widthPixels": 480,
                        "heightPixels": 280,
                    }
                },
            }
        }
    }


def top15_chart(sid, data_start_row, num_units):
    return {
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Top 15 Units by Cumulative Cost",
                    "basicChart": {
                        "chartType": "BAR",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Total Cost ($)"},
                            {"position": "LEFT_AXIS",   "title": "Unit"},
                        ],
                        "domains": [{"domain": {"sourceRange": {"sources": [{
                            "sheetId": sid,
                            "startRowIndex": data_start_row,
                            "endRowIndex":   data_start_row + num_units,
                            "startColumnIndex": 0,
                            "endColumnIndex":   1,
                        }]}}}],
                        "series": [{"series": {"sourceRange": {"sources": [{
                            "sheetId": sid,
                            "startRowIndex": data_start_row,
                            "endRowIndex":   data_start_row + num_units,
                            "startColumnIndex": 1,
                            "endColumnIndex":   2,
                        }]}}, "targetAxis": "BOTTOM_AXIS"}],
                        "headerCount": 0,
                    }
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": sid, "rowIndex": 36, "columnIndex": 5},
                        "widthPixels": 480,
                        "heightPixels": 380,
                    }
                },
            }
        }
    }


# ── main ────────────────────────────────────────────────────────────────────

def main():
    gc = get_gspread_client()
    wb = gc.open_by_key(SPREADSHEET_ID)

    # Read source data
    ut_vals = wb.worksheet(UT_SHEET).get_all_values()
    print(f"Read {len(ut_vals)-1} rows from {UT_SHEET}")

    unit_totals, unit_inv_cnt, monthly, monthly_units, qtr, qtr_units, flagged = aggregate(ut_vals)

    now_label = datetime.now().strftime("%B %d, %Y")

    # Get or create Analysis sheet
    existing = [s.title for s in wb.worksheets()]
    if ANALYSIS_SHEET in existing:
        ws = wb.worksheet(ANALYSIS_SHEET)
        print(f"Found existing '{ANALYSIS_SHEET}' sheet.")
    else:
        ws = wb.add_worksheet(title=ANALYSIS_SHEET, rows=200, cols=20)
        print(f"Created '{ANALYSIS_SHEET}' sheet.")
    sid = ws.id

    # Clear existing content
    ws.clear()
    time.sleep(1)

    # ── Build data arrays ────────────────────────────────────────────────────

    total_spend  = sum(unit_totals.values())
    total_units  = len(unit_totals)
    avg_per_unit = total_spend / total_units if total_units else 0
    high_units   = sorted([(u, t) for u, t in unit_totals.items() if t >= HIGH_COST_THRESHOLD], key=lambda x: -x[1])
    watch_units  = sorted([(u, t) for u, t in unit_totals.items() if WATCH_THRESHOLD <= t < HIGH_COST_THRESHOLD], key=lambda x: -x[1])

    sorted_months = sorted(monthly)[-12:]   # last 12 months
    sorted_qtrs   = sorted(qtr)
    top15         = sorted(unit_totals.items(), key=lambda x: -x[1])[:15]
    all_units_ranked = sorted(unit_totals.items(), key=lambda x: -x[1])

    # ROW LAYOUT (0-based for API, 1-based for display)
    ROW_TITLE       = 0   # row 1
    ROW_UPDATED     = 1   # row 2
    ROW_KPI_HEAD    = 3   # row 4  "Key Metrics"
    ROW_KPI_DATA    = 4   # row 5  KPI values (spans 4 rows)
    ROW_MON_HEAD    = 10  # row 11 Monthly header
    ROW_MON_DATA    = 11  # row 12 Monthly data start
    n_months        = len(sorted_months)
    ROW_QTR_HEAD    = ROW_MON_DATA + n_months + 1
    ROW_QTR_DATA    = ROW_QTR_HEAD + 1
    n_qtrs          = len(sorted_qtrs)
    ROW_TOP15_HEAD  = ROW_QTR_DATA + n_qtrs + 1
    ROW_TOP15_DATA  = ROW_TOP15_HEAD + 1
    ROW_FLAG_HEAD   = ROW_TOP15_DATA + 16
    ROW_FLAG_DATA   = ROW_FLAG_HEAD + 1
    n_flagged       = len(high_units)
    ROW_WATCH_HEAD  = ROW_FLAG_DATA + max(n_flagged, 1) + 1
    ROW_WATCH_DATA  = ROW_WATCH_HEAD + 1
    n_watch         = len(watch_units)
    ROW_SUMMARY     = ROW_WATCH_DATA + max(n_watch, 1) + 2

    # ── Write values ─────────────────────────────────────────────────────────
    data = []

    # Title
    data.append(["OAKS AT CREEKSIDE — UNIT TURN COST ANALYSIS"] + [""] * 4)

    # Updated
    data.append([f"Last updated: {now_label}"] + [""] * 4)

    # Blank
    data.append([""] * 5)

    # KPI section header
    data.append(["KEY METRICS"] + [""] * 4)

    # KPI rows (label, value, blank, label, value)
    kpi_pairs = [
        ("Total Spend (All-Time)",         f"${total_spend:,.0f}"),
        ("Units with Recorded Spend",      str(total_units)),
        ("Average Cost per Unit",          f"${avg_per_unit:,.0f}"),
        ("Units Exceeding $5,000",         str(len(high_units))),
        ("Units $3k–$5k (Watch List)",     str(len(watch_units))),
        ("Most Recent Month Spend",        f"${monthly.get(sorted_months[-1], 0):,.0f} ({sorted_months[-1]})"),
    ]
    for i in range(0, len(kpi_pairs), 2):
        l1, v1 = kpi_pairs[i]
        l2, v2 = kpi_pairs[i+1] if i+1 < len(kpi_pairs) else ("", "")
        data.append([l1, v1, "", l2, v2])

    # Blank rows to ROW_MON_HEAD
    while len(data) < ROW_MON_HEAD:
        data.append([""] * 5)

    # Monthly spend table
    data.append(["MONTHLY SPEND (Last 12 Months)", "", "", "", ""])
    for mk in sorted_months:
        data.append([mk, monthly[mk], len(monthly_units[mk]), "", ""])

    # Quarterly table
    while len(data) < ROW_QTR_HEAD:
        data.append([""] * 5)
    data.append(["QUARTERLY SPEND", "", "", "", ""])
    for q in sorted_qtrs:
        data.append([q, qtr[q], len(qtr_units[q]), "", ""])

    # Top 15 table
    while len(data) < ROW_TOP15_HEAD:
        data.append([""] * 5)
    data.append(["TOP 15 UNITS BY CUMULATIVE COST", "", "", "", ""])
    for rank, (u, t) in enumerate(top15, 1):
        flag = "🚨 HIGH" if t >= HIGH_COST_THRESHOLD else ("⚠ WATCH" if t >= WATCH_THRESHOLD else "")
        data.append([f"Unit {u}", t, unit_inv_cnt[u], flag, ""])

    # High-cost flags
    while len(data) < ROW_FLAG_HEAD:
        data.append([""] * 5)
    data.append(["🚨  FLAGGED UNITS (> $5,000 cumulative)", "", "", "", ""])
    if high_units:
        for u, t in high_units:
            data.append([f"Unit {u}", t, unit_inv_cnt[u], "REVIEW", ""])
    else:
        data.append(["No units flagged", "", "", "", ""])

    # Watch list
    while len(data) < ROW_WATCH_HEAD:
        data.append([""] * 5)
    data.append(["⚠  WATCH LIST ($3,000 – $5,000 cumulative)", "", "", "", ""])
    if watch_units:
        for u, t in watch_units:
            data.append([f"Unit {u}", t, unit_inv_cnt[u], "MONITOR", ""])
    else:
        data.append(["No units in watch range", "", "", "", ""])

    # CFO Summary
    while len(data) < ROW_SUMMARY:
        data.append([""] * 5)
    summary_lines = build_cfo_summary(unit_totals, monthly, qtr, flagged, now_label)
    for line in summary_lines:
        data.append([line] + [""] * 4)

    # Write everything
    print("Writing data to Analysis tab…")
    ws.update("A1", data, value_input_option="USER_ENTERED")
    time.sleep(3)

    # ── Formatting requests ──────────────────────────────────────────────────
    print("Applying formatting…")
    reqs = []

    # Title row
    reqs.append(merge_req(sid, ROW_TITLE, 0, 1, 5))
    reqs.append(cell_fmt(sid, ROW_TITLE, 0, 1, 5, bg=DARK_BLUE, fg=WHITE, bold=True, size=14, halign="CENTER"))

    # Updated row
    reqs.append(cell_fmt(sid, ROW_UPDATED, 0, 1, 5, fg={"red":0.5,"green":0.5,"blue":0.5}, italic=True))

    # KPI header
    reqs.append(merge_req(sid, ROW_KPI_HEAD, 0, 1, 5))
    reqs.append(cell_fmt(sid, ROW_KPI_HEAD, 0, 1, 5, bg=MED_BLUE, fg=WHITE, bold=True, size=11, halign="CENTER"))

    # KPI data rows
    for i in range(3):
        r = ROW_KPI_DATA + i
        reqs.append(cell_fmt(sid, r, 0, 1, 1, fg={"red":0.3,"green":0.3,"blue":0.3}))       # label col A
        reqs.append(cell_fmt(sid, r, 1, 1, 1, bold=True, fg=DARK_BLUE, halign="RIGHT"))       # value col B
        reqs.append(cell_fmt(sid, r, 3, 1, 1, fg={"red":0.3,"green":0.3,"blue":0.3}))       # label col D
        reqs.append(cell_fmt(sid, r, 4, 1, 1, bold=True, fg=DARK_BLUE, halign="RIGHT"))       # value col E

    def section_header_reqs(row, ncols=5):
        return [
            merge_req(sid, row, 0, 1, ncols),
            cell_fmt(sid, row, 0, 1, ncols, bg=MED_BLUE, fg=WHITE, bold=True, halign="LEFT"),
        ]

    def col_header_reqs(row, labels):
        reqs_out = []
        reqs_out.append(cell_fmt(sid, row, 0, 1, len(labels), bg=LIGHT_BLUE, bold=True))
        return reqs_out

    # Monthly section
    reqs += section_header_reqs(ROW_MON_HEAD)
    # Column sub-headers (Month | Spend | # Units)
    sub_row = ROW_MON_HEAD  # in the same row for now — data starts right below
    for i, (mk, amt) in enumerate(zip(sorted_months, [monthly[m] for m in sorted_months])):
        r = ROW_MON_DATA + i
        bg = LIGHT_BLUE if i % 2 == 0 else WHITE
        reqs.append(cell_fmt(sid, r, 0, 1, 1, bg=bg))
        reqs.append(cell_fmt(sid, r, 1, 1, 1, bg=bg, halign="RIGHT"))
        reqs.append(cell_fmt(sid, r, 2, 1, 1, bg=bg, halign="CENTER",
                             fg={"red":0.4,"green":0.4,"blue":0.4}))
    reqs.append(border_req(sid, ROW_MON_DATA, 0, n_months, 3))
    reqs.append(number_fmt_req(sid, ROW_MON_DATA, 1, n_months, 1, '"$"#,##0'))

    # Quarterly section
    reqs += section_header_reqs(ROW_QTR_HEAD)
    for i in range(n_qtrs):
        r = ROW_QTR_DATA + i
        bg = LIGHT_BLUE if i % 2 == 0 else WHITE
        reqs.append(cell_fmt(sid, r, 0, 1, 1, bg=bg))
        reqs.append(cell_fmt(sid, r, 1, 1, 1, bg=bg, halign="RIGHT"))
        reqs.append(cell_fmt(sid, r, 2, 1, 1, bg=bg, halign="CENTER",
                             fg={"red":0.4,"green":0.4,"blue":0.4}))
    reqs.append(border_req(sid, ROW_QTR_DATA, 0, n_qtrs, 3))
    reqs.append(number_fmt_req(sid, ROW_QTR_DATA, 1, n_qtrs, 1, '"$"#,##0'))

    # Top 15 section
    reqs += section_header_reqs(ROW_TOP15_HEAD)
    for i, (u, t) in enumerate(top15):
        r = ROW_TOP15_DATA + i
        is_high = t >= HIGH_COST_THRESHOLD
        is_watch = t >= WATCH_THRESHOLD
        bg = {"red": 1.0, "green": 0.9, "blue": 0.9} if is_high else (
             {"red": 1.0, "green": 0.97, "blue": 0.8} if is_watch else
             (LIGHT_BLUE if i % 2 == 0 else WHITE))
        reqs.append(cell_fmt(sid, r, 0, 1, 4, bg=bg))
        reqs.append(cell_fmt(sid, r, 1, 1, 1, halign="RIGHT"))
        reqs.append(cell_fmt(sid, r, 2, 1, 1, halign="CENTER"))
        if is_high:
            reqs.append(cell_fmt(sid, r, 3, 1, 1, fg=RED, bold=True))
        elif is_watch:
            reqs.append(cell_fmt(sid, r, 3, 1, 1, fg=AMBER, bold=True))
    reqs.append(border_req(sid, ROW_TOP15_DATA, 0, min(15, len(top15)), 4))
    reqs.append(number_fmt_req(sid, ROW_TOP15_DATA, 1, min(15, len(top15)), 1, '"$"#,##0'))

    # Flagged section
    reqs += section_header_reqs(ROW_FLAG_HEAD)
    if high_units:
        for i, (u, t) in enumerate(high_units):
            r = ROW_FLAG_DATA + i
            reqs.append(cell_fmt(sid, r, 0, 1, 4, bg={"red":1.0,"green":0.9,"blue":0.9}))
            reqs.append(cell_fmt(sid, r, 3, 1, 1, fg=RED, bold=True))
            reqs.append(cell_fmt(sid, r, 1, 1, 1, halign="RIGHT"))
        reqs.append(number_fmt_req(sid, ROW_FLAG_DATA, 1, n_flagged, 1, '"$"#,##0'))
        reqs.append(border_req(sid, ROW_FLAG_DATA, 0, n_flagged, 4))

    # Watch section
    reqs += section_header_reqs(ROW_WATCH_HEAD)
    if watch_units:
        for i, (u, t) in enumerate(watch_units):
            r = ROW_WATCH_DATA + i
            reqs.append(cell_fmt(sid, r, 0, 1, 4, bg={"red":1.0,"green":0.97,"blue":0.8}))
            reqs.append(cell_fmt(sid, r, 3, 1, 1, fg=AMBER, bold=True))
            reqs.append(cell_fmt(sid, r, 1, 1, 1, halign="RIGHT"))
        reqs.append(number_fmt_req(sid, ROW_WATCH_DATA, 1, n_watch, 1, '"$"#,##0'))
        reqs.append(border_req(sid, ROW_WATCH_DATA, 0, n_watch, 4))

    # Summary section
    reqs.append(merge_req(sid, ROW_SUMMARY, 0, 1, 5))
    reqs.append(cell_fmt(sid, ROW_SUMMARY, 0, 1, 5, bg=DARK_BLUE, fg=WHITE, bold=True, size=11, halign="LEFT"))
    n_sum = len(summary_lines)
    reqs.append(cell_fmt(sid, ROW_SUMMARY + 1, 0, n_sum, 5, wrap=True,
                         fg={"red":0.1,"green":0.1,"blue":0.1}))
    reqs.append(merge_req(sid, ROW_SUMMARY + 1, 0, n_sum, 5))
    # Shade alternating para blocks
    for i, line in enumerate(summary_lines):
        r = ROW_SUMMARY + 1 + i
        if line.isupper() and line:
            reqs.append(cell_fmt(sid, r, 0, 1, 5, bg=LIGHT_BLUE, bold=True))

    # Set column widths
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
        "properties": {"pixelSize": 200}, "fields": "pixelSize"}})
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2},
        "properties": {"pixelSize": 120}, "fields": "pixelSize"}})
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3},
        "properties": {"pixelSize": 90}, "fields": "pixelSize"}})
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4},
        "properties": {"pixelSize": 110}, "fields": "pixelSize"}})

    # Send formatting in chunks
    CHUNK = 40
    for start in range(0, len(reqs), CHUNK):
        wb.batch_update({"requests": reqs[start:start + CHUNK]})
        print(f"  Formatting batch {start//CHUNK + 1}/{(len(reqs)-1)//CHUNK + 1}")
        time.sleep(2)

    # ── Charts ───────────────────────────────────────────────────────────────
    print("Deleting old charts…")
    delete_existing_charts(wb, sid)
    time.sleep(2)

    print("Adding charts…")
    chart_reqs = [
        monthly_bar_chart(sid, ROW_MON_DATA, n_months),
        quarterly_chart(sid, ROW_QTR_DATA, n_qtrs),
        top15_chart(sid, ROW_TOP15_DATA, min(15, len(top15))),
    ]
    wb.batch_update({"requests": chart_reqs})
    time.sleep(2)

    # ── Print CFO summary to console ─────────────────────────────────────────
    print("\n" + "=" * 60)
    print("CFO SUMMARY (also written to Analysis tab)")
    print("=" * 60)
    for line in summary_lines:
        print(line)
    print("=" * 60)
    print(f"\n✅ Analysis tab built successfully.")


if __name__ == "__main__":
    main()
