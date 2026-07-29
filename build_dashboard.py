"""
Build / Refresh Creekside Dashboard
====================================
Reads live data from Google Sheets and regenerates creekside_dashboard.html
with embedded invoice rows + Unit Conditions meeting notes.

Usage:
    python build_dashboard.py
    python build_dashboard.py --out path/to/output.html
"""

import os, sys, re, json, argparse
from datetime import datetime
from collections import defaultdict

sys.path.insert(0, os.path.dirname(__file__))
from creekside_automation import (
    get_gspread_client, SPREADSHEET_ID,
    norm_unit, norm_amt, to_date,
)

SRC_SHEET   = "DATA Invoice By Location"
UT_SHEET    = "Unit Turn Cost by Unit"
COND_SHEET  = "Unit Conditions"
OUT_FILE    = os.path.join(os.path.dirname(__file__), "creekside_dashboard.html")

GL_CATEGORY = {
    "5566": "Paint",       "5217": "Paint",
    "5555": "Resurfacing",
    "7559": "Flooring",    "5552": "Flooring",    "5556": "Flooring",
    "7565": "Make Ready",  "5214": "Make Ready",
    "5558": "HVAC",
    "5551": "Appliances",
    "7562": "Cabinets/Counters", "5559": "Cabinets/Counters",
    "6514": "Plumbing",    "5564": "Plumbing",    "5560": "Plumbing",
    "5316": "Water Mitigation",
    "5213": "Cleaning",    "7570": "Cleaning",
    "5314": "Electrical",  "5557": "Electrical",
    "5553": "Blinds/Window Treatments",
    "5563": "Doors",
    "5223": "Structural/Exterior", "6540": "Structural/Exterior",
    "6539": "Biohazard/Specialty",
}

def quarter_of(d):
    q = (d.month - 1) // 3 + 1
    return f"{d.year}Q{q}"

def s(v):
    return str(v or "").strip()


def read_invoice_data(gc):
    """Read and join Unit Turn Cost by Unit with DATA Invoice by Location."""
    wb     = gc.open_by_key(SPREADSHEET_ID)
    ut_ws  = wb.worksheet(UT_SHEET)
    src_ws = wb.worksheet(SRC_SHEET)

    ut_vals  = ut_ws.get_all_values()
    src_vals = src_ws.get_all_values()

    # Build invoice → (vendor, gl, category) map from DATA Invoice by Location
    inv_meta = {}
    for row in src_vals[1:]:
        inv_num = s(row[3]) if len(row) > 3 else ""
        gl      = s(row[5]) if len(row) > 5 else ""
        vendor  = s(row[10]) if len(row) > 10 else ""
        cat     = GL_CATEGORY.get(gl, "Other")
        if inv_num:
            inv_meta[inv_num.upper()] = {"vendor": vendor, "category": cat, "gl": gl}

    rows = []
    cat_spend    = defaultdict(float)
    cat_cnt      = defaultdict(int)
    vendor_spend = defaultdict(float)
    vendor_cnt   = defaultdict(int)
    monthly      = defaultdict(float)
    qtr_data     = defaultdict(float)

    for row in ut_vals[1:]:
        unit  = norm_unit(row[0]) if len(row) > 0 else ""
        inv   = s(row[1]) if len(row) > 1 else ""
        inv_d = s(row[2]) if len(row) > 2 else ""
        ins_d = s(row[3]) if len(row) > 3 else ""
        desc  = s(row[4]) if len(row) > 4 else ""
        amt   = norm_amt(row[5]) if len(row) > 5 else 0
        if not unit or not amt:
            continue

        meta     = inv_meta.get(inv.upper(), {})
        vendor   = meta.get("vendor", "")
        category = meta.get("category", "Other")

        # Derive month/quarter from install date or invoice date
        d = to_date(ins_d) or to_date(inv_d)
        month  = f"{d.year}-{d.month:02d}" if d else ""
        quarter = quarter_of(d) if d else ""

        rows.append({
            "unit": unit, "invoice": inv,
            "inv_date": inv_d, "inst_date": ins_d,
            "desc": desc, "amount": amt,
            "month": month, "quarter": quarter,
            "vendor": vendor, "category": category,
        })

        if amt:
            cat_spend[category]  += amt
            cat_cnt[category]    += 1
        if vendor and amt:
            vendor_spend[vendor] += amt
            vendor_cnt[vendor]   += 1
        if month and amt:
            monthly[month] += amt
        if quarter and amt:
            qtr_data[quarter] += amt

    raw = {
        "rows": rows,
        "cat_spend":    dict(cat_spend),
        "cat_cnt":      dict(cat_cnt),
        "vendor_spend": dict(vendor_spend),
        "vendor_cnt":   dict(vendor_cnt),
        "monthly":      dict(monthly),
        "qtr":          dict(qtr_data),
        "generated":    datetime.now().strftime("%-m/%-d/%Y %-I:%M %p"),
    }
    return raw


def read_conditions(gc):
    """Read Unit Conditions sheet. Returns dict: unit -> {notes, make_ready_by, date_turned}."""
    wb = gc.open_by_key(SPREADSHEET_ID)
    ws = wb.worksheet(COND_SHEET)
    vals = ws.get_all_values()
    if not vals:
        return {}

    headers = [h.strip().lower() for h in vals[0]]

    def col(keyword):
        for i, h in enumerate(headers):
            if keyword in h:
                return i
        return -1

    unit_col     = col("unit")
    date_col     = col("date of unit")
    mready_col   = col("make ready")
    issue_col    = col("condition")
    if issue_col == -1:
        issue_col = col("issue")
    req_col      = col("required work")
    total_col    = col("total amount")

    result = {}
    for row in vals[1:]:
        unit = norm_unit(row[unit_col]) if unit_col >= 0 and unit_col < len(row) else ""
        if not unit:
            continue

        def get(c):
            return row[c].strip() if c >= 0 and c < len(row) else ""

        result[unit] = {
            "date_turned":   get(date_col),
            "make_ready_by": get(mready_col),
            "notes":         get(issue_col),
            "required_work": get(req_col),
            "total_spent":   get(total_col),
        }

    return result


def read_crosscheck(gc):
    """Read Cross-Check tab. Returns dict: unit -> {status, context, flags, ...}."""
    wb = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws   = wb.worksheet("Cross-Check")
        vals = ws.get_all_values()
    except Exception:
        return {}
    if len(vals) < 2:
        return {}

    headers = [h.strip().lower() for h in vals[0]]
    def col(kw):
        for i, h in enumerate(headers):
            if kw in h:
                return i
        return -1

    unit_col    = col("unit")
    status_col  = col("status")
    context_col = col("context")
    scope_col   = col("scope claim")
    notes_col   = col("notes summary")
    exp_col     = col("expected work")
    actual_col  = col("actual invoices")
    spend_col   = col("total spend")
    flags_col   = col("flags")

    result = {}
    for row in vals[1:]:
        def g(c): return row[c].strip() if c >= 0 and c < len(row) else ""
        unit = norm_unit(g(unit_col))
        if not unit:
            continue
        flags_raw = g(flags_col)
        result[unit] = {
            "status":        g(status_col),
            "context":       g(context_col),
            "scope_claim":   g(scope_col),
            "notes_summary": g(notes_col),
            "expected_work": g(exp_col),
            "actual_invoices": g(actual_col),
            "total_spend":   g(spend_col),
            "flags":         flags_raw,
            "flag_list":     [f.strip() for f in flags_raw.split("|") if f.strip() and f.strip() != "OK"],
        }
    return result


def build_html(raw_data, cond_data, cc_data, out_path):
    generated = raw_data["generated"]
    raw_json  = json.dumps(raw_data,  separators=(',', ':'))
    cond_json = json.dumps(cond_data, separators=(',', ':'))
    cc_json   = json.dumps(cc_data,   separators=(',', ':'))

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Oaks at Creekside — Unit Turn Analysis</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f0f4f8;color:#1a2540;font-size:14px}}
  .header{{background:linear-gradient(135deg,#1F3868 0%,#3C6BA4 100%);color:#fff;padding:18px 28px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.3)}}
  .header h1{{font-size:20px;font-weight:700;letter-spacing:.3px}}
  .header .sub{{font-size:12px;opacity:.75;margin-top:2px}}
  .header .updated{{font-size:11px;opacity:.6;text-align:right}}
  .nav{{background:#fff;border-bottom:2px solid #e2e8f0;display:flex;gap:0;padding:0 24px;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
  .nav button{{background:none;border:none;padding:14px 20px;cursor:pointer;font-size:13px;font-weight:600;color:#64748b;border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .2s}}
  .nav button:hover{{color:#3C6BA4}}
  .nav button.active{{color:#1F3868;border-bottom-color:#1F3868}}
  .page{{display:none;padding:24px;max-width:1400px;margin:0 auto}}
  .page.active{{display:block}}
  .kpi-row{{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:14px;margin-bottom:24px}}
  .kpi{{background:#fff;border-radius:10px;padding:18px 20px;box-shadow:0 1px 4px rgba(0,0,0,.08);border-left:4px solid #3C6BA4;cursor:default}}
  .kpi.red{{border-left-color:#CC0000}}
  .kpi.amber{{border-left-color:#F59E0B}}
  .kpi.green{{border-left-color:#22C55E}}
  .kpi .label{{font-size:11px;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}}
  .kpi .value{{font-size:26px;font-weight:800;color:#1F3868}}
  .kpi .sub{{font-size:11px;color:#94a3b8;margin-top:4px}}
  .grid-2{{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:20px}}
  .grid-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:18px;margin-bottom:20px}}
  .card{{background:#fff;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
  .card h3{{font-size:13px;font-weight:700;color:#1F3868;margin-bottom:14px;text-transform:uppercase;letter-spacing:.4px;padding-bottom:8px;border-bottom:2px solid #e2e8f0}}
  .chart-wrap{{position:relative;height:240px}}
  .chart-wrap.tall{{height:320px}}
  .chart-wrap.xtall{{height:420px}}
  .tbl-wrap{{overflow-x:auto;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.06)}}
  table{{width:100%;border-collapse:collapse;font-size:13px;background:#fff}}
  th{{background:#1F3868;color:#fff;padding:10px 12px;text-align:left;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px;position:sticky;top:0;z-index:2}}
  th.sort-asc::after{{content:' ↑'}}
  th.sort-desc::after{{content:' ↓'}}
  td{{padding:9px 12px;border-bottom:1px solid #f1f5f9}}
  tr:last-child td{{border-bottom:none}}
  tr:hover td{{background:#f8faff}}
  tr.clickable{{cursor:pointer}}
  tr.clickable:hover td{{background:#eff6ff}}
  .badge{{display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:700}}
  .badge.red{{background:#fee2e2;color:#CC0000}}
  .badge.amber{{background:#fef3c7;color:#b45309}}
  .badge.green{{background:#dcfce7;color:#166534}}
  .badge.blue{{background:#dbeafe;color:#1d4ed8}}
  .badge.purple{{background:#f3e8ff;color:#7c3aed}}
  .toolbar{{display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap;align-items:center}}
  .toolbar input,.toolbar select{{padding:8px 12px;border:1px solid #d1d5db;border-radius:7px;font-size:13px;background:#fff;outline:none;transition:border .15s}}
  .toolbar input:focus,.toolbar select:focus{{border-color:#3C6BA4;box-shadow:0 0 0 2px rgba(60,107,164,.15)}}
  .toolbar input{{min-width:220px}}
  .toolbar .count{{font-size:12px;color:#64748b;margin-left:4px}}
  .drill{{background:#fff;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-top:18px;border-top:3px solid #3C6BA4}}
  .drill h3{{font-size:14px;font-weight:700;color:#1F3868;margin-bottom:12px;display:flex;align-items:center;gap:10px}}
  .back-btn{{background:#eff6ff;border:1px solid #bfdbfe;color:#1d4ed8;padding:4px 12px;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600}}
  .back-btn:hover{{background:#dbeafe}}
  .hidden{{display:none}}
  .breadcrumb{{display:flex;align-items:center;gap:6px;font-size:12px;color:#64748b;margin-bottom:14px}}
  .breadcrumb span{{cursor:pointer;color:#3C6BA4;font-weight:600}}
  .breadcrumb span:hover{{text-decoration:underline}}
  .breadcrumb .sep{{color:#cbd5e1}}
  .summary-box{{background:linear-gradient(135deg,#1F3868,#2d5fa0);color:#fff;border-radius:10px;padding:22px 24px;margin-bottom:20px}}
  .summary-box h2{{font-size:15px;font-weight:700;margin-bottom:12px;opacity:.9;text-transform:uppercase;letter-spacing:.5px}}
  .summary-box p{{font-size:13px;line-height:1.7;opacity:.92;margin-bottom:8px}}
  .summary-box ul{{padding-left:18px;margin-top:6px}}
  .summary-box li{{font-size:13px;line-height:1.7;opacity:.9;margin-bottom:3px}}
  .summary-box .flag{{background:rgba(255,255,255,.15);border-left:3px solid #fbbf24;padding:8px 12px;border-radius:0 6px 6px 0;margin-top:10px;font-size:13px}}
  .summary-box .rec{{background:rgba(255,255,255,.1);border-left:3px solid #34d399;padding:8px 12px;border-radius:0 6px 6px 0;margin-top:6px;font-size:13px}}

  /* ── Cross-check panel ── */
  .cc-panel{{border-radius:8px;padding:14px 18px;margin-top:16px;border:1px solid #e2e8f0}}
  .cc-panel.cc-ok{{background:#f0fdf4;border-color:#bbf7d0}}
  .cc-panel.cc-warn{{background:#fffbeb;border-color:#fde68a}}
  .cc-panel.cc-alert{{background:#fff7ed;border-color:#fed7aa}}
  .cc-panel.cc-red{{background:#fef2f2;border-color:#fecaca}}
  .cc-hdr{{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}}
  .cc-hdr h4{{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.5px}}
  .cc-context{{font-size:11px;font-weight:600;padding:2px 10px;border-radius:10px;background:rgba(0,0,0,.07)}}
  .cc-flag{{font-size:12px;line-height:1.6;padding:6px 10px;margin-bottom:6px;border-radius:6px;border-left:3px solid}}
  .cc-flag.flag-warn{{background:#fef9c3;border-color:#ca8a04;color:#713f12}}
  .cc-flag.flag-alert{{background:#fff3e0;border-color:#f97316;color:#7c2d12}}
  .cc-flag.flag-red{{background:#fee2e2;border-color:#dc2626;color:#7f1d1d}}
  .cc-flag.flag-info{{background:#eff6ff;border-color:#3b82f6;color:#1e3a8a}}
  .cc-ok-msg{{font-size:12px;color:#16a34a;font-weight:600}}

  /* ── Meeting notes panel ── */
  .notes-panel{{background:#fafafa;border:1px solid #e2e8f0;border-radius:8px;padding:16px 18px;margin-top:16px}}
  .notes-panel .notes-hdr{{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}}
  .notes-panel .notes-hdr h4{{font-size:12px;font-weight:700;color:#1F3868;text-transform:uppercase;letter-spacing:.5px}}
  .notes-panel .notes-meta{{font-size:11px;color:#94a3b8}}
  .notes-entries{{max-height:260px;overflow-y:auto}}
  .note-entry{{border-left:3px solid #3C6BA4;padding:8px 12px;margin-bottom:8px;background:#fff;border-radius:0 6px 6px 0;font-size:12px;line-height:1.6}}
  .note-entry .note-date{{font-weight:700;color:#1F3868;font-size:11px;margin-bottom:4px;text-transform:uppercase;letter-spacing:.3px}}
  .note-entry .note-body{{color:#374151;white-space:pre-wrap}}
  .no-notes{{font-size:12px;color:#94a3b8;font-style:italic;padding:8px 0}}

  @media(max-width:768px){{
    .grid-2,.grid-3{{grid-template-columns:1fr}}
    .nav button{{padding:12px 12px;font-size:12px}}
    .header h1{{font-size:16px}}
  }}
</style>
</head>
<body>

<div class="header">
  <div>
    <h1>Oaks at Creekside — Unit Turn Analysis</h1>
    <div class="sub">Unit turn costs · vendor spend · unit conditions</div>
  </div>
  <div class="updated">Updated {generated}</div>
</div>

<nav class="nav">
  <button class="active" onclick="showPage('overview',this)">Overview</button>
  <button onclick="showPage('units',this)">Units</button>
  <button onclick="showPage('categories',this)">Categories</button>
  <button onclick="showPage('vendors',this)">Vendors</button>
  <button onclick="showPage('invoices',this)">All Invoices</button>
</nav>

<!-- ══════════════════════════════════════════════════════ OVERVIEW -->
<div id="page-overview" class="page active">
  <div class="kpi-row" id="overview-kpis"></div>
  <div class="summary-box" id="exec-summary"></div>
  <div class="grid-2">
    <div class="card"><h3>Monthly Spend</h3><div class="chart-wrap tall"><canvas id="chart-monthly"></canvas></div></div>
    <div class="card"><h3>Quarterly Spend</h3><div class="chart-wrap tall"><canvas id="chart-qtr"></canvas></div></div>
  </div>
  <div class="grid-2">
    <div class="card"><h3>Spend by Category</h3><div class="chart-wrap tall"><canvas id="chart-cat-donut"></canvas></div></div>
    <div class="card"><h3>Top 10 Vendors</h3><div class="chart-wrap tall"><canvas id="chart-vendor-bar"></canvas></div></div>
  </div>
</div>

<!-- ══════════════════════════════════════════════════════ UNITS -->
<div id="page-units" class="page">
  <div class="toolbar">
    <input id="unit-search" placeholder="🔍 Search unit # or description…" oninput="filterUnits()">
    <select id="unit-filter" onchange="filterUnits()">
      <option value="">All statuses</option>
      <option value="high">🚨 High (&gt;$5K)</option>
      <option value="watch">⚠ Watch ($3K–$5K)</option>
      <option value="ok">✅ OK (&lt;$3K)</option>
    </select>
    <span class="count" id="unit-count"></span>
  </div>
  <div class="tbl-wrap">
    <table id="unit-table">
      <thead><tr>
        <th onclick="sortUnits('unit')">Unit</th>
        <th onclick="sortUnits('total')">Total Spend</th>
        <th onclick="sortUnits('invoices')"># Invoices</th>
        <th>Top Category</th>
        <th>Last Invoice</th>
        <th>Status</th>
        <th>Conditions</th>
        <th>Cross-Check</th>
      </tr></thead>
      <tbody id="unit-tbody"></tbody>
    </table>
  </div>

  <!-- Unit drill-down -->
  <div class="drill hidden" id="unit-drill">
    <h3>
      <button class="back-btn" onclick="closeDrill('unit-drill')">← Back</button>
      <span id="unit-drill-title"></span>
    </h3>
    <div class="kpi-row" id="unit-drill-kpis"></div>
    <div class="grid-2">
      <div class="card">
        <h3>Spend by Category</h3>
        <div class="chart-wrap"><canvas id="unit-cat-chart"></canvas></div>
      </div>
      <div class="card">
        <h3>Spend Over Time</h3>
        <div class="chart-wrap"><canvas id="unit-time-chart"></canvas></div>
      </div>
    </div>
    <!-- Cross-check findings -->
    <div class="cc-panel hidden" id="unit-cc-panel">
      <div class="cc-hdr">
        <h4 id="unit-cc-title"></h4>
        <span class="cc-context" id="unit-cc-context"></span>
      </div>
      <div id="unit-cc-flags"></div>
    </div>
    <!-- Meeting notes -->
    <div class="notes-panel" id="unit-notes-panel">
      <div class="notes-hdr">
        <h4>📋 Meeting Notes — Unit Conditions</h4>
        <span class="notes-meta" id="unit-notes-meta"></span>
      </div>
      <div class="notes-entries" id="unit-notes-entries"></div>
    </div>
    <div class="card" style="margin-top:16px">
      <h3>All Invoices</h3>
      <div class="tbl-wrap" style="max-height:340px;overflow-y:auto">
        <table><thead><tr>
          <th>Invoice #</th><th>Invoice Date</th><th>Install Date</th>
          <th>Description</th><th>Category</th><th>Vendor</th><th>Amount</th>
        </tr></thead>
        <tbody id="unit-invoices-tbody"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- ══════════════════════════════════════════════════════ CATEGORIES -->
<div id="page-categories" class="page">
  <div class="grid-2" style="margin-bottom:20px">
    <div class="card"><h3>Spend by Category</h3><div class="chart-wrap xtall"><canvas id="chart-cat-bar"></canvas></div></div>
    <div class="card"><h3>Category Breakdown</h3><div class="chart-wrap xtall"><canvas id="chart-cat-pie2"></canvas></div></div>
  </div>
  <div class="toolbar"><span class="count">Click any row to drill into units for that category</span></div>
  <div class="tbl-wrap">
    <table id="cat-table">
      <thead><tr>
        <th>Category</th><th>Total Spend</th><th>% of Total</th><th># Invoices</th><th>Avg / Invoice</th>
      </tr></thead>
      <tbody id="cat-tbody"></tbody>
    </table>
  </div>
  <div class="drill hidden" id="cat-drill">
    <h3>
      <button class="back-btn" onclick="closeDrill('cat-drill')">← Back</button>
      <span id="cat-drill-title"></span>
    </h3>
    <div class="toolbar">
      <input id="cat-drill-search" placeholder="🔍 Filter units…" oninput="filterCatDrill()">
    </div>
    <div class="tbl-wrap">
      <table><thead><tr><th>Unit</th><th>Spend</th><th># Invoices</th><th>Status</th></tr></thead>
      <tbody id="cat-drill-tbody"></tbody></table>
    </div>
  </div>
</div>

<!-- ══════════════════════════════════════════════════════ VENDORS -->
<div id="page-vendors" class="page">
  <div class="grid-2" style="margin-bottom:20px">
    <div class="card"><h3>Top 15 Vendors by Spend</h3><div class="chart-wrap xtall"><canvas id="chart-vendor-bar2"></canvas></div></div>
    <div class="card">
      <h3>Vendor Summary</h3>
      <div class="toolbar" style="margin-bottom:10px">
        <input id="vendor-search" placeholder="🔍 Search vendors…" oninput="filterVendors()">
      </div>
      <div class="tbl-wrap" style="max-height:380px;overflow-y:auto">
        <table><thead><tr><th>#</th><th>Vendor</th><th>Total</th><th># Invoices</th><th>Avg</th></tr></thead>
        <tbody id="vendor-tbody"></tbody></table>
      </div>
    </div>
  </div>
  <div class="drill hidden" id="vendor-drill">
    <h3>
      <button class="back-btn" onclick="closeDrill('vendor-drill')">← Back</button>
      <span id="vendor-drill-title"></span>
    </h3>
    <div class="grid-2" style="margin-bottom:14px">
      <div class="card"><h3>Units Served</h3><div class="chart-wrap"><canvas id="vendor-unit-chart"></canvas></div></div>
      <div class="card"><h3>Category Mix</h3><div class="chart-wrap"><canvas id="vendor-cat-chart"></canvas></div></div>
    </div>
    <div class="tbl-wrap">
      <table><thead><tr><th>Unit</th><th>Invoice #</th><th>Date</th><th>Description</th><th>Category</th><th>Amount</th></tr></thead>
      <tbody id="vendor-drill-tbody"></tbody></table>
    </div>
  </div>
</div>

<!-- ══════════════════════════════════════════════════════ ALL INVOICES -->
<div id="page-invoices" class="page">
  <div class="toolbar">
    <input id="inv-search" placeholder="🔍 Search description, vendor, unit…" oninput="filterInvoices()">
    <select id="inv-cat" onchange="filterInvoices()"><option value="">All categories</option></select>
    <select id="inv-month" onchange="filterInvoices()"><option value="">All months</option></select>
    <select id="inv-qtr" onchange="filterInvoices()"><option value="">All quarters</option></select>
    <span class="count" id="inv-count"></span>
  </div>
  <div class="tbl-wrap" style="max-height:70vh;overflow-y:auto">
    <table id="inv-table">
      <thead><tr>
        <th onclick="sortInvoices('unit')">Unit</th>
        <th onclick="sortInvoices('invoice')">Invoice #</th>
        <th onclick="sortInvoices('inv_date')">Invoice Date</th>
        <th onclick="sortInvoices('inst_date')">Install Date</th>
        <th onclick="sortInvoices('desc')">Description</th>
        <th onclick="sortInvoices('category')">Category</th>
        <th onclick="sortInvoices('vendor')">Vendor</th>
        <th onclick="sortInvoices('amount')">Amount</th>
      </tr></thead>
      <tbody id="inv-tbody"></tbody>
    </table>
  </div>
  <div style="padding:10px 0;font-size:13px;color:#64748b" id="inv-total-row"></div>
</div>

<script>
// ══════════════════════════════════════════════════════════════════
// DATA
// ══════════════════════════════════════════════════════════════════
const RAW  = {raw_json};
const COND = {cond_json};
const CC   = {cc_json};

const fmt  = n => '$' + Math.round(n).toLocaleString();
const fmtD = n => '$' + n.toFixed(2).replace(/\\B(?=(\\d{{3}})+(?!\\d))/g,',');
const pct  = (n,t) => t ? (n/t*100).toFixed(1)+'%' : '0%';

// Pre-compute aggregates
const rows = RAW.rows;
const totalSpend   = rows.reduce((s,r)=>s+r.amount,0);
const unitTotals   = {{}};
const unitRows     = {{}};
rows.forEach(r=>{{
  unitTotals[r.unit] = (unitTotals[r.unit]||0)+r.amount;
  (unitRows[r.unit]=unitRows[r.unit]||[]).push(r);
}});
const vendorTotals = RAW.vendor_spend;
const vendorCnt    = RAW.vendor_cnt;
const catSpend     = RAW.cat_spend;
const catCnt       = RAW.cat_cnt;
const monthly      = RAW.monthly;
const qtrData      = RAW.qtr;

const sortedMonths = Object.keys(monthly).sort().slice(-12);
const sortedQtrs   = Object.keys(qtrData).sort();

const HIGH = 5000, WATCH = 3000;
const unitStatus = u => unitTotals[u]>=HIGH?'high': unitTotals[u]>=WATCH?'watch':'ok';
const statusBadge = s => s==='high'?'<span class="badge red">🚨 High</span>':
                         s==='watch'?'<span class="badge amber">⚠ Watch</span>':
                         '<span class="badge green">✅ OK</span>';

const PALETTE = ['#3C6BA4','#F59E0B','#22C55E','#EF4444','#8B5CF6',
                 '#06B6D4','#F97316','#10B981','#EC4899','#6366F1',
                 '#84CC16','#14B8A6','#F43F5E','#A855F7','#3B82F6',
                 '#D97706','#059669','#DC2626'];

// ── Page switching ──────────────────────────────────────────────
function showPage(id, btn){{
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav button').forEach(b=>b.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  btn.classList.add('active');
}}

// ── Drill-down helpers ──────────────────────────────────────────
const drillCharts = {{}};
function closeDrill(id){{
  document.getElementById(id).classList.add('hidden');
  if(drillCharts[id]){{drillCharts[id].forEach(c=>c.destroy());drillCharts[id]=[];}}
}}
function makeChart(id, type, labels, data, opts={{}}){{
  const ctx = document.getElementById(id);
  if(!ctx) return null;
  if(ctx._chart) ctx._chart.destroy();
  const c = new Chart(ctx, {{type, data:{{
    labels,
    datasets:[{{data, backgroundColor: opts.colors||PALETTE,
               borderColor: type==='line'?'#3C6BA4':'transparent',
               fill: opts.fill||false,
               tension:.3,
               borderWidth:2}}]
  }}, options:{{
    responsive:true, maintainAspectRatio:false,
    plugins:{{
      legend:{{display:opts.legend!==false,position:opts.legendPos||'top',
              labels:{{font:{{size:11}},boxWidth:14}}}},
      tooltip:{{callbacks:{{label:ctx2=>{{
        const v=ctx2.parsed.y??ctx2.parsed;
        return typeof v==='number'?' '+fmt(v):' '+v;
      }}}}}}
    }},
    scales: type==='bar'||type==='line' ? {{
      x:{{ticks:{{font:{{size:11}}}}}},
      y:{{ticks:{{callback:v=>fmt(v),font:{{size:11}}}},beginAtZero:true}}
    }} : {{}},
    ...(opts.extra||{{}})
  }}}});
  ctx._chart = c;
  return c;
}}

// ══════════════════════════════════════════════════════════════════
// OVERVIEW
// ══════════════════════════════════════════════════════════════════
function buildOverview(){{
  const totalUnits = Object.keys(unitTotals).length;
  const highUnits  = Object.values(unitTotals).filter(v=>v>=HIGH).length;
  const lastMo     = sortedMonths[sortedMonths.length-1]||'';
  const lastMoAmt  = monthly[lastMo]||0;
  const prevMoAmt  = monthly[sortedMonths[sortedMonths.length-2]]||0;
  const mom        = prevMoAmt ? ((lastMoAmt-prevMoAmt)/prevMoAmt*100).toFixed(1) : 0;
  const avgUnit    = totalUnits ? totalSpend/totalUnits : 0;
  const topCatE    = Object.entries(catSpend).sort((a,b)=>b[1]-a[1])[0]||['—',0];

  document.getElementById('overview-kpis').innerHTML = [
    {{l:'Total Spend',    v:fmt(totalSpend),   s:'',      c:''}},
    {{l:'Total Units',    v:totalUnits,         s:'',      c:''}},
    {{l:'Avg / Unit',     v:fmt(avgUnit),       s:'',      c:''}},
    {{l:lastMo+' Spend',  v:fmt(lastMoAmt),     s:(mom>=0?'▲ ':'▼ ')+Math.abs(mom)+'% MoM', c:mom>=0?'amber':'green'}},
    {{l:'High Cost Units',v:highUnits,           s:'>$5K cumulative', c:highUnits>0?'red':'green'}},
    {{l:'Top Category',   v:topCatE[0],         s:fmt(topCatE[1]),   c:''}},
  ].map(k=>`<div class="kpi ${{k.c}}"><div class="label">${{k.l}}</div><div class="value">${{k.v}}</div><div class="sub">${{k.s}}</div></div>`).join('');

  // Executive summary
  const sortedCats = Object.entries(catSpend).sort((a,b)=>b[1]-a[1]);
  const topVendors = Object.entries(vendorTotals).sort((a,b)=>b[1]-a[1]).slice(0,3);
  const highList   = Object.entries(unitTotals).filter(([,v])=>v>=HIGH).sort((a,b)=>b[1]-a[1]);
  document.getElementById('exec-summary').innerHTML = `
    <h2>Executive Summary</h2>
    <p>Total unit turn spend across <strong>${{totalUnits}} units</strong> is <strong>${{fmt(totalSpend)}}</strong>,
       averaging <strong>${{fmt(avgUnit)}}</strong> per unit.
       ${{lastMo}} spend was <strong>${{fmt(lastMoAmt)}}</strong> (${{mom>=0?'▲':'▼'}}${{Math.abs(mom)}}% vs prior month).</p>
    <p>Top categories: ${{sortedCats.slice(0,3).map(([c,v])=>`<strong>${{c}}</strong> (${{fmt(v)}})`).join(', ')}}.</p>
    <p>Top vendors: ${{topVendors.map(([v,a])=>`<strong>${{v}}</strong> (${{fmt(a)}})`).join(', ')}}.</p>
    ${{highList.length ? `<div class="flag">⚠️ <strong>${{highList.length}} high-cost unit${{highList.length>1?'s':''}}</strong> exceed $5K cumulative: ${{highList.slice(0,5).map(([u,v])=>`Unit ${{u}} (${{fmt(v)}})`).join(', ')}}${{highList.length>5?' and more':''}}</div>` : ''}}
    <div class="rec">✅ Use the <strong>Units</strong> tab to drill into individual units and see meeting notes alongside invoices.</div>`;

  makeChart('chart-monthly','bar',
    sortedMonths.map(m=>m.slice(5)+'/'+m.slice(2,4)),
    sortedMonths.map(m=>monthly[m]),
    {{legend:false,colors:'#3C6BA4'}});

  makeChart('chart-qtr','bar',
    sortedQtrs, sortedQtrs.map(q=>qtrData[q]),
    {{legend:false,colors:'#22C55E'}});

  makeChart('chart-cat-donut','doughnut',
    sortedCats.map(([k])=>k), sortedCats.map(([,v])=>v),
    {{extra:{{plugins:{{legend:{{position:'right',labels:{{font:{{size:11}},boxWidth:12}}}}}}}}}});

  const topV = Object.entries(vendorTotals).sort((a,b)=>b[1]-a[1]).slice(0,10);
  makeChart('chart-vendor-bar','bar',
    topV.map(([k])=>k), topV.map(([,v])=>v),
    {{legend:false,extra:{{indexAxis:'y',
      scales:{{x:{{ticks:{{callback:v=>fmt(v)}},beginAtZero:true}},y:{{ticks:{{font:{{size:10}}}}}}}}}}}});
}}

// ══════════════════════════════════════════════════════════════════
// UNITS
// ══════════════════════════════════════════════════════════════════
let unitSortCol = 'total', unitSortDir = -1;
let allUnitData = [];

function buildUnits(){{
  allUnitData = Object.entries(unitTotals).map(([unit, total])=>{{
    const rs = unitRows[unit]||[];
    const byCat = {{}};
    let lastDate = '';
    rs.forEach(r=>{{
      byCat[r.category]=(byCat[r.category]||0)+r.amount;
      if(r.inst_date > lastDate) lastDate = r.inst_date;
    }});
    const topCat = Object.entries(byCat).sort((a,b)=>b[1]-a[1])[0]||['—',0];
    const hasNotes = !!(COND[unit]&&COND[unit].notes);
    return {{unit, total, invoices:rs.length, topCat:topCat[0], lastDate, hasNotes}};
  }});
  renderUnitTable();
}}

function renderUnitTable(){{
  const q  = document.getElementById('unit-search').value.toLowerCase();
  const sf = document.getElementById('unit-filter').value;

  let data = allUnitData.filter(u=>{{
    if(sf && unitStatus(u.unit)!==sf) return false;
    if(q && !u.unit.includes(q) && !(unitRows[u.unit]||[]).some(r=>r.desc.toLowerCase().includes(q))) return false;
    return true;
  }});

  data.sort((a,b)=>{{
    let va = a[unitSortCol], vb = b[unitSortCol];
    if(unitSortCol==='unit') return unitSortDir*(parseInt(va)-parseInt(vb));
    return unitSortDir*(va-vb);
  }});

  document.getElementById('unit-count').textContent = `${{data.length}} units`;
  document.getElementById('unit-tbody').innerHTML = data.map(u=>{{
    const st  = unitStatus(u.unit);
    const cc  = CC[u.unit];
    const notesBadge = u.hasNotes ? '<span class="badge purple">📋 Notes</span>' : '<span style="color:#cbd5e1">—</span>';
    const ccBadge = cc && cc.flag_list && cc.flag_list.length
      ? `<span class="badge red">⚑ ${{cc.flag_list.length}} flag${{cc.flag_list.length>1?'s':''}}</span>`
      : (cc ? '<span class="badge green">✓ OK</span>' : '<span style="color:#cbd5e1">—</span>');
    return `<tr class="clickable" onclick="drillUnit('${{u.unit}}')">
      <td><strong>Unit ${{u.unit}}</strong></td>
      <td><strong>${{fmt(u.total)}}</strong></td>
      <td>${{u.invoices}}</td>
      <td><span class="badge blue">${{u.topCat}}</span></td>
      <td>${{u.lastDate||'—'}}</td>
      <td>${{statusBadge(st)}}</td>
      <td>${{notesBadge}}</td>
      <td>${{ccBadge}}</td>
    </tr>`;
  }}).join('');
}}

function filterUnits(){{ renderUnitTable(); }}

function sortUnits(col){{
  if(unitSortCol===col) unitSortDir*=-1; else {{ unitSortCol=col; unitSortDir=-1; }}
  document.querySelectorAll('#unit-table th').forEach(th=>{{th.classList.remove('sort-asc','sort-desc')}});
  const idx = {{'unit':0,'total':1,'invoices':2}}[col];
  const ths = document.querySelectorAll('#unit-table th');
  if(ths[idx]) ths[idx].classList.add(unitSortDir>0?'sort-asc':'sort-desc');
  renderUnitTable();
}}

// ── Cross-check renderer ────────────────────────────────────────
function renderCrossCheck(unit){{
  const panel = document.getElementById('unit-cc-panel');
  const cc    = CC[unit];

  if(!cc){{
    panel.classList.add('hidden');
    return;
  }}

  panel.classList.remove('hidden','cc-ok','cc-warn','cc-alert','cc-red');

  const flags = cc.flag_list || [];
  const statusText = cc.status || '';

  // Panel colour
  if(flags.length === 0)                         panel.classList.add('cc-ok');
  else if(statusText.includes('No Invoice'))     panel.classList.add('cc-warn');
  else if(flags.some(f=>f.toLowerCase().includes('repeat')||f.toLowerCase().includes('creep')))
                                                 panel.classList.add('cc-alert');
  else                                           panel.classList.add('cc-warn');

  // Title
  const titleEl = document.getElementById('unit-cc-title');
  const icon = flags.length===0 ? '✅' : statusText.includes('Not Discussed') ? '🔍' :
               statusText.includes('No Invoice') ? '⚠️' : '⚑';
  titleEl.textContent = `${{icon}} Cross-Check — ${{statusText.replace(/[✅⚠️🔍]/g,'').trim()}}`;
  titleEl.style.color = flags.length===0 ? '#16a34a' : '#92400e';

  // Context note badge
  const ctxEl = document.getElementById('unit-cc-context');
  ctxEl.textContent = cc.context && cc.context !== '—' ? cc.context : '';
  ctxEl.style.display = ctxEl.textContent ? 'inline-block' : 'none';

  // Flags list
  const flagsEl = document.getElementById('unit-cc-flags');
  if(flags.length === 0){{
    flagsEl.innerHTML = '<div class="cc-ok-msg">No issues detected — notes and invoices are consistent.</div>';
  }} else {{
    flagsEl.innerHTML = flags.map(f=>{{
      let cls = 'flag-warn';
      const fl = f.toLowerCase();
      if(fl.includes('repeat')||fl.includes('creep')||fl.includes('unexpected')) cls='flag-alert';
      if(fl.includes('easy turn')||fl.includes('not discussed')) cls='flag-red';
      if(fl.includes('expected but no invoice')||fl.includes('no invoice')) cls='flag-info';
      return `<div class="cc-flag ${{cls}}">${{f}}</div>`;
    }}).join('');

    // Show expected vs actual if useful
    if(cc.expected_work && cc.expected_work !== '—' && cc.actual_invoices && cc.actual_invoices !== '—'){{
      flagsEl.innerHTML += `
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px;font-size:11px">
          <div style="background:#f0f9ff;border-radius:6px;padding:8px">
            <div style="font-weight:700;color:#0369a1;margin-bottom:4px">Expected from notes</div>
            <div style="color:#374151">${{cc.expected_work}}</div>
          </div>
          <div style="background:#fafafa;border-radius:6px;padding:8px">
            <div style="font-weight:700;color:#374151;margin-bottom:4px">Actual invoices billed</div>
            <div style="color:#374151">${{cc.actual_invoices}}</div>
          </div>
        </div>`;
    }}
  }}
}}

// ── Meeting notes renderer ──────────────────────────────────────
function renderNotes(unit){{
  const cond = COND[unit];
  const panel = document.getElementById('unit-notes-panel');
  const entriesEl = document.getElementById('unit-notes-entries');
  const metaEl = document.getElementById('unit-notes-meta');

  if(!cond || !cond.notes || !cond.notes.trim()){{
    entriesEl.innerHTML = '<div class="no-notes">No meeting notes recorded for this unit yet.</div>';
    metaEl.textContent = '';
    return;
  }}

  const rawNotes = cond.notes;

  // Parse entries: look for date headers like "6/24/26" or "6/24/2026" at start of a line
  const dateHeaderRe = /^(\\d{{1,2}}\\/\\d{{1,2}}\\/\\d{{2,4}})\\s*$/gm;
  const parts = [];
  let lastIndex = 0, lastDate = null, m;
  while((m = dateHeaderRe.exec(rawNotes)) !== null){{
    if(lastDate !== null){{
      parts.push({{date: lastDate, body: rawNotes.slice(lastIndex, m.index).trim()}});
    }}
    lastDate = m[1];
    lastIndex = m.index + m[0].length;
  }}
  if(lastDate !== null){{
    parts.push({{date: lastDate, body: rawNotes.slice(lastIndex).trim()}});
  }}
  if(parts.length === 0){{
    // No date headers found — show as single block
    parts.push({{date: '', body: rawNotes.trim()}});
  }}

  // Most recent first
  parts.reverse();
  metaEl.textContent = `${{parts.length}} entr${{parts.length===1?'y':'ies'}}`;

  entriesEl.innerHTML = parts.map(p=>`
    <div class="note-entry">
      ${{p.date ? `<div class="note-date">${{p.date}}</div>` : ''}}
      <div class="note-body">${{p.body.replace(/</g,'&lt;').replace(/>/g,'&gt;')}}</div>
    </div>`).join('');

  // Extra metadata from other columns
  const extras = [];
  if(cond.date_turned)   extras.push(`Turned: ${{cond.date_turned}}`);
  if(cond.make_ready_by) extras.push(`Make-ready by: ${{cond.make_ready_by}}`);
  if(cond.required_work) extras.push(`Required work: ${{cond.required_work}}`);
  if(extras.length){{
    entriesEl.innerHTML += `<div style="font-size:11px;color:#64748b;margin-top:8px;padding-top:8px;border-top:1px solid #e2e8f0">${{extras.join(' · ')}}</div>`;
  }}
}}

let unitDrillCharts = [];
function drillUnit(unit){{
  const rs = unitRows[unit]||[];
  const total = unitTotals[unit]||0;
  const status = unitStatus(unit);
  document.getElementById('unit-drill').classList.remove('hidden');
  document.getElementById('unit-drill-title').textContent = `Unit ${{unit}} — ${{fmt(total)}} total`;

  const byCat = {{}};
  const byMonth = {{}};
  rs.forEach(r=>{{
    byCat[r.category]=(byCat[r.category]||0)+r.amount;
    byMonth[r.month]=(byMonth[r.month]||0)+r.amount;
  }});
  const topCat = Object.entries(byCat).sort((a,b)=>b[1]-a[1])[0]||['—',0];
  document.getElementById('unit-drill-kpis').innerHTML = [
    {{label:'Total Spend',   value:fmt(total),  cls:status}},
    {{label:'# Invoices',    value:rs.length,   cls:''}},
    {{label:'Avg / Invoice', value:fmt(total/rs.length), cls:''}},
    {{label:'Top Category',  value:topCat[0],   cls:''}},
    {{label:'Status',        value:status==='high'?'🚨 HIGH':status==='watch'?'⚠ WATCH':'✅ OK', cls:status}},
  ].map(k=>`<div class="kpi ${{k.cls}}"><div class="label">${{k.label}}</div><div class="value" style="font-size:20px">${{k.value}}</div></div>`).join('');

  unitDrillCharts.forEach(c=>c.destroy()); unitDrillCharts=[];
  const catSorted = Object.entries(byCat).sort((a,b)=>b[1]-a[1]);
  unitDrillCharts.push(makeChart('unit-cat-chart','doughnut',
    catSorted.map(([k])=>k), catSorted.map(([,v])=>v),
    {{legendPos:'right',extra:{{plugins:{{legend:{{position:'right',labels:{{font:{{size:11}},boxWidth:12}}}}}}}}}}));

  const months = Object.keys(byMonth).sort();
  unitDrillCharts.push(makeChart('unit-time-chart','bar',
    months.map(m=>m.slice(5)+'/'+m.slice(2,4)),
    months.map(m=>byMonth[m]),
    {{legend:false,colors:'#3C6BA4'}}));

  renderCrossCheck(unit);
  renderNotes(unit);

  document.getElementById('unit-invoices-tbody').innerHTML = rs
    .sort((a,b)=>(b.inst_date||b.inv_date).localeCompare(a.inst_date||a.inv_date))
    .map(r=>`<tr>
      <td>${{r.invoice}}</td><td>${{r.inv_date}}</td><td>${{r.inst_date}}</td>
      <td>${{r.desc}}</td><td><span class="badge blue">${{r.category}}</span></td>
      <td>${{r.vendor}}</td><td><strong>${{fmt(r.amount)}}</strong></td>
    </tr>`).join('');

  document.getElementById('unit-drill').scrollIntoView({{behavior:'smooth',block:'start'}});
}}

// ══════════════════════════════════════════════════════════════════
// CATEGORIES
// ══════════════════════════════════════════════════════════════════
let catDrillData = [];
function buildCategories(){{
  const catSorted = Object.entries(catSpend).sort((a,b)=>b[1]-a[1]);
  makeChart('chart-cat-bar','bar',
    catSorted.map(([k])=>k),catSorted.map(([,v])=>v),
    {{legend:false,extra:{{indexAxis:'y',
     scales:{{x:{{ticks:{{callback:v=>fmt(v)}},beginAtZero:true}},y:{{ticks:{{font:{{size:11}}}}}}}}}}}});
  makeChart('chart-cat-pie2','doughnut',
    catSorted.map(([k])=>k),catSorted.map(([,v])=>v),
    {{extra:{{plugins:{{legend:{{position:'right',labels:{{font:{{size:11}},boxWidth:14}}}}}}}}}});
  document.getElementById('cat-tbody').innerHTML = catSorted.map(([cat,amt])=>`
    <tr class="clickable" onclick="drillCategory('${{cat}}')">
      <td><strong>${{cat}}</strong></td><td>${{fmt(amt)}}</td>
      <td>${{pct(amt,totalSpend)}}</td><td>${{catCnt[cat]||0}}</td>
      <td>${{fmt(amt/(catCnt[cat]||1))}}</td>
    </tr>`).join('');
}}

function drillCategory(cat){{
  document.getElementById('cat-drill').classList.remove('hidden');
  document.getElementById('cat-drill-title').textContent = `${{cat}} — by Unit`;
  const byUnit = {{}};
  rows.forEach(r=>{{if(r.category===cat){{
    byUnit[r.unit]={{amt:(byUnit[r.unit]?.amt||0)+r.amount, cnt:(byUnit[r.unit]?.cnt||0)+1}};
  }}}});
  catDrillData = Object.entries(byUnit).sort((a,b)=>b[1].amt-a[1].amt);
  renderCatDrill();
  document.getElementById('cat-drill').scrollIntoView({{behavior:'smooth',block:'start'}});
}}
function renderCatDrill(){{
  const q = document.getElementById('cat-drill-search').value.toLowerCase();
  const data = catDrillData.filter(([u])=>u.includes(q));
  document.getElementById('cat-drill-tbody').innerHTML = data.map(([unit,d])=>`
    <tr class="clickable" onclick="showPage('units',document.querySelector('.nav button:nth-child(2)'));setTimeout(()=>drillUnit('${{unit}}'),100)">
      <td><strong>Unit ${{unit}}</strong></td><td>${{fmt(d.amt)}}</td>
      <td>${{d.cnt}}</td><td>${{statusBadge(unitStatus(unit))}}</td>
    </tr>`).join('');
}}
function filterCatDrill(){{ renderCatDrill(); }}

// ══════════════════════════════════════════════════════════════════
// VENDORS
// ══════════════════════════════════════════════════════════════════
let allVendorData = [];
function buildVendors(){{
  allVendorData = Object.entries(vendorTotals).sort((a,b)=>b[1]-a[1]);
  const top15 = allVendorData.slice(0,15);
  makeChart('chart-vendor-bar2','bar',
    top15.map(([k])=>k),top15.map(([,v])=>v),
    {{legend:false,extra:{{indexAxis:'y',
     scales:{{x:{{ticks:{{callback:v=>fmt(v)}},beginAtZero:true}},y:{{ticks:{{font:{{size:10}}}}}}}}}}}});
  renderVendors();
}}
function renderVendors(){{
  const q = document.getElementById('vendor-search').value.toLowerCase();
  const data = allVendorData.filter(([v])=>v.toLowerCase().includes(q));
  document.getElementById('vendor-tbody').innerHTML = data.map(([vendor,amt],i)=>`
    <tr class="clickable" onclick="drillVendor('${{vendor.replace(/'/g,"\\\\'")}}')">
      <td>${{i+1}}</td><td><strong>${{vendor}}</strong></td><td>${{fmt(amt)}}</td>
      <td>${{vendorCnt[vendor]||0}}</td><td>${{fmt(amt/(vendorCnt[vendor]||1))}}</td>
    </tr>`).join('');
}}
function filterVendors(){{ renderVendors(); }}

function drillVendor(vendor){{
  document.getElementById('vendor-drill').classList.remove('hidden');
  document.getElementById('vendor-drill-title').textContent = vendor;
  const vrows = rows.filter(r=>r.vendor===vendor).sort((a,b)=>(b.inst_date||b.inv_date).localeCompare(a.inst_date||a.inv_date));
  const byUnit={{}}, byCat={{}};
  vrows.forEach(r=>{{
    byUnit[r.unit]=(byUnit[r.unit]||0)+r.amount;
    byCat[r.category]=(byCat[r.category]||0)+r.amount;
  }});
  const unitSorted = Object.entries(byUnit).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const catSorted  = Object.entries(byCat).sort((a,b)=>b[1]-a[1]);
  drillCharts['vendor-drill']?.forEach(c=>c.destroy());
  drillCharts['vendor-drill'] = [
    makeChart('vendor-unit-chart','bar',unitSorted.map(([u])=>'Unit '+u),unitSorted.map(([,v])=>v),
      {{legend:false,colors:'#3C6BA4'}}),
    makeChart('vendor-cat-chart','doughnut',catSorted.map(([k])=>k),catSorted.map(([,v])=>v),
      {{extra:{{plugins:{{legend:{{position:'right',labels:{{font:{{size:11}},boxWidth:12}}}}}}}}}})
  ];
  document.getElementById('vendor-drill-tbody').innerHTML = vrows.map(r=>`<tr>
    <td><strong>Unit ${{r.unit}}</strong></td><td>${{r.invoice}}</td>
    <td>${{r.inst_date||r.inv_date}}</td><td>${{r.desc}}</td>
    <td><span class="badge blue">${{r.category}}</span></td>
    <td><strong>${{fmt(r.amount)}}</strong></td>
  </tr>`).join('');
  document.getElementById('vendor-drill').scrollIntoView({{behavior:'smooth',block:'start'}});
}}

// ══════════════════════════════════════════════════════════════════
// ALL INVOICES
// ══════════════════════════════════════════════════════════════════
let invSortCol='inst_date', invSortDir=-1;
let filteredInvRows = [...rows];

function buildInvoicesPage(){{
  const cats   = [...new Set(rows.map(r=>r.category))].sort();
  const months = [...new Set(rows.map(r=>r.month))].sort().reverse();
  const qtrs   = [...new Set(rows.map(r=>r.quarter))].sort().reverse();
  document.getElementById('inv-cat').innerHTML   += cats.map(c=>`<option value="${{c}}">${{c}}</option>`).join('');
  document.getElementById('inv-month').innerHTML += months.map(m=>`<option value="${{m}}">${{m}}</option>`).join('');
  document.getElementById('inv-qtr').innerHTML   += qtrs.map(q=>`<option value="${{q}}">${{q}}</option>`).join('');
  filterInvoices();
}}
function filterInvoices(){{
  const q   = document.getElementById('inv-search').value.toLowerCase();
  const cat = document.getElementById('inv-cat').value;
  const mo  = document.getElementById('inv-month').value;
  const qt  = document.getElementById('inv-qtr').value;
  filteredInvRows = rows.filter(r=>
    (!cat || r.category===cat) && (!mo || r.month===mo) && (!qt || r.quarter===qt) &&
    (!q  || r.unit.includes(q)||r.desc.toLowerCase().includes(q)||r.vendor.toLowerCase().includes(q)||r.invoice.toLowerCase().includes(q))
  );
  renderInvoices();
}}
function renderInvoices(){{
  const sorted = [...filteredInvRows].sort((a,b)=>{{
    let va=a[invSortCol]||'', vb=b[invSortCol]||'';
    if(invSortCol==='amount') return invSortDir*(a.amount-b.amount);
    return invSortDir*va.localeCompare(vb);
  }});
  document.getElementById('inv-count').textContent = `${{sorted.length}} invoices`;
  const total = sorted.reduce((s,r)=>s+r.amount,0);
  document.getElementById('inv-total-row').textContent = `Showing ${{sorted.length}} invoices · Total: ${{fmt(total)}}`;
  document.getElementById('inv-tbody').innerHTML = sorted.map(r=>`<tr>
    <td><strong>Unit ${{r.unit}}</strong></td>
    <td>${{r.invoice}}</td><td>${{r.inv_date}}</td><td>${{r.inst_date}}</td>
    <td>${{r.desc}}</td>
    <td><span class="badge blue">${{r.category}}</span></td>
    <td>${{r.vendor}}</td>
    <td><strong>${{fmt(r.amount)}}</strong></td>
  </tr>`).join('');
}}
function sortInvoices(col){{
  if(invSortCol===col) invSortDir*=-1; else {{invSortCol=col; invSortDir=-1;}}
  renderInvoices();
}}

// ══════════════════════════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════════════════════════
buildOverview();
buildUnits();
buildCategories();
buildVendors();
buildInvoicesPage();
</script>
</body>
</html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ Dashboard written to {out_path} ({len(html)//1024}KB)")


def main(argv=None):
    parser = argparse.ArgumentParser(description="Build Creekside dashboard HTML")
    parser.add_argument("--out", default=OUT_FILE, help="Output HTML file path")
    args = parser.parse_args(argv)

    print("Connecting to Google Sheets…")
    gc = get_gspread_client()

    print("Reading invoice data…")
    raw_data = read_invoice_data(gc)
    print(f"  {len(raw_data['rows'])} invoice rows, {len(raw_data['vendor_spend'])} vendors, {len(raw_data['cat_spend'])} categories")

    print("Reading Unit Conditions…")
    cond_data = read_conditions(gc)
    print(f"  {len(cond_data)} units with condition records")

    print("Reading Cross-Check findings…")
    cc_data = read_crosscheck(gc)
    print(f"  {len(cc_data)} units with cross-check data")

    print(f"Building dashboard → {args.out}")
    build_html(raw_data, cond_data, cc_data, args.out)


if __name__ == "__main__":
    main()
