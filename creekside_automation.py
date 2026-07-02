"""
Creekside Unit Turn Automation
==============================
Monthly workflow:
  1. Export "Invoice by Location" from Resman -> save as CSV
  2. Export "Invoice by Detail" from Resman   -> save as CSV
  3. Run:
       python creekside_automation.py \
           --location  invoices_by_location.csv \
           --detail    invoices_by_detail.csv \
           --month     2026-06 \
           [--dry-run]

Requires:
  - GOOGLE_SERVICE_ACCOUNT_JSON env var  -> path to service account key file
  - OR GOOGLE_CREDENTIALS_JSON           -> raw JSON string of the key
  - Sheet must be shared with the service account email (Editor access)
"""

import os
import re
import sys
import json
import argparse
import csv
import io
import time
from datetime import datetime, date
from collections import defaultdict

# ---------------------------------------------------------------------------
# Constants — sheet names and spreadsheet ID
# ---------------------------------------------------------------------------
SPREADSHEET_ID   = "1qVJ3Nz4LCgZVKlWMLP8i2xj8F0tzD6ujGDBtGhM6JhM"
SRC_SHEET        = "DATA Invoice by Location"
DET_SHEET        = "Invoice by Detail Report"
UT_SHEET         = "Unit Turn Cost by Unit"
COND_SHEET       = "Unit Conditions"
PERQ_SHEET       = "Quarterly Turn Cost Summary"

# Marketing filter keywords (same logic as AppScript)
MARKETING_KEYWORDS = [
    "LOCATOR COMMISSION", "LOCATOR",
    "APARTMENT LIST", "APARTMENTLIST", "APARTMENTLIST.COM",
    "LEAD EXPENSE", "LEADS EXPENSE",
    "LIFT LEAD", "LIFTLEAD",
    "MOVE-IN FEE", "MOVE IN FEE", "MOVEIN FEE",
    "RENTERS INSURANCE", "RENTER'S INSURANCE", "RENTER INSURANCE",
]

MARKETING_REASONS = {
    "Locator commission":          ["LOCATOR COMMISSION", "LOCATOR"],
    "Apartment List / lead expense": ["APARTMENT LIST", "APARTMENTLIST", "APARTMENTLIST.COM"],
    "Lead expense":                ["LEAD EXPENSE", "LEADS EXPENSE"],
    "Lift Lead":                   ["LIFT LEAD", "LIFTLEAD"],
    "Move-in Fee":                 ["MOVE-IN FEE", "MOVE IN FEE", "MOVEIN FEE"],
    "Renters Insurance":           ["RENTERS INSURANCE", "RENTER'S INSURANCE", "RENTER INSURANCE"],
}

HIGH_COST_THRESHOLD = 5000   # flag units over this amount in analysis report

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

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def s(v):
    return str(v or "").strip()


def norm_unit(v):
    digits = re.sub(r"[^0-9]", "", s(v))
    n = int(digits) if digits else 0
    return str(n) if 100 <= n <= 1500 else ""


def to_date(v):
    if isinstance(v, (date, datetime)):
        return v if isinstance(v, datetime) else datetime(v.year, v.month, v.day)
    raw = s(v)
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%-m/%-d/%Y", "%-m/%-d/%y"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def quarter_of(d):
    if not d:
        return ""
    q = (d.month - 1) // 3 + 1
    return f"{d.year}Q{q}"


def norm_text(v):
    return re.sub(r"\s+", " ", s(v)).upper()


def norm_inv(v):
    return s(v).upper()


def norm_desc(v):
    return re.sub(r"\s+", " ", s(v)).upper()


def norm_amt(v):
    try:
        return round(float(re.sub(r"[^0-9.\-]", "", s(v)) or "0"), 2)
    except ValueError:
        return 0.0


def build_key(unit, inv, desc):
    """Dedup key: unit|invoice|description (no amount — sheet displays truncated values)."""
    return f"{norm_unit(unit)}|{norm_inv(inv)}|{norm_desc(desc)}"


def is_marketing(row_texts):
    combined = " | ".join(norm_text(v) for v in row_texts)
    return any(kw in combined for kw in MARKETING_KEYWORDS)


def get_marketing_reason(row_texts):
    combined = " | ".join(norm_text(v) for v in row_texts)
    reasons = [r for r, kws in MARKETING_REASONS.items() if any(kw in combined for kw in kws)]
    return ", ".join(reasons)


# ---------------------------------------------------------------------------
# CSV loading  (handles Resman export quirks)
# ---------------------------------------------------------------------------

def load_csv(path):
    """Return list of dicts. Tries utf-8 then latin-1."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(path, newline="", encoding=enc) as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                if rows:
                    return rows
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Cannot decode {path}")


def _find_col(headers, *candidates):
    """Case-insensitive column finder."""
    upper = {h.upper(): h for h in headers}
    for c in candidates:
        if c.upper() in upper:
            return upper[c.upper()]
    return None


# ---------------------------------------------------------------------------
# Step 1 — Parse "Invoice by Location" CSV
# ---------------------------------------------------------------------------

def parse_location_rows(path):
    """
    Returns cleaned list of dicts with canonical keys:
      unit, invoice, install_date, acctg_date, amount, gl, description, property
    Skips marketing rows and rows with no valid unit.
    """
    raw = load_csv(path)
    if not raw:
        return [], []

    headers = list(raw[0].keys())
    col = lambda *c: _find_col(headers, *c)

    c_prop  = col("PropertyName", "Property", "Property Name")
    c_unit  = col("Unit#", "Unit #", "Unit", "ObjectName", "Object Name")
    c_otype = col("ObjectType", "Object Type")
    c_inv   = col("Invoice #", "Invoice#", "Invoice Number", "InvoiceNumber")
    c_inst  = col("Install Date", "InstallDate", "Service Date")
    c_acctg = col("Acctg Date", "AcctgDate", "Accounting Date", "AccountingDate", "Post Date")
    c_amt   = col("Total", "Amount", "Net Amount")
    c_gl    = col("GL Acc Number", "GL Account", "Account")
    c_desc  = col("Description", "Memo")

    kept = []
    removed = []

    for row in raw:
        vals = list(row.values())
        texts = [s(row.get(h, "")) for h in headers]

        # Try to extract unit; fall back to description
        raw_unit = s(row.get(c_unit, "")) if c_unit else ""
        if not raw_unit or raw_unit.lower() == "none":
            desc_val = s(row.get(c_desc, "")) if c_desc else ""
            m = re.search(r"#(\d{3,4})", desc_val)
            if m:
                raw_unit = m.group(1)

        unit = norm_unit(raw_unit)

        if is_marketing(texts):
            removed.append({
                "unit": unit or raw_unit,
                "invoice": norm_inv(row.get(c_inv, "")) if c_inv else "",
                "reason": get_marketing_reason(texts),
                "description": s(row.get(c_desc, "")) if c_desc else "",
                "amount": norm_amt(row.get(c_amt, "")) if c_amt else 0,
            })
            continue

        if not unit:
            removed.append({
                "unit": raw_unit,
                "invoice": norm_inv(row.get(c_inv, "")) if c_inv else "",
                "reason": "Missing / invalid unit number",
                "description": s(row.get(c_desc, "")) if c_desc else "",
                "amount": norm_amt(row.get(c_amt, "")) if c_amt else 0,
            })
            continue

        kept.append({
            "property":     s(row.get(c_prop, "Oaks at Creekside")) if c_prop else "Oaks at Creekside",
            "unit":         unit,
            "obj_type":     s(row.get(c_otype, "Unit")) if c_otype else "Unit",
            "invoice":      norm_inv(row.get(c_inv, "")) if c_inv else "",
            "install_date": to_date(row.get(c_inst, "")) if c_inst else None,
            "acctg_date":   to_date(row.get(c_acctg, "")) if c_acctg else None,
            "amount":       norm_amt(row.get(c_amt, "")) if c_amt else 0,
            "gl":           s(row.get(c_gl, "")) if c_gl else "",
            "description":  s(row.get(c_desc, "")) if c_desc else "",
            "category":     "Unit Turn",
        })

    return kept, removed


# ---------------------------------------------------------------------------
# Step 2 — Parse "Invoice by Detail" CSV
# ---------------------------------------------------------------------------

def parse_detail_rows(path):
    """
    Returns dict: invoice_number -> invoice_date (datetime)
    """
    raw = load_csv(path)
    if not raw:
        return {}

    headers = list(raw[0].keys())
    col = lambda *c: _find_col(headers, *c)

    c_inv  = col("InvoiceNumber", "Invoice Number", "Invoice #", "Invoice#")
    c_date = col("InvoiceDate", "Invoice Date", "Date")

    mapping = {}
    for row in raw:
        inv = norm_inv(row.get(c_inv, "")) if c_inv else ""
        d   = to_date(row.get(c_date, "")) if c_date else None
        if inv and inv not in mapping:
            mapping[inv] = d

    return mapping


# ---------------------------------------------------------------------------
# Google Sheets connection
# ---------------------------------------------------------------------------

def get_gspread_client():
    """Authenticate via service account. Returns gspread client."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("ERROR: gspread not installed. Run:  pip install gspread google-auth")
        sys.exit(1)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    raw_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    json_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

    if raw_json:
        info = json.loads(raw_json)
        creds = Credentials.from_service_account_info(info, scopes=scopes)
    elif json_path:
        creds = Credentials.from_service_account_file(json_path, scopes=scopes)
    else:
        print("ERROR: Set GOOGLE_SERVICE_ACCOUNT_JSON (path) or GOOGLE_CREDENTIALS_JSON (raw JSON).")
        sys.exit(1)

    return gspread.authorize(creds)


# ---------------------------------------------------------------------------
# Read existing sheet data
# ---------------------------------------------------------------------------

def read_sheet_as_list(ws):
    """Return all rows as list of lists (including header row)."""
    return ws.get_all_values()


def header_index(headers, *candidates):
    """Return 0-based index of first matching header (case-insensitive)."""
    upper = [h.strip().upper() for h in headers]
    for c in candidates:
        if c.upper() in upper:
            return upper.index(c.upper())
    return -1


# ---------------------------------------------------------------------------
# Update: DATA Invoice by Location
# ---------------------------------------------------------------------------

def update_src_sheet(ws, new_rows, removed_rows, dry_run=False):
    """Append new location invoice rows; add/update Category + Removal Reason cols."""
    all_vals = read_sheet_as_list(ws)
    if not all_vals:
        print(f"  [WARN] {SRC_SHEET} is empty — cannot update.")
        return

    headers = all_vals[0]

    # Ensure Category and Removal Reason columns exist
    cat_idx    = header_index(headers, "Category")
    reason_idx = header_index(headers, "Removal Reason")

    if cat_idx == -1:
        headers.append("Category")
        cat_idx = len(headers) - 1
    if reason_idx == -1:
        headers.append("Removal Reason")
        reason_idx = len(headers) - 1

    # Build existing key set to avoid duplicates
    existing_inv = set()
    for row in all_vals[1:]:
        inv_val = row[3] if len(row) > 3 else ""
        existing_inv.add(norm_inv(inv_val))

    to_append = [r for r in new_rows if r["invoice"] not in existing_inv]

    print(f"\n[{SRC_SHEET}]")
    print(f"  Existing rows : {len(all_vals) - 1}")
    print(f"  New to append : {len(to_append)}")
    print(f"  Skipped (dupe): {len(new_rows) - len(to_append)}")
    print(f"  Removed (mktg/no-unit): {len(removed_rows)}")

    if dry_run or not to_append:
        return

    append_data = []
    for r in to_append:
        row_out = [""] * max(len(headers), 10)
        row_out[0] = r["property"]
        row_out[1] = r["unit"]
        row_out[2] = r["obj_type"]
        row_out[3] = r["invoice"]
        row_out[4] = r["install_date"].strftime("%-m/%-d/%y") if r["install_date"] else ""
        row_out[5] = r["acctg_date"].strftime("%-m/%-d/%y") if r["acctg_date"] else ""
        row_out[6] = str(int(r["amount"])) if r["amount"] == int(r["amount"]) else str(r["amount"])
        row_out[7] = r["gl"]
        row_out[8] = r["description"]
        if cat_idx < len(row_out):
            row_out[cat_idx] = r["category"]
        append_data.append(row_out)

    ws.append_rows(append_data, value_input_option="USER_ENTERED")
    print(f"  ✅ Appended {len(append_data)} rows.")


# ---------------------------------------------------------------------------
# Update: Invoice by Detail Report
# ---------------------------------------------------------------------------

def update_det_sheet(ws, detail_path, dry_run=False):
    """Append new detail rows that don't already exist."""
    new_raw = load_csv(detail_path)
    all_vals = read_sheet_as_list(ws)

    existing_inv = set()
    if len(all_vals) > 1:
        headers = all_vals[0]
        inv_col = header_index(headers, "InvoiceNumber", "Invoice Number", "Invoice #")
        for row in all_vals[1:]:
            if inv_col >= 0 and inv_col < len(row):
                existing_inv.add(norm_inv(row[inv_col]))

    to_append = []
    headers_new = list(new_raw[0].keys()) if new_raw else []
    for row in new_raw:
        inv = norm_inv(row.get(_find_col(headers_new, "InvoiceNumber", "Invoice Number", "Invoice #") or "", ""))
        if inv and inv not in existing_inv:
            to_append.append(list(row.values()))
            existing_inv.add(inv)

    print(f"\n[{DET_SHEET}]")
    print(f"  New detail rows to append: {len(to_append)}")

    if dry_run or not to_append:
        return

    ws.append_rows(to_append, value_input_option="USER_ENTERED")
    print(f"  ✅ Appended {len(to_append)} detail rows.")


# ---------------------------------------------------------------------------
# Update: Unit Turn Cost by Unit
# ---------------------------------------------------------------------------

def update_unit_turn_sheet(ws, new_rows, detail_map, target_month, wb_gs=None, dry_run=False):
    """
    Merge new_rows filtered to target_month into Unit Turn Cost by Unit.
    Columns: Unit | Invoice# | Invoice Date | Install Date | Description | Amount | Unit Total | Duplicate?
    """
    month_rows = [
        r for r in new_rows
        if r["acctg_date"] and
           r["acctg_date"].year == target_month.year and
           r["acctg_date"].month == target_month.month
    ]

    all_vals = read_sheet_as_list(ws)
    existing_keys = set()
    if len(all_vals) > 1:
        for row in all_vals[1:]:
            if len(row) >= 5:
                key = build_key(row[0], row[1], row[4])
                existing_keys.add(key)

    to_add = []
    new_keys = set()
    for r in month_rows:
        key = build_key(r["unit"], r["invoice"], r["description"])
        if key not in existing_keys and key not in new_keys:
            new_keys.add(key)
            inv_date = detail_map.get(r["invoice"])
            to_add.append({
                "unit":       r["unit"],
                "invoice":    r["invoice"],
                "inv_date":   inv_date,
                "inst_date":  r["install_date"],
                "desc":       r["description"],
                "amount":     r["amount"],
                "key":        key,
            })

    print(f"\n[{UT_SHEET}]")
    print(f"  Month rows found  : {len(month_rows)}")
    print(f"  New (not in sheet): {len(to_add)}")

    if dry_run or not to_add:
        return to_add

    # Combine with existing data rows
    existing_data = all_vals[1:] if len(all_vals) > 1 else []

    def row_to_dict(r):
        return {
            "unit":      norm_unit(r[0]) if len(r) > 0 else "",
            "invoice":   r[1] if len(r) > 1 else "",
            "inv_date":  to_date(r[2]) if len(r) > 2 else None,
            "inst_date": to_date(r[3]) if len(r) > 3 else None,
            "desc":      r[4] if len(r) > 4 else "",
            "amount":    norm_amt(r[5]) if len(r) > 5 else 0,
            "key":       build_key(
                norm_unit(r[0]) if len(r) > 0 else "",
                r[1] if len(r) > 1 else "",
                r[4] if len(r) > 4 else "",
            ),
            "is_new": False,
        }

    combined = [row_to_dict(r) for r in existing_data]
    for item in to_add:
        item["is_new"] = True
        combined.append(item)

    # Sort by unit number then install date
    combined.sort(key=lambda r: (
        int(r["unit"]) if r["unit"].isdigit() else 9999,
        r["inst_date"] or datetime.min
    ))

    # Compute per-unit totals
    unit_totals = defaultdict(float)
    for r in combined:
        unit_totals[r["unit"]] += float(r.get("amount", 0) or 0)

    # Duplicate detection
    key_counts = defaultdict(int)
    for r in combined:
        key_counts[r["key"]] += 1

    # Build output rows
    fmt_date = lambda d: d.strftime("%-m/%-d/%Y") if d else ""
    out_rows = []
    seen_units = set()
    for r in combined:
        unit_total = unit_totals[r["unit"]] if r["unit"] not in seen_units else ""
        seen_units.add(r["unit"])
        dup = key_counts[r["key"]] if key_counts[r["key"]] > 1 else ""
        out_rows.append([
            r["unit"],
            r["invoice"],
            fmt_date(r.get("inv_date")),
            fmt_date(r.get("inst_date")),
            r["desc"],
            str(int(r["amount"])) if r.get("amount") and r["amount"] == int(r["amount"]) else str(r.get("amount", "")),
            str(int(unit_total)) if isinstance(unit_total, float) and unit_total == int(unit_total) else str(unit_total or ""),
            str(dup) if dup else "",
        ])

    # Write all data rows (replace rows 2 onward)
    header_row = ["Unit", "Invoice #", "Invoice Date", "Install Date", "Description", "Amount", "Unit Total", "Duplicate?"]
    ws.update("A1", [header_row], value_input_option="USER_ENTERED")

    if out_rows:
        # Clear existing data area first
        last_row = max(len(all_vals), len(out_rows) + 1)
        if last_row > 1:
            ws.batch_clear([f"A2:H{last_row + 10}"])
        ws.update("A2", out_rows, value_input_option="USER_ENTERED")

    time.sleep(2)
    print(f"  ✅ Written {len(out_rows)} rows ({len(to_add)} new).")

    # Apply formatting: pastel banding per unit + red font for new month rows
    _apply_ut_formatting(ws, wb_gs=wb_gs or ws.spreadsheet, out_rows=out_rows, new_month_keys={
        build_key(r["unit"], r["invoice"], r["description"]) for r in to_add
    })

    return to_add


PASTEL_COLORS = [
    {"red": 0.957, "green": 0.800, "blue": 0.800},  # #f4cccc
    {"red": 0.988, "green": 0.898, "blue": 0.804},  # #fce5cd
    {"red": 1.000, "green": 0.949, "blue": 0.800},  # #fff2cc
    {"red": 0.851, "green": 0.918, "blue": 0.827},  # #d9ead3
    {"red": 0.816, "green": 0.878, "blue": 0.890},  # #d0e0e3
    {"red": 0.812, "green": 0.886, "blue": 0.953},  # #cfe2f3
    {"red": 0.851, "green": 0.824, "blue": 0.914},  # #d9d2e9
    {"red": 0.918, "green": 0.820, "blue": 0.863},  # #ead1dc
]


def _apply_ut_formatting(ws, wb_gs, out_rows, new_month_keys):
    """Apply pastel background per unit group + red font for new month rows."""
    import time as _time
    sid = ws.id
    requests = []

    # Clear all formatting in data rows first
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": sid,
                "startRowIndex": 1,
                "endRowIndex": len(out_rows) + 2,
                "startColumnIndex": 0,
                "endColumnIndex": 8,
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                    "textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False},
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    })

    # Assign pastel color per unique unit (order of first appearance)
    seen_units = []
    seen_set = set()
    for row in out_rows:
        u = row[0]
        if u not in seen_set:
            seen_set.add(u)
            seen_units.append(u)
    unit_color = {u: PASTEL_COLORS[i % len(PASTEL_COLORS)] for i, u in enumerate(seen_units)}

    # Build pastel background requests (group consecutive rows of same unit)
    current_unit = None
    group_start = None
    for i, row in enumerate(out_rows):
        u = row[0]
        if u != current_unit:
            if current_unit is not None and group_start is not None:
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sid,
                            "startRowIndex": group_start + 1,
                            "endRowIndex": i + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 8,
                        },
                        "cell": {"userEnteredFormat": {"backgroundColor": unit_color.get(current_unit, {"red": 1, "green": 1, "blue": 1})}},
                        "fields": "userEnteredFormat.backgroundColor",
                    }
                })
            current_unit = u
            group_start = i
    if current_unit is not None and group_start is not None:
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sid,
                    "startRowIndex": group_start + 1,
                    "endRowIndex": len(out_rows) + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 8,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": unit_color.get(current_unit, {"red": 1, "green": 1, "blue": 1})}},
                "fields": "userEnteredFormat.backgroundColor",
            }
        })

    # Red font for new month rows (no background change, just font color)
    red = {"red": 1.0, "green": 0.0, "blue": 0.0}
    for i, row in enumerate(out_rows):
        k = build_key(row[0], row[1], row[4])
        if k in new_month_keys:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sid,
                        "startRowIndex": i + 1,
                        "endRowIndex": i + 2,
                        "startColumnIndex": 0,
                        "endColumnIndex": 8,
                    },
                    "cell": {"userEnteredFormat": {"textFormat": {"foregroundColor": red}}},
                    "fields": "userEnteredFormat.textFormat.foregroundColor",
                }
            })

    # Send in chunks
    CHUNK = 50
    for start in range(0, len(requests), CHUNK):
        wb_gs.batch_update({"requests": requests[start:start + CHUNK]})
        _time.sleep(2)
    print(f"  ✅ Formatting applied ({len(new_month_keys)} rows in red font).")


# ---------------------------------------------------------------------------
# Update: Unit Conditions — running totals
# ---------------------------------------------------------------------------

def update_unit_conditions(ws, unit_totals, dry_run=False):
    """Write per-unit running totals into 'Total Amount Spent on Unit Turn' column."""
    all_vals = read_sheet_as_list(ws)
    if not all_vals:
        return

    headers = all_vals[0]
    unit_col  = header_index(headers, "Unit")
    total_col = header_index(headers, "Total Amount Spent on Unit Turn")

    if unit_col == -1:
        print(f"  [WARN] 'Unit' column not found in {COND_SHEET}")
        return
    if total_col == -1:
        print(f"  [WARN] 'Total Amount Spent on Unit Turn' column not found in {COND_SHEET}")
        return

    updates = 0
    batch = []
    for i, row in enumerate(all_vals[1:], start=2):
        u = norm_unit(row[unit_col]) if unit_col < len(row) else ""
        if u and u in unit_totals:
            col_letter = col_num_to_letter(total_col + 1)
            batch.append({
                "range": f"{col_letter}{i}",
                "values": [[int(unit_totals[u]) if unit_totals[u] == int(unit_totals[u]) else unit_totals[u]]]
            })
            updates += 1

    print(f"\n[{COND_SHEET}]")
    print(f"  Units to update: {updates}")
    if dry_run or not batch:
        return

    ws.batch_update(batch, value_input_option="USER_ENTERED")
    print(f"  ✅ Updated {updates} unit totals.")


# ---------------------------------------------------------------------------
# Update: Quarterly Turn Cost Summary
# ---------------------------------------------------------------------------

def update_quarterly_summary(ws, all_ut_rows, dry_run=False):
    """
    Rebuild Quarterly Turn Cost Summary in-place.
    Columns: Unit | [quarter cols] | Total | Notes
    """
    # Aggregate unit-quarter totals from all Unit Turn rows
    uq_totals    = defaultdict(float)
    unit_totals  = defaultdict(float)

    for r in all_ut_rows:
        unit  = norm_unit(r[0]) if isinstance(r, list) else r.get("unit", "")
        inv_d = to_date(r[2]) if isinstance(r, list) else r.get("inv_date")
        amt   = norm_amt(r[5]) if isinstance(r, list) else float(r.get("amount", 0) or 0)
        q     = quarter_of(inv_d)
        if unit and q:
            uq_totals[f"{unit}|{q}"] += amt
            unit_totals[unit] += amt

    quarters = sorted(
        {k.split("|")[1] for k in uq_totals},
        key=lambda q: (int(q[:4]), int(q[5]))
    )
    units_with_data = set(unit_totals.keys())

    all_vals = read_sheet_as_list(ws)
    if not all_vals:
        return

    headers = all_vals[0]
    unit_col  = header_index(headers, "Unit")
    total_col = header_index(headers, "Total")
    notes_col = header_index(headers, "Notes")

    if unit_col == -1:
        print(f"  [WARN] 'Unit' column not found in {PERQ_SHEET}")
        return

    # Map units to their existing row index
    unit_to_row = {}
    for i, row in enumerate(all_vals[1:], start=2):
        u = norm_unit(row[unit_col]) if unit_col < len(row) else ""
        if u:
            unit_to_row[u] = i

    # Ensure quarter columns exist
    existing_quarters = {h: i for i, h in enumerate(headers) if re.match(r"^\d{4}Q[1-4]$", h)}

    print(f"\n[{PERQ_SHEET}]")
    print(f"  Quarters in data  : {quarters}")
    print(f"  Units with amounts: {len(units_with_data)}")

    if dry_run:
        return

    batch = []

    for q in quarters:
        if q not in existing_quarters:
            # Insert column before Total
            insert_at = total_col + 1 if total_col >= 0 else len(headers) + 1
            ws.insert_cols([[q]], col=insert_at)
            # Refresh after column insert
            all_vals = read_sheet_as_list(ws)
            headers  = all_vals[0]
            unit_col  = header_index(headers, "Unit")
            total_col = header_index(headers, "Total")
            notes_col = header_index(headers, "Notes")
            existing_quarters = {h: i for i, h in enumerate(headers) if re.match(r"^\d{4}Q[1-4]$", h)}
            unit_to_row = {}
            for i, row in enumerate(all_vals[1:], start=2):
                u = norm_unit(row[unit_col]) if unit_col < len(row) else ""
                if u:
                    unit_to_row[u] = i

    # Write per-unit quarter values and totals
    for unit, row_i in unit_to_row.items():
        for q, col_i in existing_quarters.items():
            val = uq_totals.get(f"{unit}|{q}", "")
            if val:
                batch.append({
                    "range": f"{col_num_to_letter(col_i + 1)}{row_i}",
                    "values": [[int(val) if val == int(val) else val]]
                })

        if total_col >= 0:
            tot = unit_totals.get(unit, "")
            if tot:
                batch.append({
                    "range": f"{col_num_to_letter(total_col + 1)}{row_i}",
                    "values": [[int(tot) if tot == int(tot) else tot]]
                })

    if batch:
        ws.batch_update(batch, value_input_option="USER_ENTERED")

    print(f"  ✅ Updated quarterly summary.")


def col_num_to_letter(n):
    """Convert 1-based column index to A, B, ..., Z, AA, ..."""
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


# ---------------------------------------------------------------------------
# Analysis Report
# ---------------------------------------------------------------------------

# ── Cross-check keyword maps ─────────────────────────────────────────────────

# Keywords in meeting notes → expected invoice category
NOTES_WORK_KEYWORDS = {
    "paint":        "Paint",
    "painting":     "Paint",
    "repaint":      "Paint",
    "resurface":    "Resurfacing",
    "resurfacing":  "Resurfacing",
    "tub resurface": "Resurfacing",
    "carpet":       "Flooring",
    "flooring":     "Flooring",
    "floor":        "Flooring",
    "vinyl plank":  "Flooring",
    "lvp":          "Flooring",
    "plank":        "Flooring",
    "hvac":         "HVAC",
    "ac":           "HVAC",
    "air condition": "HVAC",
    "appliance":    "Appliances",
    "stove":        "Appliances",
    "refrigerator": "Appliances",
    "fridge":       "Appliances",
    "dishwasher":   "Appliances",
    "microwave":    "Appliances",
    "cabinet":      "Cabinets/Counters",
    "counter":      "Cabinets/Counters",
    "countertop":   "Cabinets/Counters",
    "plumbing":     "Plumbing",
    "leak":         "Plumbing",
    "water heater": "Plumbing",
    "toilet":       "Plumbing",
    "electrical":   "Electrical",
    "electric":     "Electrical",
    "outlet":       "Electrical",
    "blind":        "Blinds/Window Treatments",
    "window treatment": "Blinds/Window Treatments",
    "clean":        "Cleaning",
    "cleaning":     "Cleaning",
    "mold":         "Biohazard/Specialty",
    "roach":        "Biohazard/Specialty",
    "pest":         "Biohazard/Specialty",
    "bug":          "Biohazard/Specialty",
    "door":         "Doors",
    "water damage": "Water Mitigation",
    "water mitigation": "Water Mitigation",
    "make ready":   "Make Ready",
    "make-ready":   "Make Ready",
}

# Keywords that signal a scope/complexity claim in notes
EASY_TURN_PHRASES = [
    "easy turn", "simple turn", "quick turn", "light turn",
    "minimal work", "just clean", "just paint", "just resurface",
    "good condition", "great condition", "minor work", "small turn",
    "cosmetic only", "cosmetic",
]
MAJOR_TURN_PHRASES = [
    "major turn", "heavy turn", "extensive", "full gut", "full renovation",
    "significant damage", "major damage", "complete turn", "full turn",
    "full rehab", "needs everything",
]

EASY_TURN_THRESHOLD  = 1500   # spending over this on an "easy turn" is a flag
WATCH_TURN_THRESHOLD = 3000   # spending over this on a "watch" unit is a flag

# Post-premium/renovation: flag if this much NEW spend accumulates after the note was written
POST_PREMIUM_THRESHOLD = 2000

# Durable work: category -> minimum months before it should be billed again
DURABLE_WORK_MIN_MONTHS = {
    "Flooring":           24,   # carpet/plank/LVP should last 2+ years
    "Resurfacing":        24,   # tub resurface should last 2+ years
    "Cabinets/Counters":  36,   # cabinets/countertops should last 3+ years
    "HVAC":               36,
    "Appliances":         36,
    "Electrical":         36,
    "Doors":              24,
    "Structural/Exterior": 36,
}

# Keywords that identify a unit as Premium or Renovated in Quarterly notes
PREMIUM_KEYWORDS = ["premium", "renovated", "renovation", "upgraded", "upgrade"]
SPECIAL_CIRCUMSTANCE_KEYWORDS = [
    "biohazard", "storm", "flood", "fire", "medical", "legal",
    "mold remediation", "insurance", "emergency",
]


def _read_quarterly_notes(wb):
    """
    Read 'Notes by Kryziel' column from Quarterly Turn Cost Summary.
    Returns dict: unit -> note_text (lowercase).
    """
    try:
        ws   = wb.worksheet(PERQ_SHEET)
        vals = ws.get_all_values()
    except Exception:
        return {}
    if not vals:
        return {}
    headers = [h.strip().lower() for h in vals[0]]
    unit_col  = next((i for i, h in enumerate(headers) if h == "unit"), 0)
    notes_col = next((i for i, h in enumerate(headers) if "notes by" in h), -1)
    if notes_col == -1:
        notes_col = next((i for i, h in enumerate(headers) if "notes" in h), -1)
    if notes_col == -1:
        return {}
    result = {}
    for row in vals[1:]:
        unit = norm_unit(row[unit_col]) if unit_col < len(row) else ""
        note = row[notes_col].strip() if notes_col < len(row) else ""
        if unit and note:
            result[unit] = note
    return result


def _classify_quarterly_note(note_text):
    """
    Returns tuple: (is_premium, is_special_circumstance, raw_note)
    based on keywords in the Quarterly notes.
    """
    lower = note_text.lower()
    is_premium = any(kw in lower for kw in PREMIUM_KEYWORDS)
    is_special = any(kw in lower for kw in SPECIAL_CIRCUMSTANCE_KEYWORDS)
    return is_premium, is_special


def _detect_repeat_durable_work(unit, all_ut_rows, target_month, inv_cat_map):
    """
    For a given unit, check if any durable-category work billed this month
    was also billed within the minimum lookback window.
    Returns list of flag strings.
    """
    target_yr  = target_month.year
    target_mon = target_month.month

    # Collect all invoices for this unit with dates and categories
    unit_history = []  # list of (date, category, desc, amount)
    for row in all_ut_rows:
        u = norm_unit(row[0]) if len(row) > 0 else ""
        if u != unit:
            continue
        inv_num = row[1].strip() if len(row) > 1 else ""
        raw_d   = row[3] if len(row) > 3 else ""
        desc    = row[4].strip() if len(row) > 4 else ""
        amt     = norm_amt(row[5]) if len(row) > 5 else 0
        cat     = inv_cat_map.get(inv_num.upper(), "Other")
        d = None
        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%-m/%-d/%Y", "%-m/%-d/%y", "%Y-%m-%d"):
            try:
                d = datetime.strptime(raw_d, fmt)
                break
            except ValueError:
                continue
        if d:
            unit_history.append((d, cat, desc, amt))

    # Separate this month vs prior history
    this_month_cats = {}  # category -> (date, desc, amt)
    prior_history   = []  # (date, cat, desc, amt)
    for d, cat, desc, amt in unit_history:
        if d.year == target_yr and d.month == target_mon:
            if cat not in this_month_cats:
                this_month_cats[cat] = (d, desc, amt)
        else:
            prior_history.append((d, cat, desc, amt))

    flags = []
    for cat, min_months in DURABLE_WORK_MIN_MONTHS.items():
        if cat not in this_month_cats:
            continue
        this_d, this_desc, this_amt = this_month_cats[cat]
        # Find most recent prior invoice in same category
        prior_same = [(d, desc, amt) for d, c, desc, amt in prior_history if c == cat]
        if not prior_same:
            continue
        prior_same.sort(key=lambda x: x[0], reverse=True)
        last_d, last_desc, last_amt = prior_same[0]
        months_gap = (this_d.year - last_d.year) * 12 + (this_d.month - last_d.month)
        if months_gap < min_months:
            flags.append(
                f"Repeat {cat} within {months_gap}mo (min expected gap: {min_months}mo) — "
                f"prior: {last_d.strftime('%-m/%Y')} ({_fmt_amt(last_amt)} · {last_desc[:40]})"
            )
    return flags


def _detect_premium_cost_creep(unit, quarterly_note, all_ut_rows, target_month, inv_cat_map):
    """
    For Premium/Renovated units: calculate total spend AFTER the note was written.
    Flags if post-premium spend is creeping up significantly.
    Returns (context_label, creep_flags).
    """
    is_premium, is_special = _classify_quarterly_note(quarterly_note)
    if not is_premium:
        return None, []

    label = "Premium" if "premium" in quarterly_note.lower() else "Renovated"

    # Rough heuristic: assume the premium renovation happened in the quarter with
    # the highest single-quarter spend. Post-premium = everything after that quarter.
    unit_by_quarter = defaultdict(float)
    unit_by_quarter_d = {}
    for row in all_ut_rows:
        u = norm_unit(row[0]) if len(row) > 0 else ""
        if u != unit:
            continue
        raw_d = row[3] if len(row) > 3 else ""
        amt   = norm_amt(row[5]) if len(row) > 5 else 0
        d = None
        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%-m/%-d/%Y", "%-m/%-d/%y", "%Y-%m-%d"):
            try:
                d = datetime.strptime(raw_d, fmt)
                break
            except ValueError:
                continue
        if d:
            q = f"{d.year}Q{(d.month-1)//3+1}"
            unit_by_quarter[q] += amt
            if q not in unit_by_quarter_d or d < unit_by_quarter_d[q]:
                unit_by_quarter_d[q] = d

    if not unit_by_quarter:
        return label, []

    # Quarter with highest spend = likely the premium renovation quarter
    peak_quarter = max(unit_by_quarter, key=lambda q: unit_by_quarter[q])
    peak_d = unit_by_quarter_d.get(peak_quarter)

    # Sum spend in all quarters AFTER the peak quarter
    post_spend = 0.0
    for q, amt in unit_by_quarter.items():
        if q > peak_quarter:
            post_spend += amt

    creep_flags = []
    if post_spend >= POST_PREMIUM_THRESHOLD:
        creep_flags.append(
            f"{label} unit — post-renovation spend is {_fmt_amt(post_spend)} "
            f"(since {peak_quarter}). Monitor for ongoing cost creep."
        )

    return label, creep_flags


def _extract_month_notes(notes_text, target_month):
    """Extract only the note entries for target_month from the full notes string."""
    yr_short = str(target_month.year)[-2:]
    mo       = str(target_month.month)
    lines    = notes_text.splitlines()
    result   = []
    capturing = False
    for line in lines:
        stripped = line.strip()
        # Date header like "6/24/26" or "6/1/26"
        date_match = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{2,4})$', stripped)
        if date_match:
            line_mo, _, line_yr = date_match.groups()
            line_yr_short = line_yr[-2:]
            capturing = (line_mo == mo and line_yr_short == yr_short)
            continue
        if capturing and stripped:
            result.append(stripped)
    return " ".join(result).lower()


def _detect_expected_categories(notes_lower):
    """Return set of categories expected based on keywords found in notes."""
    expected = set()
    for kw, cat in NOTES_WORK_KEYWORDS.items():
        if kw in notes_lower:
            expected.add(cat)
    return expected


def _detect_scope_claim(notes_lower):
    """Return 'easy', 'major', or None based on scope language in notes."""
    for phrase in EASY_TURN_PHRASES:
        if phrase in notes_lower:
            return "easy"
    for phrase in MAJOR_TURN_PHRASES:
        if phrase in notes_lower:
            return "major"
    return None


def crosscheck_meeting_notes_vs_invoices(gc, target_month):
    """
    Step 8: Comprehensive cross-check of meeting notes vs actual invoices.
    Writes results to 'Cross-Check' tab and prints summary.
    """
    import time as _time
    print("\n[Cross-Check: Meeting Notes vs Invoices]")
    month_label = target_month.strftime("%B %Y")
    target_yr   = target_month.year
    target_mon  = target_month.month

    try:
        wb      = gc.open_by_key(SPREADSHEET_ID)
        ws_cond = wb.worksheet(COND_SHEET)
        ws_ut   = wb.worksheet(UT_SHEET)
    except Exception as e:
        print(f"  [WARN] Could not open sheets for cross-check: {e}")
        return

    # ── Read Unit Conditions ────────────────────────────────────────
    cond_vals = ws_cond.get_all_values()
    if not cond_vals:
        print("  [WARN] Unit Conditions sheet is empty.")
        return

    headers   = cond_vals[0]
    unit_col  = next((i for i, h in enumerate(headers) if "unit" in h.lower()), 0)
    issue_col = next((i for i, h in enumerate(headers) if "condition" in h.lower() or "issue" in h.lower()), 3)

    # unit → full notes text
    all_notes = {}
    for row in cond_vals[1:]:
        unit  = norm_unit(row[unit_col]) if unit_col < len(row) else ""
        notes = row[issue_col].strip() if issue_col < len(row) else ""
        if unit and notes:
            all_notes[unit] = notes

    # Which units had notes mentioning this month
    discussed_units = {}   # unit -> notes_for_this_month (lowercase)
    for unit, notes in all_notes.items():
        month_notes = _extract_month_notes(notes, target_month)
        if month_notes.strip():
            discussed_units[unit] = month_notes

    # ── Read Unit Turn invoices for this month ──────────────────────
    ut_vals = ws_ut.get_all_values()
    unit_invoices = defaultdict(list)   # unit -> list of {desc, category, amount, date}

    # Build invoice→category from DATA Invoice By Location
    try:
        ws_src = wb.worksheet("DATA Invoice By Location")
        src_vals = ws_src.get_all_values()
        inv_cat_map = {}
        for row in src_vals[1:]:
            inv_num = row[3].strip() if len(row) > 3 else ""
            gl      = row[5].strip() if len(row) > 5 else ""
            if inv_num:
                inv_cat_map[inv_num.upper()] = GL_CATEGORY.get(gl, "Other")
    except Exception:
        inv_cat_map = {}

    for row in ut_vals[1:]:
        unit    = norm_unit(row[0]) if len(row) > 0 else ""
        inv_num = row[1].strip() if len(row) > 1 else ""
        raw_d   = row[3] if len(row) > 3 else ""
        desc    = row[4].strip() if len(row) > 4 else ""
        amt     = norm_amt(row[5]) if len(row) > 5 else 0
        if not unit:
            continue
        d = None
        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%-m/%-d/%Y", "%-m/%-d/%y", "%Y-%m-%d"):
            try:
                d = datetime.strptime(raw_d, fmt)
                break
            except ValueError:
                continue
        if d and d.year == target_yr and d.month == target_mon:
            cat = inv_cat_map.get(inv_num.upper(), "Other")
            unit_invoices[unit].append({"desc": desc, "category": cat, "amount": amt})

    invoiced_units = set(unit_invoices.keys())

    # ── Read Quarterly notes (your context notes per unit) ──────────
    quarterly_notes = _read_quarterly_notes(wb)   # unit -> "Premium", "Renovated", etc.
    all_ut_rows = ut_vals[1:]

    # ── Analyse each unit ───────────────────────────────────────────
    findings = []

    all_units = sorted(set(list(discussed_units.keys()) + list(invoiced_units)),
                       key=lambda x: int(x) if x.isdigit() else 9999)

    flags_summary = {
        "clean": [], "no_invoice": [], "unmentioned": [],
        "easy_overspend": [], "scope_creep": [], "unexpected_cat": [],
        "repeat_durable": [], "premium_creep": [],
    }

    for unit in all_units:
        in_notes    = unit in discussed_units
        in_invoice  = unit in invoiced_units
        notes_text  = discussed_units.get(unit, "")
        invoices    = unit_invoices.get(unit, [])
        total_spend = sum(i["amount"] for i in invoices)
        actual_cats = set(i["category"] for i in invoices)
        q_note      = quarterly_notes.get(unit, "")

        expected_cats = _detect_expected_categories(notes_text) if notes_text else set()
        scope_claim   = _detect_scope_claim(notes_text) if notes_text else None

        # Classify quarterly context note
        is_premium, is_special = _classify_quarterly_note(q_note) if q_note else (False, False)
        premium_label, creep_flags = _detect_premium_cost_creep(
            unit, q_note, all_ut_rows, target_month, inv_cat_map
        ) if q_note else (None, [])

        # Repeat durable work check (all-time history)
        repeat_flags = _detect_repeat_durable_work(
            unit, all_ut_rows, target_month, inv_cat_map
        ) if in_invoice else []

        flags = []

        if in_notes and in_invoice:
            status = "✅ Match"

            # Easy turn but high spend
            if scope_claim == "easy" and total_spend > EASY_TURN_THRESHOLD:
                flags.append(f"Easy turn claimed but spent {_fmt_amt(total_spend)}")
                flags_summary["easy_overspend"].append(unit)

            # Unexpected categories billed
            if expected_cats:
                unexpected = actual_cats - expected_cats - {"Other", "Make Ready", "Cleaning"}
                if unexpected:
                    flags.append(f"Unexpected work billed: {', '.join(sorted(unexpected))}")
                    flags_summary["unexpected_cat"].append(unit)

            # Expected work with no invoice
            missing_cats = expected_cats - actual_cats
            if missing_cats:
                flags.append(f"Expected but no invoice for: {', '.join(sorted(missing_cats))}")
                flags_summary["scope_creep"].append(unit)

            if not flags and not repeat_flags and not creep_flags:
                flags_summary["clean"].append(unit)

        elif in_notes and not in_invoice:
            status = "⚠️ No Invoice"
            flags.append("Discussed in meeting but no invoice received")
            flags_summary["no_invoice"].append(unit)

        else:
            status = "🔍 Not Discussed"
            flags.append("Invoice received but unit never mentioned in meeting notes")
            flags_summary["unmentioned"].append(unit)

        # Repeat durable work flags (apply regardless of match status)
        if repeat_flags:
            flags.extend(repeat_flags)
            flags_summary["repeat_durable"].append(unit)

        # Premium/Renovated cost creep flags
        if creep_flags:
            flags.extend(creep_flags)
            flags_summary["premium_creep"].append(unit)

        # Context label from your quarterly notes
        context_label = premium_label or ("Special circumstance" if is_special else "—")
        if q_note and not premium_label and not is_special:
            context_label = q_note[:60]

        inv_desc = "; ".join(
            f"{i['category']} {_fmt_amt(i['amount'])}" for i in
            sorted(invoices, key=lambda x: -x["amount"])
        ) if invoices else "—"

        findings.append({
            "unit":           unit,
            "status":         status,
            "context":        context_label,
            "scope_claim":    scope_claim or "—",
            "notes_summary":  notes_text[:200] if notes_text else "—",
            "expected_work":  ", ".join(sorted(expected_cats)) if expected_cats else "—",
            "actual_invoices": inv_desc,
            "total_spend":    _fmt_amt(total_spend) if total_spend else "—",
            "flags":          " | ".join(flags) if flags else "OK",
        })

    # ── Write to Cross-Check sheet tab ─────────────────────────────
    try:
        try:
            ws_cc = wb.worksheet("Cross-Check")
            ws_cc.clear()
        except Exception:
            ws_cc = wb.add_worksheet("Cross-Check", rows=500, cols=12)

        header_row = [
            "Unit", "Status", "Your Context Note", "Scope Claim",
            "Notes Summary (this month)", "Expected Work (from notes)",
            "Actual Invoices Billed", "Total Spend", f"Flags — {month_label}",
        ]
        out_rows = [header_row] + [[
            f["unit"], f["status"], f["context"], f["scope_claim"],
            f["notes_summary"], f["expected_work"],
            f["actual_invoices"], f["total_spend"], f["flags"],
        ] for f in findings]

        ws_cc.update("A1", out_rows, value_input_option="USER_ENTERED")
        _time.sleep(1)

        sid = ws_cc.id
        NCOLS = 9
        fmt_requests = [
            {"repeatCell": {
                "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                          "startColumnIndex": 0, "endColumnIndex": NCOLS},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": {"red": 0.118, "green": 0.220, "blue": 0.408},
                    "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1},
                                  "bold": True},
                    "wrapStrategy": "WRAP",
                }},
                "fields": "userEnteredFormat(backgroundColor,textFormat,wrapStrategy)",
            }},
            {"updateSheetProperties": {
                "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }},
            {"repeatCell": {
                "range": {"sheetId": sid, "startRowIndex": 1,
                          "endRowIndex": len(out_rows) + 1,
                          "startColumnIndex": 0, "endColumnIndex": NCOLS},
                "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP"}},
                "fields": "userEnteredFormat.wrapStrategy",
            }},
        ]

        RED    = {"red": 1.0,  "green": 0.88, "blue": 0.88}
        AMBER  = {"red": 1.0,  "green": 0.95, "blue": 0.80}
        GREEN  = {"red": 0.85, "green": 0.93, "blue": 0.85}
        YELLOW = {"red": 1.0,  "green": 0.97, "blue": 0.80}
        ORANGE = {"red": 1.0,  "green": 0.91, "blue": 0.77}  # repeat durable / creep

        for i, f in enumerate(findings):
            row_idx   = i + 1
            has_flags = f["flags"] not in ("OK", "—")
            repeat    = any(unit == f["unit"] for unit in flags_summary["repeat_durable"])
            creep     = any(unit == f["unit"] for unit in flags_summary["premium_creep"])
            if "No Invoice" in f["status"]:
                bg = AMBER
            elif "Not Discussed" in f["status"]:
                bg = RED
            elif repeat or creep:
                bg = ORANGE
            elif has_flags:
                bg = YELLOW
            else:
                bg = GREEN
            fmt_requests.append({"repeatCell": {
                "range": {"sheetId": sid, "startRowIndex": row_idx,
                          "endRowIndex": row_idx + 1,
                          "startColumnIndex": 0, "endColumnIndex": NCOLS},
                "cell": {"userEnteredFormat": {"backgroundColor": bg}},
                "fields": "userEnteredFormat.backgroundColor",
            }})

        widths = [65, 120, 140, 90, 280, 180, 250, 85, 340]
        for col_i, w in enumerate(widths):
            fmt_requests.append({"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS",
                          "startIndex": col_i, "endIndex": col_i + 1},
                "properties": {"pixelSize": w},
                "fields": "pixelSize",
            }})

        wb.batch_update({"requests": fmt_requests})
        print(f"  ✅ Cross-Check tab written ({len(findings)} units)")
        _time.sleep(1)

    except Exception as e:
        print(f"  [WARN] Could not write Cross-Check tab: {e}")

    # ── Print summary ───────────────────────────────────────────────
    print(f"\n  Month: {month_label}  |  {len(all_units)} units reviewed")
    print(f"  ✅ Clean match           : {len(flags_summary['clean'])}")
    print(f"  ⚠️  No invoice received  : {len(flags_summary['no_invoice'])} — {flags_summary['no_invoice']}")
    print(f"  🔍 Not in meeting notes  : {len(flags_summary['unmentioned'])} — {flags_summary['unmentioned']}")
    print(f"  🚨 Easy turn overspend   : {len(flags_summary['easy_overspend'])} — {flags_summary['easy_overspend']}")
    print(f"  📦 Unexpected categories : {len(flags_summary['unexpected_cat'])} — {flags_summary['unexpected_cat']}")
    print(f"  🔧 Expected work missing : {len(flags_summary['scope_creep'])} — {flags_summary['scope_creep']}")
    print(f"  🔁 Repeat durable work   : {len(flags_summary['repeat_durable'])} — {flags_summary['repeat_durable']}")
    print(f"  📈 Premium/cost creep    : {len(flags_summary['premium_creep'])} — {flags_summary['premium_creep']}")


def _fmt_amt(n):
    return f"${n:,.0f}" if n else "$0"


def print_analysis(new_rows, removed_rows, target_month, unit_totals_all):
    """Print a formatted analysis report to stdout."""
    bar = "=" * 60

    month_label = target_month.strftime("%B %Y")
    month_rows  = [
        r for r in new_rows
        if r["acctg_date"] and
           r["acctg_date"].year == target_month.year and
           r["acctg_date"].month == target_month.month
    ]

    # Per-unit cost this month
    month_unit_cost = defaultdict(float)
    for r in month_rows:
        month_unit_cost[r["unit"]] += r["amount"]

    # Duplicate invoices
    inv_counts = defaultdict(int)
    for r in month_rows:
        inv_counts[r["invoice"]] += 1
    dupes = {inv: cnt for inv, cnt in inv_counts.items() if cnt > 1}

    # High-cost units (all-time)
    high_cost = {u: t for u, t in unit_totals_all.items() if t > HIGH_COST_THRESHOLD}

    print(f"\n{bar}")
    print(f"  CREEKSIDE UNIT TURN — {month_label.upper()} ANALYSIS")
    print(f"{bar}")

    print(f"\n📋 INVOICE SUMMARY")
    print(f"  Total new invoices processed : {len(new_rows)}")
    print(f"  Added for {month_label:<12}    : {len(month_rows)}")
    print(f"  Removed (marketing/no-unit)  : {len(removed_rows)}")
    print(f"  Units active this month      : {len(month_unit_cost)}")

    if month_unit_cost:
        print(f"\n💰 COST BY UNIT — {month_label}")
        sorted_units = sorted(month_unit_cost.items(), key=lambda x: -x[1])
        for unit, amt in sorted_units:
            flag = "  ⚠️  HIGH" if amt > HIGH_COST_THRESHOLD else ""
            print(f"  Unit {unit:<6}  ${amt:>8,.0f}{flag}")
        print(f"  {'TOTAL':<10}  ${sum(month_unit_cost.values()):>8,.0f}")

    if dupes:
        print(f"\n🔁 DUPLICATE INVOICE FLAGS")
        for inv, cnt in dupes.items():
            print(f"  Invoice {inv} appears {cnt}x — please verify")

    if high_cost:
        print(f"\n🚨 HIGH-COST UNITS (All-Time > ${HIGH_COST_THRESHOLD:,})")
        for unit, tot in sorted(high_cost.items(), key=lambda x: -x[1]):
            print(f"  Unit {unit:<6}  ${tot:>8,.0f} cumulative")

    if removed_rows:
        print(f"\n🗑️  REMOVED ROWS ({len(removed_rows)} total)")
        mktg = [r for r in removed_rows if r["reason"] != "Missing / invalid unit number"]
        no_unit = [r for r in removed_rows if r["reason"] == "Missing / invalid unit number"]
        if mktg:
            print(f"  Marketing/non-turn ({len(mktg)}):")
            for r in mktg:
                print(f"    Unit {r['unit'] or '???'} | {r['invoice']} | {r['reason']}")
        if no_unit:
            print(f"  Missing unit # ({len(no_unit)}):")
            for r in no_unit:
                print(f"    Invoice {r['invoice']} | {r['description'][:50]}")

    print(f"\n{bar}\n")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    global HIGH_COST_THRESHOLD
    parser = argparse.ArgumentParser(
        description="Creekside Unit Turn — Monthly automation (processes CSVs → updates Google Sheet)"
    )
    parser.add_argument("--location", required=True,
                        help="Path to Invoice by Location CSV export from Resman")
    parser.add_argument("--detail",   required=True,
                        help="Path to Invoice by Detail CSV export from Resman")
    parser.add_argument("--month",    required=True,
                        help="Target month in YYYY-MM format (e.g. 2026-06)")
    parser.add_argument("--dry-run",  action="store_true",
                        help="Parse and report without writing to the sheet")
    parser.add_argument("--threshold", type=float, default=HIGH_COST_THRESHOLD,
                        help=f"High-cost alert threshold (default: ${HIGH_COST_THRESHOLD:,})")
    args = parser.parse_args()

    HIGH_COST_THRESHOLD = args.threshold

    try:
        target_month = datetime.strptime(args.month, "%Y-%m")
    except ValueError:
        print("ERROR: --month must be in YYYY-MM format, e.g. 2026-06")
        sys.exit(1)

    print(f"\nCreekside Unit Turn Automation")
    print(f"Target month : {target_month.strftime('%B %Y')}")
    print(f"Dry run      : {args.dry_run}")
    print(f"Location CSV : {args.location}")
    print(f"Detail CSV   : {args.detail}")

    # --- Parse inputs ---
    print("\n[Parsing CSVs...]")
    new_rows, removed_rows = parse_location_rows(args.location)
    detail_map = parse_detail_rows(args.detail)
    print(f"  Location rows (clean): {len(new_rows)}")
    print(f"  Location rows (removed): {len(removed_rows)}")
    print(f"  Detail invoice map: {len(detail_map)} entries")

    if args.dry_run:
        # Build unit totals from new rows only for report
        unit_totals_all = defaultdict(float)
        for r in new_rows:
            unit_totals_all[r["unit"]] += r["amount"]
        print_analysis(new_rows, removed_rows, target_month, unit_totals_all)
        print("(Dry run — no sheet changes made.)")
        return

    # --- Connect to sheet ---
    print("\n[Connecting to Google Sheets...]")
    gc = get_gspread_client()
    wb = gc.open_by_key(SPREADSHEET_ID)

    ws_src  = wb.worksheet(SRC_SHEET)
    ws_det  = wb.worksheet(DET_SHEET)
    ws_ut   = wb.worksheet(UT_SHEET)
    ws_cond = wb.worksheet(COND_SHEET)
    ws_perq = wb.worksheet(PERQ_SHEET)
    print("  ✅ Connected.")

    # --- Step 1: Update DATA Invoice by Location ---
    update_src_sheet(ws_src, new_rows, removed_rows, dry_run=False)

    # --- Step 2: Update Invoice by Detail ---
    update_det_sheet(ws_det, args.detail, dry_run=False)

    # --- Step 3: Update Unit Turn Cost by Unit ---
    update_unit_turn_sheet(ws_ut, new_rows, detail_map, target_month, wb_gs=wb, dry_run=False)

    # --- Step 4: Compute all-time unit totals for downstream steps ---
    ut_all_vals = read_sheet_as_list(ws_ut)
    unit_totals_all = defaultdict(float)
    for row in ut_all_vals[1:]:
        u   = norm_unit(row[0]) if len(row) > 0 else ""
        amt = norm_amt(row[5]) if len(row) > 5 else 0
        if u:
            unit_totals_all[u] += amt

    # --- Step 5: Update Unit Conditions ---
    update_unit_conditions(ws_cond, unit_totals_all, dry_run=False)

    # --- Step 6: Update Quarterly Turn Cost Summary ---
    update_quarterly_summary(ws_perq, ut_all_vals[1:], dry_run=False)

    # --- Step 7: Rebuild Analysis tab ---
    try:
        from build_analysis_tab import main as build_analysis
        build_analysis()
        print("\n✅ Analysis tab updated.")
    except Exception as e:
        print(f"\n[WARN] Analysis tab not updated: {e}")

    # --- Step 8: Cross-check meeting notes vs invoices ---
    crosscheck_meeting_notes_vs_invoices(gc, target_month)

    # --- Step 9: Rebuild interactive dashboard ---
    try:
        from build_dashboard import main as build_dashboard
        build_dashboard()
        print("\n✅ Dashboard rebuilt.")
    except Exception as e:
        print(f"\n[WARN] Dashboard not rebuilt: {e}")

    # --- Analysis Report ---
    print_analysis(new_rows, removed_rows, target_month, unit_totals_all)


if __name__ == "__main__":
    main()
