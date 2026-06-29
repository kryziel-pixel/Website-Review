"""
One-time repair for Unit Turn Cost by Unit:
  1. Remove all duplicate rows (keep first occurrence per unit|invoice|desc)
  2. Recalculate unit totals and Duplicate? flags
  3. Re-apply per-unit pastel color banding
  4. Apply red FONT (not background) for June 2026 rows
"""

import os, sys, re, json, pickle, time
from datetime import datetime
from collections import defaultdict

SPREADSHEET_ID = "1qVJ3Nz4LCgZVKlWMLP8i2xj8F0tzD6ujGDBtGhM6JhM"
UT_SHEET       = "Unit Turn Cost by Unit"

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
RED_FONT = {"red": 1.0, "green": 0.0, "blue": 0.0}
WHITE_BG = {"red": 1.0, "green": 1.0, "blue": 1.0}


def s(v):
    return str(v or "").strip()

def norm_unit(v):
    digits = re.sub(r"[^0-9]", "", s(v))
    n = int(digits) if digits else 0
    return str(n) if 100 <= n <= 1500 else ""

def norm_inv(v):
    return s(v).upper()

def norm_desc(v):
    return re.sub(r"\s+", " ", s(v)).upper()

def norm_amt(v):
    try:
        return round(float(re.sub(r"[^0-9.\-]", "", s(v)) or "0"), 2)
    except ValueError:
        return 0.0

def dedup_key(unit, inv, desc):
    return f"{norm_unit(unit)}|{norm_inv(inv)}|{norm_desc(desc)}"

def to_date(v):
    if isinstance(v, datetime):
        return v
    raw = s(v)
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%-m/%-d/%Y", "%-m/%-d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def get_client():
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    json_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    raw_json  = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if raw_json:
        creds = Credentials.from_service_account_info(json.loads(raw_json), scopes=scopes)
    elif json_path:
        creds = Credentials.from_service_account_file(json_path, scopes=scopes)
    else:
        sys.exit("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_CREDENTIALS_JSON")
    return gspread.authorize(creds)


def col_letter(n):
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def main():
    # Load june keys so we know which rows to paint red
    june_keys = set()
    pkl_path = "/tmp/june_rows.pkl"
    if os.path.exists(pkl_path):
        with open(pkl_path, "rb") as f:
            _unit_rows, june_rows, _removed, _detail_map = pickle.load(f)
        for r in june_rows:
            june_keys.add(dedup_key(r["unit"], r["invoice"], r["description"]))
        print(f"Loaded {len(june_keys)} June 2026 keys from pickle.")
    else:
        print("WARNING: /tmp/june_rows.pkl not found — will not apply red font.")

    gc = get_client()
    wb = gc.open_by_key(SPREADSHEET_ID)
    ws = wb.worksheet(UT_SHEET)

    print("Reading sheet…")
    all_vals = ws.get_all_values()
    header = all_vals[0] if all_vals else []
    data_rows = all_vals[1:] if len(all_vals) > 1 else []
    print(f"  {len(data_rows)} data rows read.")

    # --- Deduplicate ---
    seen_keys = set()
    kept = []
    dropped = 0
    for row in data_rows:
        unit = norm_unit(row[0]) if len(row) > 0 else ""
        inv  = row[1] if len(row) > 1 else ""
        desc = row[4] if len(row) > 4 else ""
        k = dedup_key(unit, inv, desc)
        if k in seen_keys:
            dropped += 1
            continue
        seen_keys.add(k)
        kept.append(row)

    print(f"  {dropped} duplicate rows removed, {len(kept)} kept.")

    # --- Sort by unit then install date ---
    def sort_key(row):
        u = norm_unit(row[0]) if len(row) > 0 else ""
        d = to_date(row[3]) if len(row) > 3 else None
        return (int(u) if u.isdigit() else 9999, d or datetime.min)

    kept.sort(key=sort_key)

    # --- Recalculate unit totals ---
    unit_totals = defaultdict(float)
    for row in kept:
        u   = norm_unit(row[0]) if len(row) > 0 else ""
        amt = norm_amt(row[5]) if len(row) > 5 else 0
        if u:
            unit_totals[u] += amt

    # --- Rebuild output rows (8 cols) ---
    seen_units = set()
    out_rows = []
    for row in kept:
        unit  = norm_unit(row[0]) if len(row) > 0 else (row[0] if row else "")
        inv   = row[1] if len(row) > 1 else ""
        invd  = row[2] if len(row) > 2 else ""
        instd = row[3] if len(row) > 3 else ""
        desc  = row[4] if len(row) > 4 else ""
        amt   = row[5] if len(row) > 5 else ""

        # Unit total in first row of each unit
        if unit and unit not in seen_units:
            seen_units.add(unit)
            tot = unit_totals.get(unit, 0)
            unit_total_str = str(int(tot)) if tot == int(tot) else str(tot)
        else:
            unit_total_str = ""

        # Duplicate flag — with correct key, there should be none
        dup_str = ""  # will be empty since we already deduped

        out_rows.append([unit, inv, invd, instd, desc, amt, unit_total_str, dup_str])

    # --- Write data back to sheet ---
    print("Writing deduplicated data to sheet…")
    header_row = ["Unit", "Invoice #", "Invoice Date", "Install Date",
                  "Description", "Amount", "Unit Total", "Duplicate?"]

    # Clear entire sheet first
    total_rows = max(len(data_rows) + 5, len(out_rows) + 5)
    ws.batch_clear([f"A1:H{total_rows}"])
    time.sleep(2)

    ws.update("A1", [header_row], value_input_option="USER_ENTERED")
    time.sleep(1)

    if out_rows:
        ws.update("A2", out_rows, value_input_option="USER_ENTERED")
        print(f"  Written {len(out_rows)} rows.")
    time.sleep(3)

    # --- Apply formatting ---
    print("Applying formatting (pastel bands + red font for June rows)…")

    requests = []

    # Clear all existing formatting on data area
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": 1,
                "endRowIndex": len(out_rows) + 2,
                "startColumnIndex": 0,
                "endColumnIndex": 8,
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": WHITE_BG,
                    "textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}, "bold": False},
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    })

    # Collect rows by unit for banding
    unit_groups = defaultdict(list)  # unit -> list of (sheet_row_index 1-based in data, row)
    for i, row in enumerate(out_rows):
        unit = row[0]
        unit_groups[unit].append(i)  # 0-based index in out_rows

    # Assign pastel per unique unit (in order of first appearance)
    seen_order = []
    seen_set = set()
    for row in out_rows:
        u = row[0]
        if u not in seen_set:
            seen_set.add(u)
            seen_order.append(u)

    unit_color = {u: PASTEL_COLORS[i % len(PASTEL_COLORS)] for i, u in enumerate(seen_order)}

    # Build pastel background requests (group consecutive rows of same unit)
    current_unit = None
    group_start = None
    for i, row in enumerate(out_rows):
        u = row[0]
        if u != current_unit:
            if current_unit is not None and group_start is not None:
                color = unit_color.get(current_unit, WHITE_BG)
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": group_start + 1,  # +1 for header
                            "endRowIndex": i + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 8,
                        },
                        "cell": {"userEnteredFormat": {"backgroundColor": color}},
                        "fields": "userEnteredFormat.backgroundColor",
                    }
                })
            current_unit = u
            group_start = i
    # Last group
    if current_unit is not None and group_start is not None:
        color = unit_color.get(current_unit, WHITE_BG)
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": group_start + 1,
                    "endRowIndex": len(out_rows) + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 8,
                },
                "cell": {"userEnteredFormat": {"backgroundColor": color}},
                "fields": "userEnteredFormat.backgroundColor",
            }
        })

    # Red font for June 2026 rows
    red_count = 0
    for i, row in enumerate(out_rows):
        unit = row[0]
        inv  = row[1]
        desc = row[4]
        k = dedup_key(unit, inv, desc)
        if k in june_keys:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": i + 1,  # +1 for header
                        "endRowIndex": i + 2,
                        "startColumnIndex": 0,
                        "endColumnIndex": 8,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {"foregroundColor": RED_FONT}
                        }
                    },
                    "fields": "userEnteredFormat.textFormat.foregroundColor",
                }
            })
            red_count += 1

    print(f"  {red_count} June 2026 rows will get red font.")

    # Send formatting in chunks to avoid API limits
    CHUNK = 50
    for start in range(0, len(requests), CHUNK):
        chunk = requests[start:start + CHUNK]
        wb.batch_update({"requests": chunk})
        print(f"  Sent formatting batch {start // CHUNK + 1}/{(len(requests) - 1) // CHUNK + 1}")
        time.sleep(2)

    print("\n✅ Repair complete.")
    print(f"   Rows in sheet : {len(out_rows)}")
    print(f"   Duplicates removed: {dropped}")
    print(f"   June 2026 rows (red font): {red_count}")


if __name__ == "__main__":
    main()
