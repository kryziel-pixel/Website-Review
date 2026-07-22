# Oaks at Creekside — Automation Project

## What This Is
A Python-based automation system for the Oaks at Creekside apartment property. It pulls invoice and unit data from Google Sheets, runs monthly cross-check analysis, tracks unit conditions from weekly meeting emails, and generates an HTML dashboard for the asset management team.

## Credentials Setup
All scripts authenticate to Google Sheets via a service account.

**Required before running anything:**
```bash
export GOOGLE_SERVICE_ACCOUNT_JSON=/root/.credentials/creekside_service_account.json
```

The credentials file must exist at that path. It is NOT in the repo (never commit it). If setting up a new environment, upload the `creekside_service_account.json` file to `/root/.credentials/` and set the env var above.

## Google Sheet
- **Spreadsheet ID:** `1qVJ3Nz4LCgZVKlWMLP8i2xj8F0tzD6ujGDBtGhM6JhM`
- **Key tabs used:**
  - `DATA Invoice By Location` — raw invoice source data
  - `Unit Turn Cost by Unit` — monthly invoice rows per unit (written by automation)
  - `Quarterly Turn Cost Summary` — quarterly rollup with "Notes by Kryziel" column (col F) — DO NOT overwrite this column, it contains manual context notes
  - `Unit Conditions` — unit physical condition notes (col D = Condition/Issue)
  - `Cross-Check` — monthly cross-check findings (written by automation)

## Scripts

### `creekside_automation.py` — Monthly Update
Runs the full monthly automation. Run once per month after the new invoice data is in the sheet.

```bash
python creekside_automation.py
```

**What it does (in order):**
1. Reads `DATA Invoice By Location` for the target month
2. Deduplicates against existing `Unit Turn Cost by Unit` rows
3. Appends new invoice rows (red font for new entries)
4. Updates `Quarterly Turn Cost Summary` totals
5. Writes the `Cross-Check` tab with color-coded findings

**Cross-check colors:**
- GREEN = clean, no flags
- YELLOW = flags worth reviewing
- AMBER = unit discussed but no invoice found
- RED = invoice found but unit not discussed in meeting
- ORANGE = repeat durable work too soon, or post-premium cost creep

**Cross-check logic reads:**
- `Unit Conditions` tab for this month's meeting notes per unit
- `Quarterly Turn Cost Summary` col F ("Notes by Kryziel") for context — Premium/Renovated labels suppress cost flags but still monitor for post-renovation creep
- Invoice history for durable work repeat detection (e.g. flooring shouldn't repeat within 24 months)

**Durable work minimum thresholds:**
- Flooring / Resurfacing: 24 months
- Cabinets/Counters, HVAC, Appliances, Electrical, Doors, Structural: 36 months

### `update_unit_conditions.py` — Weekly Meeting Notes
Run after each weekly "Oaks at Creekside Operations Update Meeting Minutes" email.

**Typical usage — tell Claude:**
> "Update unit conditions from this week's meeting"

Claude will fetch the latest meeting email from Gmail and call:
```python
from update_unit_conditions import run_from_email_body
run_from_email_body(email_body, "7/8/26", dry_run=False)
```

**What it writes:** Only physical condition details about the unit itself — damage, repairs needed, leaks, pests, appliance issues, flooring, cleaning needed, etc.

**What it skips (does NOT write):**
- Evictions, court dates
- Skips / contractor unavailability / family emergencies
- Pure make-ready scheduling status (started / completed / not started) with no condition detail
- Tenancy or leasing notes

After writing, it automatically rebuilds the dashboard.

### `build_dashboard.py` — Dashboard Generator
Reads live data from Google Sheets and regenerates the HTML dashboard.

```bash
python build_dashboard.py
# or use the one-click script:
bash dashboard.sh
```

Output: `creekside_dashboard.html` (self-contained, ~320KB)

**Dashboard tabs:**
- Overview — KPI summary cards + charts
- Units — all units with conditions badge and cross-check flag count
- Categories — spend by GL category
- Vendors — spend by vendor
- All Invoices — full invoice table with filters

**Unit drill-down shows:** KPIs, charts, Cross-Check panel (color-coded flags), Meeting Notes history, full invoice table.

### `repair_unit_turn.py` — One-Time Dedup Fix
One-time script used to clean duplicate rows in `Unit Turn Cost by Unit`. Already run — do not re-run unless there's a new deduplication issue.

### `build_analysis_tab.py`
Builds additional analysis tab in the sheet. Run as needed.

## Weekly Workflow (Every Tuesday after meeting email arrives)

1. Open Claude Code
2. Say: **"Update unit conditions from this week's meeting"**
3. Claude fetches the latest Gmail meeting email, extracts physical condition notes, writes them to `Unit Conditions` tab col D, and rebuilds the dashboard
4. Claude sends you the updated `creekside_dashboard.html`
5. Review and confirm

## Monthly Workflow (After new invoice data is loaded into the sheet)

1. Open Claude Code
2. Say: **"Run the monthly Creekside update for [Month Year]"**
3. Claude runs `creekside_automation.py`, which:
   - Appends new invoice rows to `Unit Turn Cost by Unit`
   - Updates `Quarterly Turn Cost Summary`
   - Writes `Cross-Check` tab findings
   - Rebuilds the dashboard
4. Claude sends you the updated `creekside_dashboard.html`
5. Review the Cross-Check tab for flags — RED and ORANGE flags need attention
6. Add context notes to col F of `Quarterly Turn Cost Summary` as needed ("Premium", "Renovated", "Biohazard", etc.)

## Important Rules
- **Never commit** `creekside_service_account.json` or any credentials file
- **Never overwrite** col F ("Notes by Kryziel") in `Quarterly Turn Cost Summary` — the system reads these notes for context but must not replace them
- The `Unit Conditions` tab structure should remain as-is; the system only writes to col D
- Always push changes to branch `claude/creekside-unit-turn-automation-hykws7` (or the current active branch)

