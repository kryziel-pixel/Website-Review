# Creekside Unit Turn Automation — Setup Guide

## One-Time Setup (do this once)

### 1. Install dependencies
```bash
pip install gspread google-auth
```

### 2. Create a Google Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create or select a project
3. Enable **Google Sheets API** and **Google Drive API**
4. Go to **IAM & Admin → Service Accounts → Create Service Account**
5. Give it a name (e.g. `creekside-automation`)
6. Click **Create and Continue**, skip optional steps, click **Done**
7. Click the service account → **Keys** → **Add Key** → **Create new key** → **JSON**
8. Save the downloaded JSON file as `service_account.json` (do NOT commit this file)

### 3. Share the Google Sheet with the service account

- Open the service account JSON — copy the `client_email` value
  (looks like: `creekside-automation@your-project.iam.gserviceaccount.com`)
- Open the [Creekside Unit Turn spreadsheet](https://docs.google.com/spreadsheets/d/1qVJ3Nz4LCgZVKlWMLP8i2xj8F0tzD6ujGDBtGhM6JhM/)
- Click **Share** → paste the service account email → set to **Editor** → Share

### 4. Set the environment variable

```bash
export GOOGLE_SERVICE_ACCOUNT_JSON=/path/to/service_account.json
```

Or store it in your `.env` file:
```
GOOGLE_SERVICE_ACCOUNT_JSON=/path/to/service_account.json
```

---

## Monthly Workflow

### Step 1 — Export from Resman (you do this)
1. Export **Invoice by Location** report → save as CSV  
   (columns needed: PropertyName, Unit#, ObjectType, Invoice#, Install Date, Acctg Date, Total, GL Acc Number, Description)
2. Export **Invoice by Detail** report → save as CSV  
   (columns needed: InvoiceNumber, InvoiceDate)

### Step 2 — Send files to Claude
Tell Claude:
> "Here are the Resman exports for June 2026: [attach files]. Please run the Creekside unit turn update."

Claude will:
- Parse and clean the data
- Filter out marketing items (Locator, Apartment List, Lift Lead, Move-in Fee, Renters Insurance)
- Remove rows with invalid/missing unit numbers
- Append new rows to DATA Invoice by Location and Invoice by Detail Report
- Update Unit Turn Cost by Unit (deduped, sorted, color-banded)
- Update Unit Conditions running totals
- Rebuild Quarterly Turn Cost Summary
- Print a full analysis report

### Step 3 — Review flags
Claude will highlight:
- ⚠️ Units exceeding $5,000 (adjustable)
- 🔁 Duplicate invoice numbers
- 🗑️ All removed rows with reasons

---

## Manual Run (without Claude)

```bash
python creekside_automation.py \
    --location  "Invoice_by_Location_Jun2026.csv" \
    --detail    "Invoice_by_Detail_Jun2026.csv" \
    --month     2026-06
```

### Dry run (no sheet changes — just see the report):
```bash
python creekside_automation.py \
    --location  "Invoice_by_Location_Jun2026.csv" \
    --detail    "Invoice_by_Detail_Jun2026.csv" \
    --month     2026-06 \
    --dry-run
```

### Custom cost threshold:
```bash
python creekside_automation.py ... --threshold 8000
```

---

## Marketing Filter Rules

The following are automatically removed and logged:

| Category | Keywords |
|---|---|
| Locator commission | LOCATOR COMMISSION, LOCATOR |
| Apartment List | APARTMENT LIST, APARTMENTLIST, APARTMENTLIST.COM |
| Lead expenses | LEAD EXPENSE, LEADS EXPENSE |
| Lift Lead | LIFT LEAD, LIFTLEAD |
| Move-in Fee | MOVE-IN FEE, MOVE IN FEE, MOVEIN FEE |
| Renters Insurance | RENTERS INSURANCE, RENTER'S INSURANCE, RENTER INSURANCE |

To add new keywords, edit `MARKETING_KEYWORDS` and `MARKETING_REASONS` in `creekside_automation.py`.
