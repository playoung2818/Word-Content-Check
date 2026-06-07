# WO Content Check

Proofreader for daily Word work orders.

The tool checks today's WO `.docx` files against the receiving-log database and the `Open Sales Order` Google Sheet, then prints a concise report where the real errors are visible immediately.

## Critical vs Warning

Critical errors are the items that should stop or fix the WO:

- Sales order row count mismatch
- Part number mismatch
- Quantity mismatch

Everything else is reported as a warning:

- Serial number not found in receiving log
- No serial numbers listed
- Missing/unreadable WO table rows
- Missing sales-order lookup columns
- Duplicate serial inside one WO
- Duplicate serial across checked WOs
- Old file modified today
- Suspicious quantity value

## Files

- `WO_Content_Check.ipynb`: notebook workflow for testing and ad hoc debugging
- `wo_check.py`: reusable checker logic
- `check_today.py`: CLI entry point for daily use

## Daily CLI Use

From this folder:

```powershell
python check_today.py
```

Example output:

```text
TODAY'S WO CHECK
Date checked: 2026-06-07
Checked WO: 7 | Critical: 3 | Warnings: 8 | WO with critical errors: 2 | Clean WO: 5

CRITICAL
1. WO06-20260520-Glen Bojsza.docx
   - Row count mismatch: Word 12 | Open Sales Order 0
   - Part mismatch x2: Q1901134, Q1901135 | Word SEMIL-2007 | DB SEMIL-2007-i7HC

WARNINGS
8 warnings hidden. SERIAL_NOT_FOUND: 5; NO_SN: 3
Use --show-warnings to print warning details.
```

The CLI exits with:

- `0`: no critical errors
- `1`: one or more critical errors found

## CLI Options

Check today:

```powershell
python check_today.py
```

Check yesterday:

```powershell
python check_today.py --days 1
```

Check one WO file:

```powershell
python check_today.py --file "D:\path\to\WO06-20260520-Test.docx"
```

Print warning details:

```powershell
python check_today.py --show-warnings
```

Export issue details to CSV:

```powershell
python check_today.py --export-csv report_output
```

Skip Google Sheet sales-order lookup:

```powershell
python check_today.py --no-sales-order
```

Override the WO folder:

```powershell
python check_today.py --folder "D:\path\to\Work Order 2026-06"
```

## Notebook Use

The notebook still works with the same daily cell:

```python
df_sales_order = read_sales_order()
wo_results = validate_wo_folder(folder=WO_FOLDER, sales_order=df_sales_order, days=0)
```

Detailed report data is stored in:

```python
display(wo_results["_report"]["critical"])
display(wo_results["_report"]["warnings"])
display(wo_results["_report"]["issues"])
```

## Configuration

Defaults are defined in `wo_check.py` and can be overridden with environment variables:

- `WO_DB_URL`
- `WO_GOOGLE_CRED_PATH`
- `WO_RECEIVING_LOG_PATH`
- `WO_FOLDER`
- `WO_SHEET_NAME`
- `WO_SALES_ORDER_TAB`

Current defaults:

- Google Sheet name: `PDF_WO`
- Sales order tab: `Open Sales Order`

## Data Sources

The checker uses:

- Word WO files from the configured WO folder
- PostgreSQL `receiving_log` table for serial-to-part lookup
- Google Sheet `Open Sales Order` tab for row-count comparison

Only files created or modified on the target date are checked when scanning a folder.
