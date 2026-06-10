# WO Content Check

Verifying daily Word work orders.

The tool checks today's WO `.docx` files against the Part Info database and the Work Order data in quickbooks system, then prints a concise report where the real errors are visible immediately.

## Critical vs Warning

Critical errors are the items that should stop or fix the WO:

- Sales order row count mismatch
- Sales order item name mismatch
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

When Google Sheet sales-order lookup is enabled, the checker compares both the number of WO item rows and the item names against the Open Sales Order `Item` values. Item order is ignored, so reordering rows will not create a critical error. Missing, extra, or different item names are reported in the critical section.

Override the WO folder:

```powershell
python check_today.py --folder "D:\path\to\Work Order 2026-06"
```

Only files created or modified on the target date are checked when scanning a folder.
