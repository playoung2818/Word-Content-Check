# WO Content Check

Template README for introducing [`WO_Content_Check.ipynb`](/D:/OneDrive%20-%20neousys-tech/Desktop/Python/WO_Content_Check/WO_Content_Check.ipynb).

## Overview

`WO_Content_Check.ipynb` is a working notebook for validating and cross-checking work order content against operational data sources such as:

- Google Sheets open sales order data
- Receiving log data stored in Supabase/PostgreSQL
- Work order files in Word and PDF format
- Warehouse inventory snapshots and daily update files

Use this section to explain the business problem clearly:

> This notebook is used to review work orders, compare document content with source records, and identify mismatches before downstream warehouse or fulfillment processing.

## What This Notebook Does

Based on the current notebook structure, the workflow includes:

1. Loading configuration and Python dependencies
2. Connecting to Supabase through Flask and SQLAlchemy
3. Reading open sales orders from a Google Sheet
4. Loading receiving log data into a database table
5. Processing work order files for the current day
6. Running ad hoc debugging and individual WO checks
7. Parsing PDF-based work orders
8. Reordering or reconciling output dataframes for validation

If needed, replace this with a more precise workflow description for your team.

## Notebook Structure

Current sections in the notebook:

- `Config`
- `Load Open Sales Order From Google Sheet`
- `Load Receiving Log into DB`
- `Go Through Today's WOs`
- `Debug`
- `Go Through WO`
- `Update changes in excel to Supabase`
- `Parse PDF`

## Environment

Recommended environment:

- Python `3.10+` or your team standard
- Jupyter Notebook or VS Code Notebook support

Main libraries currently imported in the notebook:

- `flask`
- `flask_sqlalchemy`
- `pandas`
- `pdfplumber`
- `python-docx`
- `openpyxl`
- `numpy`

Example install command:

```powershell
pip install pandas numpy flask flask_sqlalchemy pdfplumber python-docx openpyxl
```

If your project uses a managed environment, replace this section with the correct setup steps.

## Required Inputs

Document the exact sources the notebook expects.

### External Data Sources

- Google Sheet: `PDF_WO`
- Worksheet: `Open Sales Order`
- Supabase/PostgreSQL database
- Shared warehouse folders containing:
  - work order Word files
  - work order PDF files
  - inventory check Excel files
  - daily update CSV files

### Credentials

The notebook currently references external credentials and database connection details. Move sensitive values to environment variables or a local config file that is not committed.

Example:

```powershell
$env:GOOGLE_APPLICATION_CREDENTIALS="C:\path\to\service-account.json"
$env:SUPABASE_DB_URL="postgresql://<user>:<password>@<host>:<port>/<db>"
```

## How To Use

1. Open [`WO_Content_Check.ipynb`](/D:/OneDrive%20-%20neousys-tech/Desktop/Python/WO_Content_Check/WO_Content_Check.ipynb).
2. Update all local file paths so they match your machine or shared drive mapping.
3. Confirm Google Sheets and database credentials are available.
4. Run the notebook section by section from top to bottom.
5. Review validation output, mismatch logs, and any generated tables or dataframes.

## Expected Output

Describe what a successful run produces. For example:

- validated work order records
- mismatch reports between WO content and source data
- parsed product details from Word or PDF work orders
- receiving log records loaded or updated in Supabase

## File Path Notes

This notebook currently uses machine-specific and shared-drive paths. Standardize these before broader team use.

Suggested improvement:

- move repeated paths into a single config cell
- use environment variables for credentials
- avoid hardcoded usernames and secrets

## Known Risks / Cleanup Items

Use this section to track technical debt before production use:

- hardcoded local paths
- hardcoded secrets or connection strings
- notebook cells that mix reusable logic with one-off debugging
- dependencies not captured in `requirements.txt`
- unclear expected output format for each validation step

## Recommended Repo Layout

If this notebook will stay in the repository, consider evolving toward:

```text
WO_Content_Check/
├─ WO_Content_Check.ipynb
├─ README.md
├─ requirements.txt
├─ .env.example
├─ src/
│  ├─ loaders.py
│  ├─ parsers.py
│  ├─ validators.py
│  └─ db.py
└─ data/
   └─ sample_files/
```

## Maintainer

- Owner: `[Name / Team]`
- Last reviewed: `[YYYY-MM-DD]`
- Contact: `[email or chat channel]`

## Revision Notes

- `[YYYY-MM-DD]` Initial README template created
