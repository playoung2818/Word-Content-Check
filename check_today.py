from __future__ import annotations

import argparse
from pathlib import Path
import sys

from wo_check import (
    GOOGLE_CRED_PATH,
    SHEET_NAME,
    SALES_ORDER_TAB,
    WO_FOLDER,
    export_report,
    read_sales_order,
    validate_single_file,
    validate_wo_folder,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Check work order Word files and print a concise critical-first report.",
    )
    parser.add_argument("--folder", default=str(WO_FOLDER), help="WO folder to scan. Defaults to WO_FOLDER/env setting.")
    parser.add_argument("--days", type=int, default=0, help="0 checks today, 1 checks yesterday, etc.")
    parser.add_argument("--file", help="Check one specific Word .docx file instead of scanning the folder.")
    parser.add_argument("--no-sales-order", action="store_true", help="Skip Google Sheet row-count lookup.")
    parser.add_argument("--google-cred-path", default=str(GOOGLE_CRED_PATH), help="Google service-account JSON path.")
    parser.add_argument("--sheet-name", default=SHEET_NAME, help="Google Sheet name.")
    parser.add_argument("--sales-order-tab", default=SALES_ORDER_TAB, help="Open Sales Order worksheet/tab name.")
    parser.add_argument("--show-warnings", action="store_true", help="Print warning details after the critical summary.")
    parser.add_argument("--export-csv", help="Directory where issues/critical/warnings CSV files should be written.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    sales_order = None
    if not args.no_sales_order:
        sales_order = read_sales_order(
            google_cred_path=Path(args.google_cred_path),
            sheet_name=args.sheet_name,
            sales_order_tab=args.sales_order_tab,
        )

    if args.file:
        results = validate_single_file(args.file, sales_order=sales_order, show_warnings=args.show_warnings)
    else:
        results = validate_wo_folder(
            folder=Path(args.folder),
            sales_order=sales_order,
            days=args.days,
            show_warnings=args.show_warnings,
        )

    if args.export_csv:
        exported = export_report(results, args.export_csv)
        if exported:
            print()
            print("EXPORTED")
            for name, path in exported.items():
                print(f"{name}: {path}")

    critical = results.get("_report", {}).get("critical")
    return 1 if critical is not None and not critical.empty else 0


if __name__ == "__main__":
    sys.exit(main())
