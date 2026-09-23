"""
Main execution script for GST Reconciliation
Run:
    python run_reconciliation.py
"""

import os
import sys
from pathlib import Path
from reconciler import (
    load_books_data,
    load_portal_data,
    GSTReconciler,
    export_reconciliation_workbook
)

def main():
    print("=" * 70)
    print("   GST RECONCILIATION ENGINE (PORTAL 2B vs SAP BOOKS)")
    print("   V. Singhi & Associates | Delhi")
    print("=" * 70)

    # File Paths (Defaults)
    books_file = r"C:\Users\lenovo\Downloads\Combined PR GST REPORT_06062023 Creation date upto 17-Apr-2026_Books.xlsx"
    portal_file = r"C:\Users\lenovo\Downloads\Bangalore IOT_GST portal.xlsx"
    branch_name = "Bangalore IOT"
    output_report = r"C:\Users\lenovo\Downloads\Bangalore_IOT_GST_Reconciliation_Report.xlsx"

    print(f"\n[1] Configuration:")
    print(f"    - SAP Books File : {books_file}")
    print(f"    - GST Portal File: {portal_file}")
    print(f"    - Branch Filter  : {branch_name}")
    print(f"    - Output Report  : {output_report}\n")

    if not os.path.exists(books_file):
        print(f"[!] Error: Books file not found at: {books_file}")
        sys.exit(1)

    if not os.path.exists(portal_file):
        print(f"[!] Error: Portal file not found at: {portal_file}")
        sys.exit(1)

    # 1. Load Data
    print("[2] Loading Books data (Filtering branch)...")
    books_rows = load_books_data(books_file, branch_filter=branch_name)
    print(f"    -> Extracted {len(books_rows)} lines from SAP Books for '{branch_name}'.")

    print("[3] Loading Portal data...")
    portal_rows = load_portal_data(portal_file)
    print(f"    -> Extracted {len(portal_rows)} lines from GST Portal.")

    # 2. Reconcile
    reconciler = GSTReconciler(books_rows, portal_rows, tolerance=2.00)
    reconciler.run_reconciliation()

    # 3. Export Excel Report
    print(f"[4] Exporting formatted reconciliation report...")
    try:
        export_reconciliation_workbook(reconciler, output_report)
    except PermissionError:
        output_report = r"C:\Users\lenovo\Downloads\Bangalore_IOT_GST_Reconciliation_Report_Updated.xlsx"
        print(f"[!] Target file was open in Excel. Saving to new path: {output_report}")
        export_reconciliation_workbook(reconciler, output_report)

    # 4. Summary Output
    print("\n" + "=" * 70)
    print("                 RECONCILIATION SUMMARY")
    print("=" * 70)
    total_books_tax = sum(r['total_tax'] for r in reconciler.books_rows)
    total_portal_tax = sum(r['total_tax'] for r in reconciler.portal_rows)
    matched_count = sum(1 for r in reconciler.books_recon if "Matched" in r['match_status'])
    matched_tax = sum(r['total_tax'] for r in reconciler.books_recon if "Matched" in r['match_status'])
    missing_books_count = sum(1 for r in reconciler.books_recon if r['match_status'] == "Only in Books (Missing in 2B)")
    missing_books_tax = sum(r['total_tax'] for r in reconciler.books_recon if r['match_status'] == "Only in Books (Missing in 2B)")
    unbooked_portal_count = sum(1 for r in reconciler.portal_recon if r['match_status'] == "Only in Portal (Unbooked)")
    unbooked_portal_tax = sum(r['total_tax'] for r in reconciler.portal_recon if r['match_status'] == "Only in Portal (Unbooked)")

    print(f"  * Total Books Tax (Bangalore IOT)   : Rs. {total_books_tax:,.2f}  ({len(reconciler.books_rows)} entries)")
    print(f"  * Total Portal Tax (2B)             : Rs. {total_portal_tax:,.2f}  ({len(reconciler.portal_rows)} entries)")
    print(f"  * Net Difference (Books - Portal)   : Rs. {(total_books_tax - total_portal_tax):,.2f}")
    print(f"  -------------------------------------------------------------")
    print(f"  * Matched Invoices (Safe ITC)       : {matched_count} entries | Rs. {matched_tax:,.2f}")
    print(f"  * Missing in 2B (Vendor Follow-up)  : {missing_books_count} entries | Rs. {missing_books_tax:,.2f}")
    print(f"  * Unbooked in Books (Unclaimed ITC) : {unbooked_portal_count} entries | Rs. {unbooked_portal_tax:,.2f}")
    print("=" * 70)
    print(f"[OK] Completed successfully! Open file to review:\n    {output_report}\n")

if __name__ == "__main__":
    main()
