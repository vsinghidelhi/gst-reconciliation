"""
GST Reconciliation Engine (GSTR-2B vs SAP Books)
Developed for V. Singhi & Associates / vsinghidelhi
Author: AI Assistant
Description:
  Automates 2-Way Reconciliation between GST Portal (GSTR-2B) and SAP Purchase Register.
  Implements a 2-Check Architecture:
    - Check 1: Vendor GSTIN Pivot Macro Reconciliation
    - Check 2: Invoice & Amount Level Micro Reconciliation
"""

import os
import re
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# -------------------------------------------------------------
# 1. HELPER & CLEANING UTILITIES
# -------------------------------------------------------------

def clean_str(val):
    if val is None:
        return ""
    return str(val).strip()

def clean_gstin(val):
    if val is None:
        return ""
    s = str(val).strip().upper().replace(" ", "")
    # Remove non-alphanumeric
    s = re.sub(r'[^A-Z0-9]', '', s)
    return s

def clean_invoice_no(val):
    if val is None:
        return ""
    s = str(val).strip().upper()
    # 1. Normalize Financial Year: 2025-26 -> 25-26, 2024-25 -> 24-25
    s = re.sub(r'20(\d{2})[-/]?(\d{2})', r'\1\2', s)
    # 2. Normalize single digits between separators: e.g. /5/ -> /05/, -4- -> -04-
    s = re.sub(r'([/-])([0-9])([/-])', r'\g<1>0\2\3', s)
    # 3. Remove all special characters, slashes, dashes, spaces, underscores
    s = re.sub(r'[^A-Z0-9]', '', s)
    # 4. Remove leading zeros
    s = s.lstrip('0')
    # 5. Handle accidental double paste: e.g. ABCABC -> ABC
    half = len(s) // 2
    if half >= 5 and s[:half] == s[half:]:
        s = s[:half]
    return s if s else "0"

def safe_float(val):
    if val is None or val == "":
        return 0.0
    try:
        return round(float(val), 2)
    except (ValueError, TypeError):
        return 0.0

def string_similarity(s1, s2):
    """Deep similarity checker for accounting invoice typos & abbreviations"""
    if not s1 or not s2:
        return 0.0
    if s1 == s2:
        return 1.0
    if s1 in s2 or s2 in s1:
        return 0.90
    # Token prefix + suffix match (e.g. ESIVPLTechBOT2509 vs ESIVPLTB2509)
    if len(s1) >= 6 and len(s2) >= 6 and s1[:4] == s2[:4] and s1[-4:] == s2[-4:]:
        return 0.88
    # 1 character substitution / typo (e.g. 0048 vs 004B)
    if len(s1) == len(s2) and sum(1 for x, y in zip(s1, s2) if x != y) <= 1:
        return 0.85
    # Common suffix/prefix (last 4 characters)
    if len(s1) >= 4 and len(s2) >= 4 and s1[-4:] == s2[-4:]:
        return 0.80
    return 0.0


# -------------------------------------------------------------
# 2. DATA LOADERS
# -------------------------------------------------------------

def load_books_data(file_path, branch_filter=None):
    """
    Loads SAP PR GST Report (Books) from 'Input' sheet.
    Filters by branch if branch_filter is provided (case-insensitive).
    """
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    if 'Input' not in wb.sheetnames:
        raise ValueError(f"Sheet 'Input' not found in Books file: {file_path}")
    sheet = wb['Input']

    rows = []
    for i, r in enumerate(sheet.iter_rows(values_only=True)):
        if i == 0:
            continue
        branch = clean_str(r[29])
        if branch_filter and branch.lower() != branch_filter.lower():
            continue

        raw_bill_no = clean_str(r[8])
        gstin = clean_gstin(r[11])
        cgst = safe_float(r[24])
        sgst = safe_float(r[25])
        igst = safe_float(r[26])
        taxable = safe_float(r[23])
        total_val = safe_float(r[27])

        row_dict = {
            'row_id': i,
            'doc_type': clean_str(r[1]),
            'trans_no': clean_str(r[2]),
            'sap_trans_no': clean_str(r[3]),
            'posting_date': clean_str(r[4])[:10] if r[4] else "",
            'inv_date': clean_str(r[5])[:10] if r[5] else "",
            'bill_no': raw_bill_no,
            'clean_bill_no': clean_invoice_no(raw_bill_no),
            'vendor_code': clean_str(r[9]),
            'vendor_name': clean_str(r[10]),
            'gstin': gstin,
            'rcm': clean_str(r[12]),
            'taxable': taxable,
            'cgst': cgst,
            'sgst': sgst,
            'igst': igst,
            'total_tax': round(cgst + sgst + igst, 2),
            'total_val': total_val,
            'branch': branch
        }
        rows.append(row_dict)
    wb.close()
    return rows


def load_portal_data(file_path):
    """
    Loads GST Portal / GSTR-2B data.
    Automatically detects column headers across different portal formats.
    """
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheet_names = wb.sheetnames

    # Detect active sheet
    target_sheet = None
    for name in ['Bangalore IOT Portal', 'Purchase', 'Overview', 'Sheet1']:
        if name in sheet_names:
            target_sheet = name
            break
    if not target_sheet:
        target_sheet = sheet_names[0]

    sheet = wb[target_sheet]
    rows = []
    col_idx = {}

    for i, r in enumerate(sheet.iter_rows(values_only=True)):
        if i == 0:
            for idx, col in enumerate(r):
                if col is not None:
                    col_idx[str(col).strip().lower()] = idx
            continue

        gst_idx = col_idx.get('gstin', col_idx.get('supplier gstin', 2))
        inv_idx = col_idx.get('invoice no', col_idx.get('doc no', 5))
        name_idx = col_idx.get('supplier name', 1)
        date_idx = col_idx.get('invoice date', col_idx.get('doc date', 7))
        taxable_idx = col_idx.get('taxable value', col_idx.get('item taxable value', 9))
        igst_idx = col_idx.get('igst', 10)
        cgst_idx = col_idx.get('cgst', 11)
        sgst_idx = col_idx.get('sgst', 12)
        val_idx = col_idx.get('invoice value', col_idx.get('doc value', 8))
        rc_idx = col_idx.get('rc', col_idx.get('reverse charge', -1))

        gstin = clean_gstin(r[gst_idx]) if gst_idx < len(r) else ""
        if not gstin and not any(r):
            continue

        bill_no = clean_str(r[inv_idx]) if inv_idx < len(r) else ""
        cgst = safe_float(r[cgst_idx]) if cgst_idx < len(r) else 0.0
        sgst = safe_float(r[sgst_idx]) if sgst_idx < len(r) else 0.0
        igst = safe_float(r[igst_idx]) if igst_idx < len(r) else 0.0
        taxable = safe_float(r[taxable_idx]) if taxable_idx < len(r) else 0.0
        val = safe_float(r[val_idx]) if val_idx < len(r) else 0.0

        row_dict = {
            'portal_id': i,
            'supplier_name': clean_str(r[name_idx]) if name_idx < len(r) else "",
            'gstin': gstin,
            'doc_no': bill_no,
            'clean_doc_no': clean_invoice_no(bill_no),
            'doc_date': clean_str(r[date_idx])[:10] if date_idx < len(r) and r[date_idx] else "",
            'doc_val': val,
            'taxable': taxable,
            'cgst': cgst,
            'sgst': sgst,
            'igst': igst,
            'total_tax': round(cgst + sgst + igst, 2),
            'rc': clean_str(r[rc_idx]) if rc_idx != -1 and rc_idx < len(r) else "N"
        }
        rows.append(row_dict)
    wb.close()
    return rows


# -------------------------------------------------------------
# 3. RECONCILIATION ENGINE (CHECK 1 & CHECK 2)
# -------------------------------------------------------------

class GSTReconciler:
    def __init__(self, books_rows, portal_rows, tolerance=2.00):
        self.books_rows = books_rows
        self.portal_rows = portal_rows
        self.tolerance = tolerance

        self.vendor_summary = []
        self.books_recon = []
        self.portal_recon = []
        self.actionable_list = []

    def run_reconciliation(self):
        print(f"[*] Starting Reconciliation...")
        print(f"    Books Records: {len(self.books_rows)}")
        print(f"    Portal Records: {len(self.portal_rows)}")

        # -------------------------------------------------------------
        # CHECK 1: VENDOR GSTIN PIVOT MACRO RECONCILIATION
        # -------------------------------------------------------------
        print("[*] Executing Check 1: Vendor GSTIN Pivot Reconciliation...")
        books_vendor_agg = {}
        for r in self.books_rows:
            g = r['gstin']
            if g not in books_vendor_agg:
                books_vendor_agg[g] = {
                    'name': r['vendor_name'],
                    'count': 0,
                    'taxable': 0.0,
                    'total_tax': 0.0,
                    'igst': 0.0,
                    'cgst': 0.0,
                    'sgst': 0.0
                }
            books_vendor_agg[g]['count'] += 1
            books_vendor_agg[g]['taxable'] += r['taxable']
            books_vendor_agg[g]['total_tax'] += r['total_tax']
            books_vendor_agg[g]['igst'] += r['igst']
            books_vendor_agg[g]['cgst'] += r['cgst']
            books_vendor_agg[g]['sgst'] += r['sgst']

        portal_vendor_agg = {}
        for r in self.portal_rows:
            g = r['gstin']
            if g not in portal_vendor_agg:
                portal_vendor_agg[g] = {
                    'name': r['supplier_name'],
                    'count': 0,
                    'taxable': 0.0,
                    'total_tax': 0.0,
                    'igst': 0.0,
                    'cgst': 0.0,
                    'sgst': 0.0
                }
            portal_vendor_agg[g]['count'] += 1
            portal_vendor_agg[g]['taxable'] += r['taxable']
            portal_vendor_agg[g]['total_tax'] += r['total_tax']
            portal_vendor_agg[g]['igst'] += r['igst']
            portal_vendor_agg[g]['cgst'] += r['cgst']
            portal_vendor_agg[g]['sgst'] += r['sgst']

        all_gstins = sorted(list(set(books_vendor_agg.keys()) | set(portal_vendor_agg.keys())))

        for g in all_gstins:
            b_info = books_vendor_agg.get(g, {'name': '', 'count': 0, 'taxable': 0.0, 'total_tax': 0.0, 'igst': 0.0, 'cgst': 0.0, 'sgst': 0.0})
            p_info = portal_vendor_agg.get(g, {'name': '', 'count': 0, 'taxable': 0.0, 'total_tax': 0.0, 'igst': 0.0, 'cgst': 0.0, 'sgst': 0.0})

            vendor_name = b_info['name'] or p_info['name']
            b_tax = round(b_info['total_tax'], 2)
            p_tax = round(p_info['total_tax'], 2)
            tax_diff = round(b_tax - p_tax, 2)
            taxable_diff = round(b_info['taxable'] - p_info['taxable'], 2)

            vendor_bill_count = max(b_info['count'], p_info['count'])
            dynamic_vendor_tol = max(self.tolerance, round(vendor_bill_count * 0.75, 2))

            if g in books_vendor_agg and g in portal_vendor_agg:
                if abs(tax_diff) <= self.tolerance:
                    status = "100% Matched"
                elif abs(tax_diff) <= dynamic_vendor_tol:
                    status = "100% Matched (Round-off)"
                else:
                    status = "Tax Variance"
            elif g in books_vendor_agg:
                status = "Only in Books"
            else:
                status = "Only in Portal"

            self.vendor_summary.append({
                'gstin': g,
                'vendor_name': vendor_name,
                'status': status,
                'books_tax': b_tax,
                'portal_tax': p_tax,
                'tax_variance': tax_diff,
                'books_taxable': round(b_info['taxable'], 2),
                'portal_taxable': round(p_info['taxable'], 2),
                'taxable_variance': taxable_diff,
                'books_inv_count': b_info['count'],
                'portal_inv_count': p_info['count']
            })

        # -------------------------------------------------------------
        # CHECK 2: INVOICE & AMOUNT LEVEL RECONCILIATION
        # -------------------------------------------------------------
        print("[*] Executing Check 2: Invoice & Amount Level Matching...")

        # Multi-line invoice aggregation in Books
        # key = (GSTIN, Clean_Bill_No)
        books_by_key = {}
        for idx, r in enumerate(self.books_rows):
            k = (r['gstin'], r['clean_bill_no'])
            if k not in books_by_key:
                books_by_key[k] = {
                    'gstin': r['gstin'],
                    'clean_bill_no': r['clean_bill_no'],
                    'bill_no': r['bill_no'],
                    'vendor_name': r['vendor_name'],
                    'taxable': 0.0,
                    'total_tax': 0.0,
                    'igst': 0.0,
                    'cgst': 0.0,
                    'sgst': 0.0,
                    'row_indices': []
                }
            books_by_key[k]['taxable'] += r['taxable']
            books_by_key[k]['total_tax'] += r['total_tax']
            books_by_key[k]['igst'] += r['igst']
            books_by_key[k]['cgst'] += r['cgst']
            books_by_key[k]['sgst'] += r['sgst']
            books_by_key[k]['row_indices'].append(idx)

        # Index Portal invoices by (GSTIN, Clean_Doc_No)
        portal_by_key = {}
        portal_by_gstin = {}
        for idx, r in enumerate(self.portal_rows):
            k = (r['gstin'], r['clean_doc_no'])
            portal_by_key[k] = idx
            portal_by_gstin.setdefault(r['gstin'], []).append(idx)

        matched_books_keys = {}  # k -> portal_idx, match_type
        matched_portal_indices = set()

        # PASS 1: Exact Key Match (GSTIN + Clean Invoice No)
        for k, b_agg in books_by_key.items():
            if k in portal_by_key:
                p_idx = portal_by_key[k]
                p_row = self.portal_rows[p_idx]
                tax_diff = abs(b_agg['total_tax'] - p_row['total_tax'])

                if tax_diff <= self.tolerance:
                    matched_books_keys[k] = (p_idx, "Matched (Exact)")
                else:
                    matched_books_keys[k] = (p_idx, "Value Mismatch")
                matched_portal_indices.add(p_idx)

        # PASS 2: Smart / Typo Match (Same GSTIN, Same Tax, Invoice Typo / Substring)
        for k, b_agg in books_by_key.items():
            if k in matched_books_keys:
                continue

            gst = b_agg['gstin']
            c_bill = b_agg['clean_bill_no']
            cand_indices = portal_by_gstin.get(gst, [])

            best_p_idx = None
            for p_idx in cand_indices:
                if p_idx in matched_portal_indices:
                    continue
                p_row = self.portal_rows[p_idx]
                tax_diff = abs(b_agg['total_tax'] - p_row['total_tax'])

                if tax_diff <= self.tolerance:
                    p_doc = p_row['clean_doc_no']
                    sim = string_similarity(c_bill, p_doc)
                    # Check 1: Similarity / Substring
                    is_match = (sim >= 0.80 or c_bill in p_doc or p_doc in c_bill)
                    # Check 2: Date placed in Doc No field by vendor (e.g. 16/07/2025 in doc no)
                    clean_p_raw = re.sub(r'[^0-9]', '', p_row['doc_no'])
                    if not is_match and len(clean_p_raw) >= 6:
                        for row_i in b_agg['row_indices']:
                            b_date_clean = re.sub(r'[^0-9]', '', self.books_rows[row_i]['inv_date'])
                            if clean_p_raw in b_date_clean or b_date_clean in clean_p_raw:
                                is_match = True
                                break
                    # Check 3: Empty bill in SAP with unique tax match for that vendor
                    if not is_match and (not c_bill or c_bill == "0"):
                        same_tax_count = sum(1 for pi in cand_indices if abs(self.portal_rows[pi]['total_tax'] - b_agg['total_tax']) <= self.tolerance)
                        if same_tax_count == 1:
                            is_match = True

                    if is_match:
                        best_p_idx = p_idx
                        break

            if best_p_idx is not None:
                matched_books_keys[k] = (best_p_idx, "Matched (Smart/Typo)")
                matched_portal_indices.add(best_p_idx)

        # -------------------------------------------------------------
        # BUILD VIEW 1: BOOKS VS PORTAL (Mapped for every row in Books)
        # -------------------------------------------------------------
        print("[*] Generating Books vs Portal detailed mapping...")
        for r in self.books_rows:
            k = (r['gstin'], r['clean_bill_no'])
            if k in matched_books_keys:
                p_idx, match_status = matched_books_keys[k]
                p_row = self.portal_rows[p_idx]
                b_agg = books_by_key[k]

                b_recon_row = dict(r)
                b_recon_row['portal_doc_no'] = p_row['doc_no']
                b_recon_row['portal_doc_date'] = p_row['doc_date']
                b_recon_row['portal_taxable'] = p_row['taxable']
                b_recon_row['portal_total_tax'] = p_row['total_tax']
                b_recon_row['portal_igst'] = p_row['igst']
                b_recon_row['portal_cgst'] = p_row['cgst']
                b_recon_row['portal_sgst'] = p_row['sgst']
                b_recon_row['tax_diff'] = round(b_agg['total_tax'] - p_row['total_tax'], 2)
                b_recon_row['match_status'] = match_status
                self.books_recon.append(b_recon_row)
            else:
                b_recon_row = dict(r)
                b_recon_row['portal_doc_no'] = "-"
                b_recon_row['portal_doc_date'] = "-"
                b_recon_row['portal_taxable'] = 0.0
                b_recon_row['portal_total_tax'] = 0.0
                b_recon_row['portal_igst'] = 0.0
                b_recon_row['portal_cgst'] = 0.0
                b_recon_row['portal_sgst'] = 0.0
                b_recon_row['tax_diff'] = r['total_tax']
                b_recon_row['match_status'] = "Only in Books (Missing in 2B)"
                self.books_recon.append(b_recon_row)

        # -------------------------------------------------------------
        # BUILD VIEW 2: PORTAL VS BOOKS (Mapped for every row in Portal)
        # -------------------------------------------------------------
        print("[*] Generating Portal vs Books detailed mapping...")
        # Invert matched_books_keys: p_idx -> list of k
        portal_to_books_key = {}
        for k, (p_idx, m_status) in matched_books_keys.items():
            portal_to_books_key[p_idx] = (k, m_status)

        for p_idx, p_row in enumerate(self.portal_rows):
            p_recon_row = dict(p_row)
            if p_idx in portal_to_books_key:
                k, m_status = portal_to_books_key[p_idx]
                b_agg = books_by_key[k]
                p_recon_row['books_bill_no'] = b_agg['bill_no']
                p_recon_row['books_taxable'] = round(b_agg['taxable'], 2)
                p_recon_row['books_total_tax'] = round(b_agg['total_tax'], 2)
                p_recon_row['books_igst'] = round(b_agg['igst'], 2)
                p_recon_row['books_cgst'] = round(b_agg['cgst'], 2)
                p_recon_row['books_sgst'] = round(b_agg['sgst'], 2)
                p_recon_row['tax_diff'] = round(p_row['total_tax'] - b_agg['total_tax'], 2)
                p_recon_row['match_status'] = m_status
            else:
                p_recon_row['books_bill_no'] = "-"
                p_recon_row['books_taxable'] = 0.0
                p_recon_row['books_total_tax'] = 0.0
                p_recon_row['books_igst'] = 0.0
                p_recon_row['books_cgst'] = 0.0
                p_recon_row['books_sgst'] = 0.0
                p_recon_row['tax_diff'] = p_row['total_tax']
                p_recon_row['match_status'] = "Only in Portal (Unbooked)"

            self.portal_recon.append(p_recon_row)

        # -------------------------------------------------------------
        # BUILD ACTIONABLE MISSING INVOICES LIST
        # -------------------------------------------------------------
        print("[*] Building Actionable Discrepancy List...")
        # 1. Missing in Portal (Books entries needing vendor follow-up)
        for r in self.books_recon:
            if r['match_status'] == "Only in Books (Missing in 2B)":
                self.actionable_list.append({
                    'action_type': 'Vendor Follow-up (Missing in 2B)',
                    'gstin': r['gstin'],
                    'party_name': r['vendor_name'],
                    'invoice_no': r['bill_no'],
                    'invoice_date': r['inv_date'],
                    'taxable_value': r['taxable'],
                    'total_tax': r['total_tax'],
                    'impact': 'ITC at Risk (Vendor GSTR-1 not filed)',
                    'recommended_action': 'Send follow-up email/reminder to Vendor with Bill details'
                })
            elif r['match_status'] == "Value Mismatch":
                self.actionable_list.append({
                    'action_type': 'Value Discrepancy Investigation',
                    'gstin': r['gstin'],
                    'party_name': r['vendor_name'],
                    'invoice_no': r['bill_no'],
                    'invoice_date': r['inv_date'],
                    'taxable_value': r['taxable'],
                    'total_tax': r['total_tax'],
                    'impact': f"Tax Variance: ₹{r['tax_diff']:.2f}",
                    'recommended_action': f"Check tax rate/partial booking (Portal Tax: ₹{r['portal_total_tax']:.2f})"
                })

        # 2. Missing in Books (Portal entries needing accounting booking)
        for r in self.portal_recon:
            if r['match_status'] == "Only in Portal (Unbooked)":
                self.actionable_list.append({
                    'action_type': 'Accounting Booking Pending',
                    'gstin': r['gstin'],
                    'party_name': r['supplier_name'],
                    'invoice_no': r['doc_no'],
                    'invoice_date': r['doc_date'],
                    'taxable_value': r['taxable'],
                    'total_tax': r['total_tax'],
                    'impact': f"Unclaimed ITC Available: ₹{r['total_tax']:.2f}",
                    'recommended_action': 'Check bill physical copy and book in SAP to claim ITC'
                })

        print(f"[OK] Reconciliation complete! Ready to export Excel.")


# -------------------------------------------------------------
# 4. EXCEL EXPORTER WITH PROFESSIONAL STYLING
# -------------------------------------------------------------

def export_reconciliation_workbook(reconciler, output_path):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Styles
    navy_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    dark_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)
    
    # Status fills
    matched_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    matched_font = Font(name="Calibri", size=10, color="006100", bold=True)

    smart_fill = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")
    smart_font = Font(name="Calibri", size=10, color="0C5460", bold=True)

    mismatch_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    mismatch_font = Font(name="Calibri", size=10, color="856404", bold=True)

    missing_books_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    missing_books_font = Font(name="Calibri", size=10, color="721C24", bold=True)

    missing_portal_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    missing_portal_font = Font(name="Calibri", size=10, color="C65911", bold=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    currency_format = "#,##0.00"

    def apply_header_style(ws, headers):
        ws.append(headers)
        ws.row_dimensions[1].height = 26
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = navy_fill
            cell.font = dark_header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def autofit_columns(ws):
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    # ---------------------------------------------------------
    # TAB 1: EXECUTIVE DASHBOARD & SUMMARY
    # ---------------------------------------------------------
    ws_dash = wb.create_sheet(title="Dashboard & KPI Summary")
    ws_dash.views.sheetView[0].showGridLines = True
    
    # Title
    ws_dash.merge_cells("A1:G1")
    title_cell = ws_dash["A1"]
    title_cell.value = "GST RECONCILIATION SUMMARY (PORTAL vs SAP BOOKS)"
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws_dash.row_dimensions[1].height = 35

    # Timestamp
    ws_dash["A2"] = f"Report Generated: {datetime.now().strftime('%d-%b-%Y %H:%M:%S')} | Tolerance: ₹{reconciler.tolerance:.2f}"
    ws_dash["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")

    # Metrics
    total_books_tax = sum(r['total_tax'] for r in reconciler.books_rows)
    total_portal_tax = sum(r['total_tax'] for r in reconciler.portal_rows)
    matched_tax = sum(r['total_tax'] for r in reconciler.books_recon if "Matched" in r['match_status'])
    unmatched_books_tax = sum(r['total_tax'] for r in reconciler.books_recon if r['match_status'] == "Only in Books (Missing in 2B)")
    unmatched_portal_tax = sum(r['total_tax'] for r in reconciler.portal_recon if r['match_status'] == "Only in Portal (Unbooked)")

    metrics_table = [
        ("Key Metric", "Count / Invoices", "Total Amount (₹)"),
        ("Total ITC Recorded in SAP Books", len(reconciler.books_rows), total_books_tax),
        ("Total ITC Available in Portal (2B)", len(reconciler.portal_rows), total_portal_tax),
        ("Net Variance (Books - Portal)", "-", round(total_books_tax - total_portal_tax, 2)),
        ("Reconciled ITC (Matched Invoices)", sum(1 for r in reconciler.books_recon if "Matched" in r['match_status']), matched_tax),
        ("ITC at Risk (In Books, Missing in 2B)", sum(1 for r in reconciler.books_recon if r['match_status'] == "Only in Books (Missing in 2B)"), unmatched_books_tax),
        ("Unclaimed ITC (In Portal, Unbooked in SAP)", sum(1 for r in reconciler.portal_recon if r['match_status'] == "Only in Portal (Unbooked)"), unmatched_portal_tax),
    ]

    ws_dash.append([])
    for row_idx, row_data in enumerate(metrics_table, start=4):
        ws_dash.append(row_data)
        ws_dash.row_dimensions[row_idx].height = 22
        for c_idx in range(1, 4):
            cell = ws_dash.cell(row=row_idx, column=c_idx)
            cell.border = thin_border
            if row_idx == 4:
                cell.fill = navy_fill
                cell.font = dark_header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = bold_font if c_idx == 1 else regular_font
                if c_idx == 3:
                    cell.number_format = currency_format
                    cell.alignment = Alignment(horizontal="right")
                else:
                    cell.alignment = Alignment(horizontal="center" if c_idx == 2 else "left")

    autofit_columns(ws_dash)

    # ---------------------------------------------------------
    # TAB 2: CHECK 1 - VENDOR GSTIN PIVOT SUMMARY
    # ---------------------------------------------------------
    ws_vendor = wb.create_sheet(title="1_Vendor_Pivot_Check")
    ws_vendor.views.sheetView[0].showGridLines = True
    v_headers = [
        "Supplier GSTIN", "Vendor / Supplier Name", "Status",
        "Books Total Tax", "Portal Total Tax", "Tax Variance (Books - Portal)",
        "Books Taxable", "Portal Taxable", "Taxable Variance",
        "Books Inv Count", "Portal Inv Count"
    ]
    apply_header_style(ws_vendor, v_headers)

    for r in reconciler.vendor_summary:
        ws_vendor.append([
            r['gstin'], r['vendor_name'], r['status'],
            r['books_tax'], r['portal_tax'], r['tax_variance'],
            r['books_taxable'], r['portal_taxable'], r['taxable_variance'],
            r['books_inv_count'], r['portal_inv_count']
        ])
        curr_row = ws_vendor.max_row
        ws_vendor.row_dimensions[curr_row].height = 20

        # Status styling
        status_cell = ws_vendor.cell(row=curr_row, column=3)
        if r['status'] == "100% Matched":
            status_cell.fill = matched_fill
            status_cell.font = matched_font
        elif r['status'] == "Tax Variance":
            status_cell.fill = mismatch_fill
            status_cell.font = mismatch_font
        elif r['status'] == "Only in Books":
            status_cell.fill = missing_books_fill
            status_cell.font = missing_books_font
        else:
            status_cell.fill = missing_portal_fill
            status_cell.font = missing_portal_font

        # Borders and number formats
        for c_idx in range(1, len(v_headers) + 1):
            cell = ws_vendor.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if c_idx in [4, 5, 6, 7, 8, 9]:
                cell.number_format = currency_format
                cell.alignment = Alignment(horizontal="right")
            elif c_idx in [10, 11]:
                cell.alignment = Alignment(horizontal="center")

    autofit_columns(ws_vendor)

    # ---------------------------------------------------------
    # TAB 3: CHECK 2 - BOOKS VS PORTAL (DETAILED)
    # ---------------------------------------------------------
    ws_bvp = wb.create_sheet(title="2_Books_vs_Portal")
    ws_bvp.views.sheetView[0].showGridLines = True
    b_headers = [
        "Branch", "SAP Trans No", "Posting Date", "Invoice Date",
        "Vendor Bill No", "Vendor Code", "Vendor Name", "Vendor GSTIN",
        "Books Taxable", "Books IGST", "Books CGST", "Books SGST", "Books Total Tax",
        "Portal Doc No", "Portal Doc Date", "Portal Taxable", "Portal Total Tax",
        "Tax Difference", "Reconciliation Status"
    ]
    apply_header_style(ws_bvp, b_headers)

    for r in reconciler.books_recon:
        ws_bvp.append([
            r['branch'], r['sap_trans_no'], r['posting_date'], r['inv_date'],
            r['bill_no'], r['vendor_code'], r['vendor_name'], r['gstin'],
            r['taxable'], r['igst'], r['cgst'], r['sgst'], r['total_tax'],
            r['portal_doc_no'], r['portal_doc_date'], r['portal_taxable'], r['portal_total_tax'],
            r['tax_diff'], r['match_status']
        ])
        curr_row = ws_bvp.max_row
        ws_bvp.row_dimensions[curr_row].height = 19

        status_cell = ws_bvp.cell(row=curr_row, column=len(b_headers))
        st = r['match_status']
        if st == "Matched (Exact)":
            status_cell.fill = matched_fill
            status_cell.font = matched_font
        elif st == "Matched (Smart/Typo)":
            status_cell.fill = smart_fill
            status_cell.font = smart_font
        elif st == "Value Mismatch":
            status_cell.fill = mismatch_fill
            status_cell.font = mismatch_font
        else:
            status_cell.fill = missing_books_fill
            status_cell.font = missing_books_font

        for c_idx in range(1, len(b_headers) + 1):
            cell = ws_bvp.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if c_idx in [9, 10, 11, 12, 13, 16, 17, 18]:
                cell.number_format = currency_format
                cell.alignment = Alignment(horizontal="right")

    autofit_columns(ws_bvp)

    # ---------------------------------------------------------
    # TAB 4: CHECK 2 - PORTAL VS BOOKS (DETAILED)
    # ---------------------------------------------------------
    ws_pvb = wb.create_sheet(title="3_Portal_vs_Books")
    ws_pvb.views.sheetView[0].showGridLines = True
    p_headers = [
        "Supplier GSTIN", "Supplier Name", "Portal Doc No", "Portal Doc Date",
        "Portal Taxable", "Portal IGST", "Portal CGST", "Portal SGST", "Portal Total Tax",
        "Reverse Charge", "Matched Books Bill No", "Books Taxable", "Books Total Tax",
        "Tax Difference", "Booking / Match Status"
    ]
    apply_header_style(ws_pvb, p_headers)

    for r in reconciler.portal_recon:
        ws_pvb.append([
            r['gstin'], r['supplier_name'], r['doc_no'], r['doc_date'],
            r['taxable'], r['igst'], r['cgst'], r['sgst'], r['total_tax'],
            r['rc'], r['books_bill_no'], r['books_taxable'], r['books_total_tax'],
            r['tax_diff'], r['match_status']
        ])
        curr_row = ws_pvb.max_row
        ws_pvb.row_dimensions[curr_row].height = 19

        status_cell = ws_pvb.cell(row=curr_row, column=len(p_headers))
        st = r['match_status']
        if st == "Matched (Exact)":
            status_cell.fill = matched_fill
            status_cell.font = matched_font
        elif st == "Matched (Smart/Typo)":
            status_cell.fill = smart_fill
            status_cell.font = smart_font
        elif st == "Value Mismatch":
            status_cell.fill = mismatch_fill
            status_cell.font = mismatch_font
        else:
            status_cell.fill = missing_portal_fill
            status_cell.font = missing_portal_font

        for c_idx in range(1, len(p_headers) + 1):
            cell = ws_pvb.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if c_idx in [5, 6, 7, 8, 9, 12, 13, 14]:
                cell.number_format = currency_format
                cell.alignment = Alignment(horizontal="right")

    autofit_columns(ws_pvb)

    # ---------------------------------------------------------
    # TAB 5: ACTIONABLE DISCREPANCY & FOLLOW-UP LIST
    # ---------------------------------------------------------
    ws_act = wb.create_sheet(title="4_Actionable_Discrepancies")
    ws_act.views.sheetView[0].showGridLines = True
    act_headers = [
        "Category", "Supplier GSTIN", "Party Name", "Invoice / Doc No",
        "Invoice Date", "Taxable Value", "Total Tax", "Audit Impact / Variance", "Recommended Action"
    ]
    apply_header_style(ws_act, act_headers)

    for r in reconciler.actionable_list:
        ws_act.append([
            r['action_type'], r['gstin'], r['party_name'], r['invoice_no'],
            r['invoice_date'], r['taxable_value'], r['total_tax'], r['impact'], r['recommended_action']
        ])
        curr_row = ws_act.max_row
        ws_act.row_dimensions[curr_row].height = 20

        cat_cell = ws_act.cell(row=curr_row, column=1)
        if "Missing in 2B" in r['action_type']:
            cat_cell.fill = missing_books_fill
            cat_cell.font = missing_books_font
        elif "Unbooked" in r['action_type'] or "Booking Pending" in r['action_type']:
            cat_cell.fill = missing_portal_fill
            cat_cell.font = missing_portal_font
        else:
            cat_cell.fill = mismatch_fill
            cat_cell.font = mismatch_font

        for c_idx in range(1, len(act_headers) + 1):
            cell = ws_act.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if c_idx in [6, 7]:
                cell.number_format = currency_format
                cell.alignment = Alignment(horizontal="right")

    autofit_columns(ws_act)

    wb.save(output_path)
    print(f"[OK] Workbook successfully saved to: {output_path}")
