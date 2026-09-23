"""
GST Reconciliation Engine (GSTR-2B vs SAP Books)
Developed for V. Singhi & Associates / vsinghidelhi
Author: AI Assistant
Description:
  Automates 2-Way Reconciliation between GST Portal (GSTR-2B) and SAP Purchase Register.
  Implements a 2-Check Architecture:
    - Check 1: Vendor GSTIN Pivot Macro Reconciliation (with IGST, CGST, SGST tracking)
    - Check 2: Invoice & Amount Level Micro Reconciliation
      * Individual Tax Head Matching: IGST, CGST, SGST matched independently
      * Place of Supply / Tax Head Mismatch detection (IGST vs CGST/SGST)
      * Explainable Smart/Typo matching with side-by-side comparison columns
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

def get_levenshtein_distance(s1, s2):
    if len(s1) < len(s2):
        return get_levenshtein_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)
    prev = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        curr = [i + 1]
        for j, c2 in enumerate(s2):
            ins = prev[j + 1] + 1
            dels = curr[j] + 1
            subs = prev[j] + (c1 != c2)
            curr.append(min(ins, dels, subs))
        prev = curr
    return prev[-1]

def diagnose_smart_match(b_bill, p_doc, b_clean, p_clean, b_date, p_date):
    """
    Evaluates whether an unlinked invoice pair represents a legitimate typo / format variance.
    Returns (is_match: bool, reason: str).
    Guaranteed never to match grouped bills (with commas) or dissimilar invoice patterns.
    """
    if not b_clean or not p_clean:
        return False, ""
    
    # Do not match grouped bills with commas/semicolons into single portal docs
    if ',' in b_bill or ';' in b_bill:
        return False, ""

    # Exact after normalization (FY standardizing 2025-26 -> 25-26 or separator removal)
    if b_clean == p_clean:
        return True, "Normalized: FY format (2025-26 -> 25-26) / leading zeros / separators"

    # 1. 1-character typo (e.g. 0048 vs 004B -> 48 vs 4B)
    if len(b_clean) == len(p_clean) and len(b_clean) >= 2:
        diffs = [(b, p) for b, p in zip(b_clean, p_clean) if b != p]
        if len(diffs) == 1:
            return True, f"1-Char Typo: '{b_bill}' vs '{p_doc}' (Char '{diffs[0][0]}' vs '{diffs[0][1]}')"

    # 2. Levenshtein edit distance = 1 (1 character inserted/deleted)
    if abs(len(b_clean) - len(p_clean)) <= 1 and len(b_clean) >= 3 and len(p_clean) >= 3:
        if get_levenshtein_distance(b_clean, p_clean) == 1:
            return True, f"1-Char Edit: '{b_bill}' vs '{p_doc}'"

    # 3. Prefix omission (e.g. Books entered '079' while Portal has 'PSB-079')
    if len(b_clean) >= 2 and len(p_clean) >= 4:
        b_num = b_clean.lstrip('0')
        p_num = re.sub(r'^[A-Z]+', '', p_clean).lstrip('0')
        if b_num and p_num and b_num == p_num:
            prefix_match = re.match(r'^[A-Z]+', p_clean)
            p_str = prefix_match.group(0) if prefix_match else ""
            return True, f"Prefix Omitted in Books: '{p_str}' ('{b_bill}' vs '{p_doc}')"
        if p_clean.endswith(b_clean):
            omitted = p_clean[:-len(b_clean)]
            return True, f"Prefix Omitted in Books: '{omitted}' ('{b_bill}' vs '{p_doc}')"

    # 4. Known Abbreviation (e.g. TechBOT in Books vs TB in Portal)
    if len(b_clean) >= 6 and len(p_clean) >= 6 and b_clean[:4] == p_clean[:4] and b_clean[-4:] == p_clean[-4:]:
        return True, f"Abbreviation Match: '{b_bill}' vs '{p_doc}'"

    # 5. Accidental double typing / substring (e.g. ECO/25-26/0ECO/25-26/0258258 vs ECO/25-26/0258)
    if len(b_clean) >= 8 and len(p_clean) >= 6:
        if p_clean in b_clean:
            return True, f"Double-paste in Books corrected: '{b_bill}' vs '{p_doc}'"

    # 6. Vendor entered Invoice Date into Doc No field (e.g. 16/07/2025 in doc no)
    clean_p_nums = re.sub(r'[^0-9]', '', p_doc)
    if len(clean_p_nums) >= 6 and b_date:
        clean_b_date = re.sub(r'[^0-9]', '', b_date)
        if clean_p_nums in clean_b_date or clean_b_date in clean_p_nums:
            return True, f"Date entered in Doc No field: '{p_doc}'"

    return False, ""


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
        # (Tracking Taxable, Total Tax, IGST, CGST, SGST)
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
            
            b_igst = round(b_info['igst'], 2)
            p_igst = round(p_info['igst'], 2)
            igst_diff = round(b_igst - p_igst, 2)

            b_cgst = round(b_info['cgst'], 2)
            p_cgst = round(p_info['cgst'], 2)
            cgst_diff = round(b_cgst - p_cgst, 2)

            b_sgst = round(b_info['sgst'], 2)
            p_sgst = round(p_info['sgst'], 2)
            sgst_diff = round(b_sgst - p_sgst, 2)

            taxable_diff = round(b_info['taxable'] - p_info['taxable'], 2)

            dynamic_vendor_tol = self.tolerance

            if g in books_vendor_agg and g in portal_vendor_agg:
                heads_match = (abs(igst_diff) <= dynamic_vendor_tol and abs(cgst_diff) <= dynamic_vendor_tol and abs(sgst_diff) <= dynamic_vendor_tol)
                if abs(tax_diff) <= self.tolerance and heads_match:
                    status = "100% Matched"
                elif abs(tax_diff) <= dynamic_vendor_tol and not heads_match:
                    status = "Tax Head Mismatch (IGST vs CGST/SGST)"
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
                'books_igst': b_igst,
                'portal_igst': p_igst,
                'igst_diff': igst_diff,
                'books_cgst': b_cgst,
                'portal_cgst': p_cgst,
                'cgst_diff': cgst_diff,
                'books_sgst': b_sgst,
                'portal_sgst': p_sgst,
                'sgst_diff': sgst_diff,
                'books_taxable': round(b_info['taxable'], 2),
                'portal_taxable': round(p_info['taxable'], 2),
                'taxable_variance': taxable_diff,
                'books_inv_count': b_info['count'],
                'portal_inv_count': p_info['count']
            })

        # -------------------------------------------------------------
        # CHECK 2: INVOICE & AMOUNT LEVEL RECONCILIATION
        # (Independent 3-Way Tax Head Verification: IGST, CGST, SGST)
        # -------------------------------------------------------------
        print("[*] Executing Check 2: Invoice & Amount Level Matching (3 Tax Heads)...")

        # Aggregate multi-line invoices in Books at Invoice level
        books_by_key = {}
        for idx, r in enumerate(self.books_rows):
            c_bill = r['clean_bill_no']
            if not c_bill or c_bill == "0":
                doc_ref = r['sap_trans_no'] if r['sap_trans_no'] else f"ROW_{idx}"
                k = (r['gstin'], f"DOC_{doc_ref}")
            else:
                k = (r['gstin'], c_bill)

            if k not in books_by_key:
                books_by_key[k] = {
                    'row_id': r.get('row_id', idx),
                    'branch': r.get('branch', ''),
                    'doc_type': r.get('doc_type', ''),
                    'trans_no': r.get('trans_no', ''),
                    'sap_trans_no': r.get('sap_trans_no', ''),
                    'posting_date': r.get('posting_date', ''),
                    'inv_date': r.get('inv_date', ''),
                    'bill_no': r.get('bill_no', ''),
                    'clean_bill_no': c_bill,
                    'vendor_code': r.get('vendor_code', ''),
                    'vendor_name': r.get('vendor_name', ''),
                    'gstin': r.get('gstin', ''),
                    'rcm': r.get('rcm', ''),
                    'line_count': 0,
                    'taxable': 0.0,
                    'total_tax': 0.0,
                    'igst': 0.0,
                    'cgst': 0.0,
                    'sgst': 0.0,
                    'total_val': 0.0,
                    'row_indices': []
                }
            books_by_key[k]['line_count'] += 1
            books_by_key[k]['taxable'] = round(books_by_key[k]['taxable'] + r['taxable'], 2)
            books_by_key[k]['total_tax'] = round(books_by_key[k]['total_tax'] + r['total_tax'], 2)
            books_by_key[k]['igst'] = round(books_by_key[k]['igst'] + r['igst'], 2)
            books_by_key[k]['cgst'] = round(books_by_key[k]['cgst'] + r['cgst'], 2)
            books_by_key[k]['sgst'] = round(books_by_key[k]['sgst'] + r['sgst'], 2)
            books_by_key[k]['total_val'] = round(books_by_key[k]['total_val'] + r.get('total_val', 0.0), 2)
            books_by_key[k]['row_indices'].append(idx)

        # Index Portal invoices by (GSTIN, Clean_Doc_No)
        portal_by_key = {}
        portal_by_gstin = {}
        for idx, r in enumerate(self.portal_rows):
            k = (r['gstin'], r['clean_doc_no'])
            portal_by_key[k] = idx
            portal_by_gstin.setdefault(r['gstin'], []).append(idx)

        matched_books_keys = {}  # k -> (portal_idx, match_status, match_reason)
        matched_portal_indices = set()

        # PASS 1: Exact Key Match (GSTIN + Clean Invoice No)
        for k, b_agg in books_by_key.items():
            if k in portal_by_key:
                p_idx = portal_by_key[k]
                p_row = self.portal_rows[p_idx]
                diff_tot = round(abs(b_agg['total_tax'] - p_row['total_tax']), 2)
                diff_igst = round(abs(b_agg['igst'] - p_row['igst']), 2)
                diff_cgst = round(abs(b_agg['cgst'] - p_row['cgst']), 2)
                diff_sgst = round(abs(b_agg['sgst'] - p_row['sgst']), 2)

                if diff_tot <= self.tolerance:
                    if diff_igst <= self.tolerance and diff_cgst <= self.tolerance and diff_sgst <= self.tolerance:
                        status = "Matched (Exact)"
                        reason = "Exact Match: Invoice No & All Tax Heads (IGST/CGST/SGST)"
                    else:
                        status = "Tax Head Mismatch (IGST vs CGST/SGST)"
                        reason = f"Total Tax matches, but Head Mismatch: Books(I:{b_agg['igst']:.2f}, C:{b_agg['cgst']:.2f}, S:{b_agg['sgst']:.2f}) vs Portal(I:{p_row['igst']:.2f}, C:{p_row['cgst']:.2f}, S:{p_row['sgst']:.2f})"
                else:
                    status = "Value Mismatch"
                    reason = f"Invoice matches, but Tax difference is ₹{diff_tot:.2f} (Books: ₹{b_agg['total_tax']:.2f}, Portal: ₹{p_row['total_tax']:.2f})"

                matched_books_keys[k] = (p_idx, status, reason)
                matched_portal_indices.add(p_idx)

        # PASS 2: Explainable Smart / Typo Match
        for k, b_agg in books_by_key.items():
            if k in matched_books_keys:
                continue

            gst = b_agg['gstin']
            c_bill = b_agg['clean_bill_no']
            cand_indices = portal_by_gstin.get(gst, [])

            best_p_idx = None
            best_status = None
            best_reason = None

            for p_idx in cand_indices:
                if p_idx in matched_portal_indices:
                    continue
                p_row = self.portal_rows[p_idx]
                diff_tot = round(abs(b_agg['total_tax'] - p_row['total_tax']), 2)
                diff_igst = round(abs(b_agg['igst'] - p_row['igst']), 2)
                diff_cgst = round(abs(b_agg['cgst'] - p_row['cgst']), 2)
                diff_sgst = round(abs(b_agg['sgst'] - p_row['sgst']), 2)

                if diff_tot <= self.tolerance:
                    is_match, reason = diagnose_smart_match(
                        b_agg['bill_no'], p_row['doc_no'],
                        c_bill, p_row['clean_doc_no'],
                        b_agg['inv_date'], p_row['doc_date']
                    )

                    # Also handle blank bill no in Books if exactly one unique bill exists for that vendor with same tax
                    is_blank_bill = (not c_bill or c_bill == "0" or str(c_bill).startswith("DOC_"))
                    if not is_match and is_blank_bill:
                        same_tax_count = sum(1 for pi in cand_indices if abs(self.portal_rows[pi]['total_tax'] - b_agg['total_tax']) <= self.tolerance)
                        if same_tax_count == 1:
                            is_match = True
                            reason = f"Blank Bill No in Books: Unique bill for vendor {b_agg['vendor_name']}"

                    if is_match:
                        if diff_igst <= self.tolerance and diff_cgst <= self.tolerance and diff_sgst <= self.tolerance:
                            status = "Matched (Smart/Typo)"
                        else:
                            status = "Tax Head Mismatch (IGST vs CGST/SGST)"
                            reason += f" | Head Mismatch: Books(I:{b_agg['igst']:.2f}, C:{b_agg['cgst']:.2f}, S:{b_agg['sgst']:.2f}) vs Portal(I:{p_row['igst']:.2f}, C:{p_row['cgst']:.2f}, S:{p_row['sgst']:.2f})"

                        best_p_idx = p_idx
                        best_status = status
                        best_reason = reason
                        break

            if best_p_idx is not None:
                matched_books_keys[k] = (best_p_idx, best_status, best_reason)
                matched_portal_indices.add(best_p_idx)

        # -------------------------------------------------------------
        # BUILD VIEW 1: BOOKS VS PORTAL (Strictly Side-by-Side, 1 Row per Invoice)
        # -------------------------------------------------------------
        print("[*] Generating Books vs Portal detailed side-by-side mapping...")
        for k, b_agg in books_by_key.items():
            b_recon_row = dict(b_agg)
            if b_agg['line_count'] > 1:
                b_recon_row['sap_trans_no'] = f"{b_agg['sap_trans_no']} ({b_agg['line_count']} items)"

            if k in matched_books_keys:
                p_idx, match_status, match_reason = matched_books_keys[k]
                p_row = self.portal_rows[p_idx]

                b_recon_row['portal_doc_no'] = p_row['doc_no']
                b_recon_row['portal_doc_date'] = p_row['doc_date']
                b_recon_row['portal_taxable'] = p_row['taxable']
                b_recon_row['taxable_diff'] = round(b_agg['taxable'] - p_row['taxable'], 2)
                b_recon_row['portal_igst'] = p_row['igst']
                b_recon_row['igst_diff'] = round(b_agg['igst'] - p_row['igst'], 2)
                b_recon_row['portal_cgst'] = p_row['cgst']
                b_recon_row['cgst_diff'] = round(b_agg['cgst'] - p_row['cgst'], 2)
                b_recon_row['portal_sgst'] = p_row['sgst']
                b_recon_row['sgst_diff'] = round(b_agg['sgst'] - p_row['sgst'], 2)
                b_recon_row['portal_total_tax'] = p_row['total_tax']
                b_recon_row['tax_diff'] = round(b_agg['total_tax'] - p_row['total_tax'], 2)
                b_recon_row['match_status'] = match_status
                if b_agg['line_count'] > 1:
                    b_recon_row['match_reason'] = f"{match_reason} (Consolidated {b_agg['line_count']} line items in Books)"
                else:
                    b_recon_row['match_reason'] = match_reason
                self.books_recon.append(b_recon_row)
            else:
                b_recon_row['portal_doc_no'] = "-"
                b_recon_row['portal_doc_date'] = "-"
                b_recon_row['portal_taxable'] = 0.0
                b_recon_row['taxable_diff'] = b_agg['taxable']
                b_recon_row['portal_igst'] = 0.0
                b_recon_row['igst_diff'] = b_agg['igst']
                b_recon_row['portal_cgst'] = 0.0
                b_recon_row['cgst_diff'] = b_agg['cgst']
                b_recon_row['portal_sgst'] = 0.0
                b_recon_row['sgst_diff'] = b_agg['sgst']
                b_recon_row['portal_total_tax'] = 0.0
                b_recon_row['tax_diff'] = b_agg['total_tax']
                b_recon_row['match_status'] = "Only in Books (Missing in 2B)"
                b_recon_row['match_reason'] = "Vendor has not filed invoice in GSTR-1 or GSTIN mismatch" + (f" ({b_agg['line_count']} items)" if b_agg['line_count'] > 1 else "")
                self.books_recon.append(b_recon_row)

        # -------------------------------------------------------------
        # BUILD VIEW 2: PORTAL VS BOOKS (Side-by-Side Presentation)
        # -------------------------------------------------------------
        print("[*] Generating Portal vs Books detailed side-by-side mapping...")
        portal_to_books_key = {}
        for k, (p_idx, m_status, m_reason) in matched_books_keys.items():
            portal_to_books_key[p_idx] = (k, m_status, m_reason)

        for p_idx, p_row in enumerate(self.portal_rows):
            p_recon_row = dict(p_row)
            if p_idx in portal_to_books_key:
                k, m_status, m_reason = portal_to_books_key[p_idx]
                b_agg = books_by_key[k]
                p_recon_row['books_bill_no'] = b_agg['bill_no']
                p_recon_row['books_inv_date'] = b_agg['inv_date']
                p_recon_row['books_taxable'] = round(b_agg['taxable'], 2)
                p_recon_row['taxable_diff'] = round(p_row['taxable'] - b_agg['taxable'], 2)
                p_recon_row['books_igst'] = round(b_agg['igst'], 2)
                p_recon_row['igst_diff'] = round(p_row['igst'] - b_agg['igst'], 2)
                p_recon_row['books_cgst'] = round(b_agg['cgst'], 2)
                p_recon_row['cgst_diff'] = round(p_row['cgst'] - b_agg['cgst'], 2)
                p_recon_row['books_sgst'] = round(b_agg['sgst'], 2)
                p_recon_row['sgst_diff'] = round(p_row['sgst'] - b_agg['sgst'], 2)
                p_recon_row['books_total_tax'] = round(b_agg['total_tax'], 2)
                p_recon_row['tax_diff'] = round(p_row['total_tax'] - b_agg['total_tax'], 2)
                p_recon_row['match_status'] = m_status
                p_recon_row['match_reason'] = m_reason
            else:
                p_recon_row['books_bill_no'] = "-"
                p_recon_row['books_inv_date'] = "-"
                p_recon_row['books_taxable'] = 0.0
                p_recon_row['taxable_diff'] = p_row['taxable']
                p_recon_row['books_igst'] = 0.0
                p_recon_row['igst_diff'] = p_row['igst']
                p_recon_row['books_cgst'] = 0.0
                p_recon_row['cgst_diff'] = p_row['cgst']
                p_recon_row['books_sgst'] = 0.0
                p_recon_row['sgst_diff'] = p_row['sgst']
                p_recon_row['books_total_tax'] = 0.0
                p_recon_row['tax_diff'] = p_row['total_tax']
                p_recon_row['match_status'] = "Only in Portal (Unbooked)"
                p_recon_row['match_reason'] = "Bill available in 2B but not recorded in SAP Books"

            self.portal_recon.append(p_recon_row)

        # -------------------------------------------------------------
        # BUILD ACTIONABLE DISCREPANCY & FOLLOW-UP LIST
        # -------------------------------------------------------------
        print("[*] Building Actionable Discrepancy List...")
        for r in self.books_recon:
            if r['match_status'] == "Only in Books (Missing in 2B)":
                self.actionable_list.append({
                    'action_type': 'Vendor Follow-up (Missing in 2B)',
                    'gstin': r['gstin'],
                    'party_name': r['vendor_name'],
                    'books_bill_no': r['bill_no'],
                    'portal_doc_no': '-',
                    'invoice_date': r['inv_date'],
                    'taxable_value': r['taxable'],
                    'total_tax': r['total_tax'],
                    'heads_breakup': f"Books(I:{r['igst']}, C:{r['cgst']}, S:{r['sgst']}) | Portal(-)",
                    'impact': 'ITC at Risk (Vendor GSTR-1 not filed)',
                    'recommended_action': 'Send follow-up email/reminder to Vendor with Bill details'
                })
            elif r['match_status'] == "Tax Head Mismatch (IGST vs CGST/SGST)":
                self.actionable_list.append({
                    'action_type': 'Tax Head POS Correction',
                    'gstin': r['gstin'],
                    'party_name': r['vendor_name'],
                    'books_bill_no': r['bill_no'],
                    'portal_doc_no': r['portal_doc_no'],
                    'invoice_date': r['inv_date'],
                    'taxable_value': r['taxable'],
                    'total_tax': r['total_tax'],
                    'heads_breakup': f"Books(I:{r['igst']}, C:{r['cgst']}, S:{r['sgst']}) vs Portal(I:{r['portal_igst']}, C:{r['portal_cgst']}, S:{r['portal_sgst']})",
                    'impact': 'Place of Supply (POS) Error / Incorrect Tax Head Booked',
                    'recommended_action': 'Reclassify entry in SAP: Intra-state vs Inter-state tax heads'
                })
            elif r['match_status'] == "Value Mismatch":
                self.actionable_list.append({
                    'action_type': 'Value Discrepancy Investigation',
                    'gstin': r['gstin'],
                    'party_name': r['vendor_name'],
                    'books_bill_no': r['bill_no'],
                    'portal_doc_no': r['portal_doc_no'],
                    'invoice_date': r['inv_date'],
                    'taxable_value': r['taxable'],
                    'total_tax': r['total_tax'],
                    'heads_breakup': f"Books Tax: ₹{r['total_tax']:.2f} vs Portal Tax: ₹{r['portal_total_tax']:.2f}",
                    'impact': f"Tax Variance: ₹{r['tax_diff']:.2f}",
                    'recommended_action': f"Check tax rate or partial booking (Portal Tax: ₹{r['portal_total_tax']:.2f})"
                })

        for r in self.portal_recon:
            if r['match_status'] == "Only in Portal (Unbooked)":
                self.actionable_list.append({
                    'action_type': 'Accounting Booking Pending',
                    'gstin': r['gstin'],
                    'party_name': r['supplier_name'],
                    'books_bill_no': '-',
                    'portal_doc_no': r['doc_no'],
                    'invoice_date': r['doc_date'],
                    'taxable_value': r['taxable'],
                    'total_tax': r['total_tax'],
                    'heads_breakup': f"Portal(I:{r['igst']}, C:{r['cgst']}, S:{r['sgst']}) | Books(-)",
                    'impact': f"Unclaimed ITC Available: ₹{r['total_tax']:.2f}",
                    'recommended_action': 'Check physical bill copy and book in SAP to claim ITC'
                })

        print(f"[OK] Reconciliation complete! Ready to export Excel.")


# -------------------------------------------------------------
# 4. EXCEL EXPORTER WITH PROFESSIONAL SIDE-BY-SIDE STYLING
# -------------------------------------------------------------

def export_reconciliation_workbook(reconciler, output_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styling Palettes
    navy_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    dark_header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=10, bold=True)
    regular_font = Font(name="Calibri", size=10)

    # Status fills
    matched_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    matched_font = Font(name="Calibri", size=10, color="006100", bold=True)

    smart_fill = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")
    smart_font = Font(name="Calibri", size=10, color="0C5460", bold=True)

    pos_mismatch_fill = PatternFill(start_color="E2D9F3", end_color="E2D9F3", fill_type="solid")
    pos_mismatch_font = Font(name="Calibri", size=10, color="512DA8", bold=True)

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
        ws.row_dimensions[1].height = 28
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
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 45)

    # ---------------------------------------------------------
    # TAB 1: EXECUTIVE DASHBOARD & SUMMARY
    # ---------------------------------------------------------
    ws_dash = wb.create_sheet(title="Dashboard & KPI Summary")
    ws_dash.views.sheetView[0].showGridLines = True

    ws_dash.merge_cells("A1:G1")
    title_cell = ws_dash["A1"]
    title_cell.value = "GST 2-WAY RECONCILIATION SUMMARY (PORTAL vs SAP BOOKS)"
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws_dash.row_dimensions[1].height = 35

    ws_dash["A2"] = f"Report Generated: {datetime.now().strftime('%d-%b-%Y %H:%M:%S')} | Tolerance: ₹{reconciler.tolerance:.2f} | 3 Tax Heads Independently Verified"
    ws_dash["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")

    total_books_tax = sum(r['total_tax'] for r in reconciler.books_rows)
    total_portal_tax = sum(r['total_tax'] for r in reconciler.portal_rows)
    matched_tax = sum(r['total_tax'] for r in reconciler.books_recon if "Matched" in r['match_status'])
    head_mismatch_tax = sum(r['total_tax'] for r in reconciler.books_recon if "Head Mismatch" in r['match_status'])
    unmatched_books_tax = sum(r['total_tax'] for r in reconciler.books_recon if r['match_status'] == "Only in Books (Missing in 2B)")
    unmatched_portal_tax = sum(r['total_tax'] for r in reconciler.portal_recon if r['match_status'] == "Only in Portal (Unbooked)")

    metrics_table = [
        ("Key Metric", "Count / Invoices", "Total Amount (₹)"),
        ("Total ITC Recorded in SAP Books", len(reconciler.books_rows), total_books_tax),
        ("Total ITC Available in Portal (2B)", len(reconciler.portal_rows), total_portal_tax),
        ("Net Variance (Books - Portal)", "-", round(total_books_tax - total_portal_tax, 2)),
        ("Reconciled ITC (Exact & Typo Matched)", sum(1 for r in reconciler.books_recon if "Matched" in r['match_status']), matched_tax),
        ("Tax Head Mismatch (IGST vs CGST/SGST POS Issue)", sum(1 for r in reconciler.books_recon if "Head Mismatch" in r['match_status']), head_mismatch_tax),
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
    # TAB 2: CHECK 1 - VENDOR GSTIN PIVOT SUMMARY (WITH HEADS)
    # ---------------------------------------------------------
    ws_vendor = wb.create_sheet(title="1_Vendor_Pivot_Check")
    ws_vendor.views.sheetView[0].showGridLines = True
    v_headers = [
        "Supplier GSTIN", "Vendor / Supplier Name", "Status",
        "Books Total Tax", "Portal Total Tax", "Tax Variance (Books - Portal)",
        "Books IGST", "Portal IGST", "IGST Variance",
        "Books CGST", "Portal CGST", "CGST Variance",
        "Books SGST", "Portal SGST", "SGST Variance",
        "Books Taxable", "Portal Taxable", "Taxable Variance",
        "Books Inv Count", "Portal Inv Count"
    ]
    apply_header_style(ws_vendor, v_headers)

    for r in reconciler.vendor_summary:
        ws_vendor.append([
            r['gstin'], r['vendor_name'], r['status'],
            r['books_tax'], r['portal_tax'], r['tax_variance'],
            r['books_igst'], r['portal_igst'], r['igst_diff'],
            r['books_cgst'], r['portal_cgst'], r['cgst_diff'],
            r['books_sgst'], r['portal_sgst'], r['sgst_diff'],
            r['books_taxable'], r['portal_taxable'], r['taxable_variance'],
            r['books_inv_count'], r['portal_inv_count']
        ])
        curr_row = ws_vendor.max_row
        ws_vendor.row_dimensions[curr_row].height = 20

        status_cell = ws_vendor.cell(row=curr_row, column=3)
        if "100% Matched" in r['status']:
            status_cell.fill = matched_fill
            status_cell.font = matched_font
        elif "Head Mismatch" in r['status']:
            status_cell.fill = pos_mismatch_fill
            status_cell.font = pos_mismatch_font
        elif r['status'] == "Tax Variance":
            status_cell.fill = mismatch_fill
            status_cell.font = mismatch_font
        elif r['status'] == "Only in Books":
            status_cell.fill = missing_books_fill
            status_cell.font = missing_books_font
        else:
            status_cell.fill = missing_portal_fill
            status_cell.font = missing_portal_font

        for c_idx in range(1, len(v_headers) + 1):
            cell = ws_vendor.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if c_idx in range(4, 19):
                cell.number_format = currency_format
                cell.alignment = Alignment(horizontal="right")
            elif c_idx in [19, 20]:
                cell.alignment = Alignment(horizontal="center")

    autofit_columns(ws_vendor)

    # ---------------------------------------------------------
    # TAB 3: CHECK 2 - BOOKS VS PORTAL (STRICTLY SIDE-BY-SIDE)
    # ---------------------------------------------------------
    ws_bvp = wb.create_sheet(title="2_Books_vs_Portal")
    ws_bvp.views.sheetView[0].showGridLines = True
    b_headers = [
        "Supplier GSTIN", "Vendor Name",
        "Books Bill No", "Portal Doc No",  # Side-by-Side Invoices!
        "Match Status", "Match Reason / Typo Basis",  # Explainable Basis!
        "Books Inv Date", "Portal Doc Date",  # Side-by-Side Dates!
        "Books Taxable", "Portal Taxable", "Taxable Diff",
        "Books IGST", "Portal IGST", "IGST Diff",
        "Books CGST", "Portal CGST", "CGST Diff",
        "Books SGST", "Portal SGST", "SGST Diff",
        "Books Total Tax", "Portal Total Tax", "Total Tax Diff",
        "Branch", "SAP Trans No", "Posting Date"
    ]
    apply_header_style(ws_bvp, b_headers)

    for r in reconciler.books_recon:
        ws_bvp.append([
            r['gstin'], r['vendor_name'],
            r['bill_no'], r['portal_doc_no'],
            r['match_status'], r['match_reason'],
            r['inv_date'], r['portal_doc_date'],
            r['taxable'], r['portal_taxable'], r['taxable_diff'],
            r['igst'], r['portal_igst'], r['igst_diff'],
            r['cgst'], r['portal_cgst'], r['cgst_diff'],
            r['sgst'], r['portal_sgst'], r['sgst_diff'],
            r['total_tax'], r['portal_total_tax'], r['tax_diff'],
            r['branch'], r['sap_trans_no'], r['posting_date']
        ])
        curr_row = ws_bvp.max_row
        ws_bvp.row_dimensions[curr_row].height = 20

        st = r['match_status']
        status_cell = ws_bvp.cell(row=curr_row, column=5)
        if st == "Matched (Exact)":
            status_cell.fill = matched_fill
            status_cell.font = matched_font
        elif st == "Matched (Smart/Typo)":
            status_cell.fill = smart_fill
            status_cell.font = smart_font
        elif "Head Mismatch" in st:
            status_cell.fill = pos_mismatch_fill
            status_cell.font = pos_mismatch_font
        elif st == "Value Mismatch":
            status_cell.fill = mismatch_fill
            status_cell.font = mismatch_font
        else:
            status_cell.fill = missing_books_fill
            status_cell.font = missing_books_font

        # Highlight Portal Doc No column when matched via smart/typo
        if st == "Matched (Smart/Typo)":
            ws_bvp.cell(row=curr_row, column=4).font = Font(name="Calibri", size=10, bold=True, color="0C5460")

        for c_idx in range(1, len(b_headers) + 1):
            cell = ws_bvp.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if c_idx in range(9, 24):
                cell.number_format = currency_format
                cell.alignment = Alignment(horizontal="right")

    autofit_columns(ws_bvp)

    # ---------------------------------------------------------
    # TAB 4: CHECK 2 - PORTAL VS BOOKS (STRICTLY SIDE-BY-SIDE)
    # ---------------------------------------------------------
    ws_pvb = wb.create_sheet(title="3_Portal_vs_Books")
    ws_pvb.views.sheetView[0].showGridLines = True
    p_headers = [
        "Supplier GSTIN", "Supplier Name",
        "Portal Doc No", "Books Bill No",  # Side-by-Side Invoices!
        "Match Status", "Match Reason / Typo Basis",  # Explainable Basis!
        "Portal Doc Date", "Books Inv Date",  # Side-by-Side Dates!
        "Portal Taxable", "Books Taxable", "Taxable Diff",
        "Portal IGST", "Books IGST", "IGST Diff",
        "Portal CGST", "Books CGST", "CGST Diff",
        "Portal SGST", "Books SGST", "SGST Diff",
        "Portal Total Tax", "Books Total Tax", "Total Tax Diff",
        "Reverse Charge"
    ]
    apply_header_style(ws_pvb, p_headers)

    for r in reconciler.portal_recon:
        ws_pvb.append([
            r['gstin'], r['supplier_name'],
            r['doc_no'], r['books_bill_no'],
            r['match_status'], r['match_reason'],
            r['doc_date'], r['books_inv_date'],
            r['taxable'], r['books_taxable'], r['taxable_diff'],
            r['igst'], r['books_igst'], r['igst_diff'],
            r['cgst'], r['books_cgst'], r['cgst_diff'],
            r['sgst'], r['books_sgst'], r['sgst_diff'],
            r['total_tax'], r['books_total_tax'], r['tax_diff'],
            r['rc']
        ])
        curr_row = ws_pvb.max_row
        ws_pvb.row_dimensions[curr_row].height = 20

        st = r['match_status']
        status_cell = ws_pvb.cell(row=curr_row, column=5)
        if st == "Matched (Exact)":
            status_cell.fill = matched_fill
            status_cell.font = matched_font
        elif st == "Matched (Smart/Typo)":
            status_cell.fill = smart_fill
            status_cell.font = smart_font
        elif "Head Mismatch" in st:
            status_cell.fill = pos_mismatch_fill
            status_cell.font = pos_mismatch_font
        elif st == "Value Mismatch":
            status_cell.fill = mismatch_fill
            status_cell.font = mismatch_font
        else:
            status_cell.fill = missing_portal_fill
            status_cell.font = missing_portal_font

        for c_idx in range(1, len(p_headers) + 1):
            cell = ws_pvb.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if c_idx in range(9, 24):
                cell.number_format = currency_format
                cell.alignment = Alignment(horizontal="right")

    autofit_columns(ws_pvb)

    # ---------------------------------------------------------
    # TAB 5: ACTIONABLE DISCREPANCY & FOLLOW-UP LIST
    # ---------------------------------------------------------
    ws_act = wb.create_sheet(title="4_Actionable_Discrepancies")
    ws_act.views.sheetView[0].showGridLines = True
    act_headers = [
        "Category", "Supplier GSTIN", "Party Name",
        "Books Bill No", "Portal Doc No", "Date",
        "Taxable Value", "Total Tax", "Heads Breakup (IGST/CGST/SGST)",
        "Audit Impact / Variance", "Recommended Action"
    ]
    apply_header_style(ws_act, act_headers)

    for r in reconciler.actionable_list:
        ws_act.append([
            r['action_type'], r['gstin'], r['party_name'],
            r['books_bill_no'], r['portal_doc_no'], r['invoice_date'],
            r['taxable_value'], r['total_tax'], r['heads_breakup'],
            r['impact'], r['recommended_action']
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
        elif "Head" in r['action_type']:
            cat_cell.fill = pos_mismatch_fill
            cat_cell.font = pos_mismatch_font
        else:
            cat_cell.fill = mismatch_fill
            cat_cell.font = mismatch_font

        for c_idx in range(1, len(act_headers) + 1):
            cell = ws_act.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if c_idx in [7, 8]:
                cell.number_format = currency_format
                cell.alignment = Alignment(horizontal="right")

    autofit_columns(ws_act)

    wb.save(output_path)
    print(f"[OK] Workbook successfully saved to: {output_path}")


# -------------------------------------------------------------
# 5. CLI EXECUTION ENTRYPOINT
# -------------------------------------------------------------

def main():
    import sys
    books_file = r"C:\Users\lenovo\Downloads\Combined PR GST REPORT_06062023 Creation date upto 17-Apr-2026_Books.xlsx"
    portal_file = r"C:\Users\lenovo\Downloads\Bangalore IOT_GST portal.xlsx"
    output_file = r"C:\Users\lenovo\Downloads\Bangalore_IOT_GST_Reconciliation_Report.xlsx"

    print("=== V. SINGHI & ASSOCIATES GST RECONCILIATION ENGINE ===")
    books = load_books_data(books_file, branch_filter="Bangalore IOT")
    portal = load_portal_data(portal_file)

    reconciler = GSTReconciler(books, portal, tolerance=2.00)
    reconciler.run_reconciliation()
    export_reconciliation_workbook(reconciler, output_file)
    print(f"[SUCCESS] Reconciliation completed! Output generated at:\n  {output_file}")


if __name__ == "__main__":
    main()
