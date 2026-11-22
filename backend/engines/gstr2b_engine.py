"""
GSTR-2B vs Purchase Register Reconciliation Engine:
High-speed matching engine comparing accountant's Purchase Register against GST Portal GSTR-2B.
"""
import io
import json
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from typing import Dict, Any, List, Optional

def normalize_inv_number(inv_str: Any) -> str:
    """Normalizes invoice numbers by stripping non-alphanumeric chars and leading zeroes."""
    if pd.isna(inv_str):
        return ""
    clean = re.sub(r'[^a-zA-Z0-9]', '', str(inv_str)).upper()
    return clean.lstrip('0') or clean

def parse_purchase_register(file_content_or_path: Any) -> pd.DataFrame:
    """Parses Excel or CSV purchase register from Tally, Zoho, Busy, etc."""
    if isinstance(file_content_or_path, str) and file_content_or_path.endswith('.csv'):
        df = pd.read_csv(file_content_or_path)
    else:
        df = pd.read_excel(file_content_or_path)

    # Normalize column names
    col_dict = {}
    for col in df.columns:
        c_clean = str(col).strip().lower().replace(' ', '_').replace('/', '_').replace('.', '')
        col_dict[c_clean] = col

    def get_col(possible_names):
        for name in possible_names:
            if name in col_dict:
                return col_dict[name]
        return None

    gstin_col = get_col(['gstin', 'supplier_gstin', 'vendor_gstin', 'party_gstin', 'tin'])
    name_col = get_col(['supplier_name', 'vendor_name', 'party_name', 'name', 'particulars'])
    inv_col = get_col(['invoice_number', 'invoice_no', 'inv_no', 'bill_no', 'voucher_no'])
    date_col = get_col(['invoice_date', 'bill_date', 'date', 'voucher_date'])
    taxable_col = get_col(['taxable_value', 'taxable_amount', 'net_amount', 'taxable_val', 'assessable_value'])
    igst_col = get_col(['igst', 'integrated_tax', 'igst_amount', 'igst_amt'])
    cgst_col = get_col(['cgst', 'central_tax', 'cgst_amount', 'cgst_amt'])
    sgst_col = get_col(['sgst', 'state_tax', 'sgst_amount', 'sgst_amt', 'utgst'])
    total_col = get_col(['invoice_value', 'total_amount', 'gross_amount', 'inv_val', 'total'])

    standard_df = pd.DataFrame()
    standard_df['supplier_gstin'] = df[gstin_col].astype(str).str.strip().str.upper() if gstin_col else 'UNKNOWN'
    standard_df['supplier_name'] = df[name_col].astype(str).str.strip() if name_col else ''
    standard_df['invoice_number'] = df[inv_col].astype(str).str.strip() if inv_col else ''
    standard_df['invoice_date'] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%d-%m-%Y') if date_col else ''
    standard_df['taxable_value'] = pd.to_numeric(df[taxable_col], errors='coerce').fillna(0) if taxable_col else 0.0
    standard_df['igst'] = pd.to_numeric(df[igst_col], errors='coerce').fillna(0) if igst_col else 0.0
    standard_df['cgst'] = pd.to_numeric(df[cgst_col], errors='coerce').fillna(0) if cgst_col else 0.0
    standard_df['sgst'] = pd.to_numeric(df[sgst_col], errors='coerce').fillna(0) if sgst_col else 0.0
    standard_df['total_tax'] = standard_df['igst'] + standard_df['cgst'] + standard_df['sgst']
    standard_df['total_value'] = pd.to_numeric(df[total_col], errors='coerce').fillna(standard_df['taxable_value'] + standard_df['total_tax']) if total_col else (standard_df['taxable_value'] + standard_df['total_tax'])
    standard_df['norm_inv'] = standard_df['invoice_number'].apply(normalize_inv_number)

    return standard_df

def parse_gstr2b(file_content_or_path: Any) -> pd.DataFrame:
    """Parses official GSTR-2B JSON or Excel file downloaded from GST Portal."""
    records = []

    # Check if JSON
    is_json = False
    if isinstance(file_content_or_path, str) and file_content_or_path.endswith('.json'):
        is_json = True
    elif isinstance(file_content_or_path, (bytes, io.BytesIO)):
        try:
            content = file_content_or_path.getvalue() if isinstance(file_content_or_path, io.BytesIO) else file_content_or_path
            data = json.loads(content.decode('utf-8'))
            is_json = True
        except Exception:
            is_json = False

    if is_json:
        if isinstance(file_content_or_path, str):
            with open(file_content_or_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        
        # Traverse GSTR-2B JSON schema
        data_block = data.get('data', data)
        docdata = data_block.get('docdata', data_block)
        
        # Table 3: B2B Invoices
        b2b_list = docdata.get('b2b', [])
        for supplier in b2b_list:
            sup_gstin = str(supplier.get('ctin', '')).strip().upper()
            cname = str(supplier.get('cname', ''))
            for inv in supplier.get('inv', []):
                inv_no = str(inv.get('inum', '')).strip()
                inv_date = str(inv.get('dt', ''))
                val = float(inv.get('val', 0.0))
                itc_avl = inv.get('itcavl', 'Y')
                
                # Sum items
                txval = 0.0
                igst = 0.0
                cgst = 0.0
                sgst = 0.0
                for item in inv.get('items', []):
                    txval += float(item.get('txval', 0.0))
                    igst += float(item.get('iamt', 0.0))
                    cgst += float(item.get('camt', 0.0))
                    sgst += float(item.get('samt', 0.0))
                
                records.append({
                    'supplier_gstin': sup_gstin,
                    'supplier_name': cname,
                    'invoice_number': inv_no,
                    'invoice_date': inv_date,
                    'taxable_value': round(txval, 2),
                    'igst': round(igst, 2),
                    'cgst': round(cgst, 2),
                    'sgst': round(sgst, 2),
                    'total_tax': round(igst + cgst + sgst, 2),
                    'total_value': round(val, 2),
                    'itc_availability': itc_avl,
                    'norm_inv': normalize_inv_number(inv_no)
                })
        return pd.DataFrame(records)

    # Else parse Excel
    xls = pd.ExcelFile(file_content_or_path)
    sheet_name = 'B2B' if 'B2B' in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name, header=4 if 'B2B' in xls.sheet_names else 0)
    
    # Clean headers
    df.columns = [str(c).strip().lower().replace(' ', '_').replace('/', '_') for c in df.columns]
    
    col_map = {
        'gstin': [c for c in df.columns if 'gstin' in c or 'ctin' in c],
        'name': [c for c in df.columns if 'name' in c or 'trade' in c or 'legal' in c],
        'inv': [c for c in df.columns if 'invoice_number' in c or 'inum' in c or 'inv' in c],
        'date': [c for c in df.columns if 'invoice_date' in c or 'dt' in c or 'date' in c],
        'val': [c for c in df.columns if 'invoice_value' in c or 'val' in c],
        'txval': [c for c in df.columns if 'taxable_value' in c or 'txval' in c],
        'igst': [c for c in df.columns if 'integrated_tax' in c or 'iamt' in c or 'igst' in c],
        'cgst': [c for c in df.columns if 'central_tax' in c or 'camt' in c or 'cgst' in c],
        'sgst': [c for c in df.columns if 'state' in c or 'samt' in c or 'sgst' in c]
    }

    res_df = pd.DataFrame()
    res_df['supplier_gstin'] = df[col_map['gstin'][0]].astype(str).str.strip().str.upper() if col_map['gstin'] else ''
    res_df['supplier_name'] = df[col_map['name'][0]].astype(str).str.strip() if col_map['name'] else ''
    res_df['invoice_number'] = df[col_map['inv'][0]].astype(str).str.strip() if col_map['inv'] else ''
    res_df['invoice_date'] = df[col_map['date'][0]].astype(str) if col_map['date'] else ''
    res_df['taxable_value'] = pd.to_numeric(df[col_map['txval'][0]], errors='coerce').fillna(0) if col_map['txval'] else 0.0
    res_df['igst'] = pd.to_numeric(df[col_map['igst'][0]], errors='coerce').fillna(0) if col_map['igst'] else 0.0
    res_df['cgst'] = pd.to_numeric(df[col_map['cgst'][0]], errors='coerce').fillna(0) if col_map['cgst'] else 0.0
    res_df['sgst'] = pd.to_numeric(df[col_map['sgst'][0]], errors='coerce').fillna(0) if col_map['sgst'] else 0.0
    res_df['total_tax'] = res_df['igst'] + res_df['cgst'] + res_df['sgst']
    res_df['total_value'] = pd.to_numeric(df[col_map['val'][0]], errors='coerce').fillna(res_df['taxable_value'] + res_df['total_tax']) if col_map['val'] else (res_df['taxable_value'] + res_df['total_tax'])
    res_df['itc_availability'] = 'Y'
    res_df['norm_inv'] = res_df['invoice_number'].apply(normalize_inv_number)

    return res_df

def reconcile_gstr2b_vs_books(
    purchase_file: Any,
    gstr2b_file: Any,
    tolerance: float = 1.0
) -> Dict[str, Any]:
    """
    Performs end-to-end reconciliation between Purchase Register and GSTR-2B.
    Returns executive metrics and matched/mismatched record sets.
    """
    df_books = parse_purchase_register(purchase_file)
    df_2b = parse_gstr2b(gstr2b_file)

    matched = []
    mismatched = []
    missing_in_2b = []
    missing_in_books = []

    matched_2b_indices = set()

    # Match each record in Books
    for b_idx, b_row in df_books.iterrows():
        b_gstin = b_row['supplier_gstin']
        b_norm_inv = b_row['norm_inv']
        b_txval = float(b_row['taxable_value'])
        b_tax = float(b_row['total_tax'])

        # Candidate matches in 2B
        candidates = df_2b[(df_2b['supplier_gstin'] == b_gstin) & (df_2b['norm_inv'] == b_norm_inv)]

        if not candidates.empty:
            best_match = None
            best_diff = float('inf')
            best_idx = None

            for c_idx, c_row in candidates.iterrows():
                if c_idx in matched_2b_indices:
                    continue
                diff = abs(b_txval - float(c_row['taxable_value'])) + abs(b_tax - float(c_row['total_tax']))
                if diff < best_diff:
                    best_diff = diff
                    best_match = c_row
                    best_idx = c_idx

            if best_match is not None and best_diff <= tolerance:
                matched_2b_indices.add(best_idx)
                matched.append({
                    'status': 'MATCHED',
                    'supplier_gstin': b_gstin,
                    'supplier_name': b_row['supplier_name'] or best_match['supplier_name'],
                    'invoice_number': b_row['invoice_number'],
                    'invoice_date': b_row['invoice_date'],
                    'portal_inv_date': best_match['invoice_date'],
                    'books_taxable': b_txval,
                    'portal_taxable': float(best_match['taxable_value']),
                    'books_tax': b_tax,
                    'portal_tax': float(best_match['total_tax']),
                    'itc_status': 'Eligible for GSTR-3B Claim'
                })
            elif best_match is not None:
                matched_2b_indices.add(best_idx)
                mismatched.append({
                    'status': 'VALUE_MISMATCH',
                    'supplier_gstin': b_gstin,
                    'supplier_name': b_row['supplier_name'] or best_match['supplier_name'],
                    'invoice_number': b_row['invoice_number'],
                    'books_taxable': b_txval,
                    'portal_taxable': float(best_match['taxable_value']),
                    'taxable_diff': round(b_txval - float(best_match['taxable_value']), 2),
                    'books_tax': b_tax,
                    'portal_tax': float(best_match['total_tax']),
                    'tax_diff': round(b_tax - float(best_match['total_tax']), 2),
                    'action': 'Discrepancy with Supplier'
                })
            else:
                missing_in_2b.append({
                    'status': 'MISSING_IN_2B',
                    'supplier_gstin': b_gstin,
                    'supplier_name': b_row['supplier_name'],
                    'invoice_number': b_row['invoice_number'],
                    'invoice_date': b_row['invoice_date'],
                    'taxable_value': b_txval,
                    'tax_amount': b_tax,
                    'risk': 'Cannot Claim ITC under Rule 36(4)',
                    'action': 'Notify Vendor to File GSTR-1'
                })
        else:
            missing_in_2b.append({
                'status': 'MISSING_IN_2B',
                'supplier_gstin': b_gstin,
                'supplier_name': b_row['supplier_name'],
                'invoice_number': b_row['invoice_number'],
                'invoice_date': b_row['invoice_date'],
                'taxable_value': b_txval,
                'tax_amount': b_tax,
                'risk': 'Cannot Claim ITC under Rule 36(4)',
                'action': 'Notify Vendor to File GSTR-1'
            })

    # Unmatched records in 2B (Missing in Books)
    for c_idx, c_row in df_2b.iterrows():
        if c_idx not in matched_2b_indices:
            missing_in_books.append({
                'status': 'MISSING_IN_BOOKS',
                'supplier_gstin': c_row['supplier_gstin'],
                'supplier_name': c_row['supplier_name'],
                'invoice_number': c_row['invoice_number'],
                'invoice_date': c_row['invoice_date'],
                'taxable_value': float(c_row['taxable_value']),
                'tax_amount': float(c_row['total_tax']),
                'opportunity': 'Unclaimed ITC available on portal',
                'action': 'Book Bill in Accounting Software'
            })

    # Metrics
    total_books_itc = round(float(df_books['total_tax'].sum()), 2)
    total_2b_itc = round(float(df_2b['total_tax'].sum()), 2)
    matched_itc = round(sum(m['books_tax'] for m in matched), 2)
    missing_2b_itc = round(sum(m['tax_amount'] for m in missing_in_2b), 2)
    unclaimed_2b_itc = round(sum(m['tax_amount'] for m in missing_in_books), 2)

    return {
        'summary': {
            'total_books_invoices': len(df_books),
            'total_2b_invoices': len(df_2b),
            'matched_count': len(matched),
            'mismatched_count': len(mismatched),
            'missing_in_2b_count': len(missing_in_2b),
            'missing_in_books_count': len(missing_in_books),
            'total_books_itc': total_books_itc,
            'total_2b_itc': total_2b_itc,
            'matched_itc': matched_itc,
            'missing_2b_itc_at_risk': missing_2b_itc,
            'unclaimed_itc_in_portal': unclaimed_2b_itc,
            'itc_match_percentage': round((matched_itc / total_books_itc * 100), 1) if total_books_itc > 0 else 100.0
        },
        'matched': matched,
        'mismatched': mismatched,
        'missing_in_2b': missing_in_2b,
        'missing_in_books': missing_in_books
    }
