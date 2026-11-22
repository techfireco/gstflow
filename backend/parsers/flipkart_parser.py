"""
Flipkart Report Parser:
Processes Flipkart Sales / GSTR / Section 52 Tax reports (CSV / Excel).
"""
import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional
from backend.config import normalize_state_code, GST_STATE_CODES, ECO_GSTINS

def parse_flipkart_report(
    report_file: Any,
    seller_gstin: Optional[str] = None
) -> Dict[str, Any]:
    """
    Parses Flipkart sales and returns reports for GSTR-1 and Tally XML generation.
    """
    if str(report_file).endswith('.csv') or hasattr(report_file, 'filename') and report_file.filename.endswith('.csv'):
        df = pd.read_csv(report_file, low_memory=False)
    else:
        df = pd.read_excel(report_file)

    # Clean column headers
    df.columns = [str(c).strip().lower().replace(' ', '_').replace('/', '_') for c in df.columns]

    # Map possible column names
    col_map = {
        'seller_gstin': ['seller_gstin', 'tin', 'seller_tin', 'supplier_gstin'],
        'state': ['customer_delivery_state', 'delivery_state', 'customer_state', 'destination_state', 'state'],
        'event': ['event_type', 'event_sub_type', 'transaction_type', 'type'],
        'inv_num': ['invoice_number', 'invoice_id', 'inv_no', 'invoice_no'],
        'taxable': ['taxable_value', 'taxable_amount', 'net_taxable_value', 'item_taxable_value'],
        'igst_rate': ['igst_rate', 'igst_tax_rate', 'integrated_tax_rate'],
        'cgst_rate': ['cgst_rate', 'cgst_tax_rate', 'central_tax_rate'],
        'sgst_rate': ['sgst_rate', 'sgst_tax_rate', 'state_tax_rate'],
        'igst_amt': ['igst_amount', 'igst_amt', 'integrated_tax_amount'],
        'cgst_amt': ['cgst_amount', 'cgst_amt', 'central_tax_amount'],
        'sgst_amt': ['sgst_amount', 'sgst_amt', 'state_tax_amount'],
        'hsn': ['hsn', 'hsn_code', 'hsn_sac', 'product_hsn'],
        'qty': ['quantity', 'qty', 'item_quantity']
    }

    def find_col(target):
        for candidate in col_map.get(target, []):
            if candidate in df.columns:
                return candidate
        return None

    state_col = find_col('state') or 'delivery_state'
    event_col = find_col('event') or 'event_type'
    inv_col = find_col('inv_num') or 'invoice_number'
    taxable_col = find_col('taxable') or 'taxable_value'
    hsn_col = find_col('hsn') or 'hsn_code'
    qty_col = find_col('qty') or 'quantity'

    # Detect seller GSTIN
    gstin_col = find_col('seller_gstin')
    if not seller_gstin:
        if gstin_col and not df[gstin_col].dropna().empty:
            seller_gstin = str(df[gstin_col].dropna().iloc[0]).strip().upper()
        else:
            seller_gstin = "07XXXXXXXXXXXZX"

    seller_state_code = seller_gstin[:2] if len(seller_gstin) >= 2 else "07"
    eco_gstin = ECO_GSTINS['FLIPKART']['default']

    # Convert numeric fields
    for c in [taxable_col, qty_col]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    # State Normalization
    df['state_code'] = df[state_col].apply(normalize_state_code) if state_col in df.columns else '97'
    df['pos_name'] = df['state_code'].map(GST_STATE_CODES).fillna('97-Other Territory')

    # Detect Rate
    igst_rate_col = find_col('igst_rate')
    cgst_rate_col = find_col('cgst_rate')
    sgst_rate_col = find_col('sgst_rate')

    def extract_rate(row):
        r = 0.0
        if igst_rate_col and row.get(igst_rate_col, 0) > 0:
            r = float(row[igst_rate_col])
        elif cgst_rate_col and sgst_rate_col:
            r = float(row.get(cgst_rate_col, 0)) + float(row.get(sgst_rate_col, 0))
        # Sometimes rate is expressed as 0.05 or 5
        if 0 < r < 1:
            r = r * 100
        return round(r, 2)

    df['gst_rate'] = df.apply(extract_rate, axis=1)

    # Filter Sales vs Returns
    # In Flipkart, event_sub_type usually 'Sale' / 'Return' / 'Cancellation'
    # Taxable value can also be signed (negative for returns)
    b2cs_dict = {}
    nil_intra = 0.0
    nil_inter = 0.0

    eco_net_taxable = 0.0
    eco_igst = 0.0
    eco_cgst = 0.0
    eco_sgst = 0.0

    for _, row in df.iterrows():
        is_return = False
        if event_col in row:
            val_str = str(row[event_col]).upper()
            if 'RETURN' in val_str or 'CANCEL' in val_str or 'REFUND' in val_str:
                is_return = True

        txval = float(row.get(taxable_col, 0))
        if is_return and txval > 0:
            txval = -txval

        pos_code = row['state_code']
        pos_name = row['pos_name']
        rate = float(row['gst_rate'])

        is_intra = (pos_code == seller_state_code)
        
        # Calculate taxes
        tax = round(abs(txval) * (rate / 100), 2)
        if txval < 0:
            tax = -tax

        if is_intra:
            camt = round(tax / 2, 2)
            samt = round(tax / 2, 2)
            iamt = 0.0
        else:
            iamt = tax
            camt = 0.0
            samt = 0.0

        if rate == 0:
            if pos_code == seller_state_code:
                nil_intra += txval
            else:
                nil_inter += txval
            eco_net_taxable += txval
            continue

        key = (pos_code, pos_name, rate)
        if key not in b2cs_dict:
            b2cs_dict[key] = {
                'sply_ty': 'INTRA' if is_intra else 'INTER',
                'pos': pos_code,
                'pos_name': pos_name,
                'rt': int(rate) if rate.is_integer() else rate,
                'typ': 'OE',
                'txval': 0.0,
                'iamt': 0.0,
                'camt': 0.0,
                'samt': 0.0,
                'csamt': 0.0
            }

        b2cs_dict[key]['txval'] += txval
        b2cs_dict[key]['iamt'] += iamt
        b2cs_dict[key]['camt'] += camt
        b2cs_dict[key]['samt'] += samt

        eco_net_taxable += txval
        eco_igst += iamt
        eco_cgst += camt
        eco_sgst += samt

    b2cs_records = []
    for item in b2cs_dict.values():
        item['txval'] = round(item['txval'], 2)
        item['iamt'] = round(item['iamt'], 2)
        item['camt'] = round(item['camt'], 2)
        item['samt'] = round(item['samt'], 2)
        b2cs_records.append(item)

    # HSN Summary
    hsn_dict = {}
    for _, row in df.iterrows():
        raw_hsn = str(row.get(hsn_col, '')).strip()
        if not raw_hsn or raw_hsn.lower() in ['nan', 'none', '']:
            raw_hsn = 'Default'
        hsn = raw_hsn.replace('.0', '')
        rate = float(row['gst_rate'])
        qty = int(row.get(qty_col, 1))
        txval = float(row.get(taxable_col, 0))

        key = (hsn, rate)
        if key not in hsn_dict:
            hsn_dict[key] = {
                'hsn_sc': hsn,
                'uqc': 'PCS',
                'uqc_name': 'PCS-PIECES',
                'qty': 0,
                'rt': int(rate) if rate.is_integer() else rate,
                'txval': 0.0,
                'iamt': 0.0,
                'camt': 0.0,
                'samt': 0.0,
                'csamt': 0.0
            }

        hsn_dict[key]['qty'] += qty
        hsn_dict[key]['txval'] += txval

    hsn_records = []
    for val in hsn_dict.values():
        val['txval'] = round(val['txval'], 2)
        val['iamt'] = round(val['iamt'], 2)
        val['camt'] = round(val['camt'], 2)
        val['samt'] = round(val['samt'], 2)
        val['total_val'] = round(val['txval'] + val['iamt'] + val['camt'] + val['samt'], 2)
        if abs(val['txval']) > 0 or abs(val['qty']) > 0:
            hsn_records.append(val)

    # Documents
    docs_records = []
    if inv_col in df.columns:
        invs = df[inv_col].dropna().astype(str).tolist()
        if invs:
            invs.sort()
            docs_records.append({
                'doc_num': 1,
                'doc_typ': 'Invoices for outward supply',
                'sub_type': 'FLIPKART_INVOICE',
                'from': invs[0],
                'to': invs[-1],
                'totnum': len(invs),
                'cancel': 0,
                'net_issue': len(invs)
            })

    return {
        'marketplace': 'Flipkart',
        'seller_gstin': seller_gstin,
        'eco_gstin': eco_gstin,
        'eco_name': 'flipkart',
        'b2cs': b2cs_records,
        'nil_intra': round(nil_intra, 2),
        'nil_inter': round(nil_inter, 2),
        'hsn': hsn_records,
        'docs': docs_records,
        'eco_summary': {
            'nature': 'Liable to collect tax u/s 52(TCS)',
            'etin': eco_gstin,
            'name': 'flipkart',
            'suppval': round(eco_net_taxable, 2),
            'igst': round(eco_igst, 2),
            'cgst': round(eco_cgst, 2),
            'sgst': round(eco_sgst, 2),
            'cess': 0.0
        },
        'raw_counts': {
            'records': len(df)
        }
    }
