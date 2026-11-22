"""
Amazon Report Parser:
Processes Amazon's Merchant Tax Report (MTR B2C CSV).
"""
import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional
from backend.config import normalize_state_code, GST_STATE_CODES, ECO_GSTINS

def parse_amazon_mtr(
    csv_file: Any,
    seller_gstin: Optional[str] = None
) -> Dict[str, Any]:
    """
    Parses Amazon MTR B2C CSV report for GSTR-1 and Tally XML generation.
    """
    df = pd.read_csv(csv_file, low_memory=False)

    # 1. Detect seller GSTIN
    if not seller_gstin:
        if 'Seller Gstin' in df.columns and not df['Seller Gstin'].dropna().empty:
            seller_gstin = str(df['Seller Gstin'].dropna().iloc[0]).strip().upper()
        else:
            seller_gstin = "07XXXXXXXXXXXZX"

    seller_state_code = seller_gstin[:2] if len(seller_gstin) >= 2 else "07"
    eco_gstin = ECO_GSTINS['AMAZON']['default']

    # Filter out Cancel and FreeReplacement (taxable impact is zero)
    valid_df = df[df['Transaction Type'].isin(['Shipment', 'Refund'])].copy()

    # Normalize numeric columns
    num_cols = [
        'Quantity', 'Tax Exclusive Gross', 'Total Tax Amount',
        'Cgst Rate', 'Sgst Rate', 'Igst Rate',
        'Cgst Tax', 'Sgst Tax', 'Igst Tax', 'Invoice Amount'
    ]
    for col in num_cols:
        if col in valid_df.columns:
            valid_df[col] = pd.to_numeric(valid_df[col], errors='coerce').fillna(0)

    # State Normalization
    valid_df['state_code'] = valid_df['Ship To State'].apply(normalize_state_code)
    valid_df['pos_name'] = valid_df['state_code'].map(GST_STATE_CODES).fillna('97-Other Territory')

    # Overall GST Rate: IGST rate if inter, or (CGST rate + SGST rate) if intra
    def calc_rate(row):
        if row['Igst Rate'] > 0:
            return round(row['Igst Rate'] * 100, 2)
        elif (row['Cgst Rate'] + row['Sgst Rate']) > 0:
            return round((row['Cgst Rate'] + row['Sgst Rate']) * 100, 2)
        return 0.0

    valid_df['gst_rate'] = valid_df.apply(calc_rate, axis=1)

    # 2. B2CS Aggregation (by state & rate)
    b2cs_dict = {}
    nil_intra = 0.0
    nil_inter = 0.0

    eco_net_taxable = 0.0
    eco_igst = 0.0
    eco_cgst = 0.0
    eco_sgst = 0.0

    for _, row in valid_df.iterrows():
        pos_code = row['state_code']
        pos_name = row['pos_name']
        rate = float(row['gst_rate'])
        txval = float(row['Tax Exclusive Gross'])
        igst = float(row['Igst Tax'])
        cgst = float(row['Cgst Tax'])
        sgst = float(row['Sgst Tax'])

        if rate == 0:
            if pos_code == seller_state_code:
                nil_intra += txval
            else:
                nil_inter += txval
            eco_net_taxable += txval
            continue

        key = (pos_code, pos_name, rate)
        if key not in b2cs_dict:
            is_intra = (pos_code == seller_state_code)
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
        b2cs_dict[key]['iamt'] += igst
        b2cs_dict[key]['camt'] += cgst
        b2cs_dict[key]['samt'] += sgst

        eco_net_taxable += txval
        eco_igst += igst
        eco_cgst += cgst
        eco_sgst += sgst

    b2cs_records = []
    for key, item in b2cs_dict.items():
        item['txval'] = round(item['txval'], 2)
        item['iamt'] = round(item['iamt'], 2)
        item['camt'] = round(item['camt'], 2)
        item['samt'] = round(item['samt'], 2)
        b2cs_records.append(item)

    # 3. HSN Summary (Table 12)
    hsn_dict = {}
    for _, row in valid_df.iterrows():
        raw_hsn = str(row.get('Hsn/sac', '')).strip()
        if not raw_hsn or raw_hsn.lower() in ['nan', 'none', '']:
            # If HSN is missing, fallback or assign standard e-commerce HSN
            raw_hsn = '4820'
        
        # Clean HSN formatting (strip .0)
        hsn = raw_hsn.replace('.0', '')
        rate = float(row['gst_rate'])
        qty = int(row['Quantity'])
        txval = float(row['Tax Exclusive Gross'])
        igst = float(row['Igst Tax'])
        cgst = float(row['Cgst Tax'])
        sgst = float(row['Sgst Tax'])

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
        hsn_dict[key]['iamt'] += igst
        hsn_dict[key]['camt'] += cgst
        hsn_dict[key]['samt'] += sgst

    hsn_records = []
    for key, val in hsn_dict.items():
        val['txval'] = round(val['txval'], 2)
        val['iamt'] = round(val['iamt'], 2)
        val['camt'] = round(val['camt'], 2)
        val['samt'] = round(val['samt'], 2)
        val['total_val'] = round(val['txval'] + val['iamt'] + val['camt'] + val['samt'], 2)
        if abs(val['txval']) > 0 or abs(val['qty']) > 0:
            hsn_records.append(val)

    # 4. Document Series (Table 13)
    docs_records = []
    # Invoices for outward supply (Shipments)
    ship_invs = valid_df[valid_df['Transaction Type'] == 'Shipment']['Invoice Number'].dropna().astype(str).tolist()
    if ship_invs:
        ship_invs.sort()
        docs_records.append({
            'doc_num': 1,
            'doc_typ': 'Invoices for outward supply',
            'sub_type': 'AMAZON_SHIPMENT',
            'from': ship_invs[0],
            'to': ship_invs[-1],
            'totnum': len(ship_invs),
            'cancel': 0,
            'net_issue': len(ship_invs)
        })

    # Credit Notes (Refunds)
    ref_cns = valid_df[valid_df['Transaction Type'] == 'Refund']['Credit Note No'].dropna().astype(str).tolist()
    if ref_cns:
        ref_cns.sort()
        docs_records.append({
            'doc_num': 5,
            'doc_typ': 'Credit Note',
            'sub_type': 'AMAZON_REFUND',
            'from': ref_cns[0],
            'to': ref_cns[-1],
            'totnum': len(ref_cns),
            'cancel': 0,
            'net_issue': len(ref_cns)
        })

    return {
        'marketplace': 'Amazon',
        'seller_gstin': seller_gstin,
        'eco_gstin': eco_gstin,
        'eco_name': 'amazon',
        'b2cs': b2cs_records,
        'nil_intra': round(nil_intra, 2),
        'nil_inter': round(nil_inter, 2),
        'hsn': hsn_records,
        'docs': docs_records,
        'eco_summary': {
            'nature': 'Liable to collect tax u/s 52(TCS)',
            'etin': eco_gstin,
            'name': 'amazon',
            'suppval': round(eco_net_taxable, 2),
            'igst': round(eco_igst, 2),
            'cgst': round(eco_cgst, 2),
            'sgst': round(eco_sgst, 2),
            'cess': 0.0
        },
        'raw_counts': {
            'shipments': len(ship_invs),
            'refunds': len(ref_cns)
        }
    }
