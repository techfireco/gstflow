"""
Meesho Report Parser:
Processes Meesho's `tcs_sales.xlsx`, `tcs_sales_return.xlsx`, and `Tax_invoice_details.xlsx`.
"""
import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional
from backend.config import normalize_state_code, GST_STATE_CODES, ECO_GSTINS

def parse_meesho_reports(
    sales_file: Any,
    return_file: Any,
    invoice_details_file: Optional[Any] = None,
    seller_gstin: Optional[str] = None
) -> Dict[str, Any]:
    """
    Parses Meesho sales and return files and extracts data for GSTR-1 and Tally XML.
    """
    df_sales = pd.read_excel(sales_file)
    df_return = pd.read_excel(return_file)

    # 1. Detect seller GSTIN if not provided
    if not seller_gstin:
        if 'gstin' in df_sales.columns and not df_sales['gstin'].dropna().empty:
            seller_gstin = str(df_sales['gstin'].dropna().iloc[0]).strip().upper()
        elif 'gstin' in df_return.columns and not df_return['gstin'].dropna().empty:
            seller_gstin = str(df_return['gstin'].dropna().iloc[0]).strip().upper()
        else:
            seller_gstin = "07XXXXXXXXXXXZX"

    seller_state_code = seller_gstin[:2] if len(seller_gstin) >= 2 else "07"

    # ECO GSTIN
    eco_gstin = ECO_GSTINS['MEESHO']['default']
    if 'eco_tcs_gstin' in df_sales.columns and not df_sales['eco_tcs_gstin'].dropna().empty:
        found_eco = str(df_sales['eco_tcs_gstin'].dropna().iloc[0]).strip().upper()
        if len(found_eco) == 15:
            eco_gstin = found_eco

    # 2. Normalize Sales & Return columns
    for df in [df_sales, df_return]:
        df['state_code'] = df['end_customer_state_new'].apply(normalize_state_code)
        df['pos_name'] = df['state_code'].map(GST_STATE_CODES).fillna('97-Other Territory')
        df['gst_rate'] = pd.to_numeric(df['gst_rate'], errors='coerce').fillna(0)
        df['total_taxable_sale_value'] = pd.to_numeric(df['total_taxable_sale_value'], errors='coerce').fillna(0)
        df['tax_amount'] = pd.to_numeric(df['tax_amount'], errors='coerce').fillna(0)
        df['quantity'] = pd.to_numeric(df['quantity'], errors='coerce').fillna(0)
        df['hsn_code'] = df['hsn_code'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    # 3. State & Rate Aggregation (B2CS & Nil Rated)
    sales_grp = df_sales.groupby(['state_code', 'pos_name', 'gst_rate']).agg({
        'total_taxable_sale_value': 'sum',
        'tax_amount': 'sum',
        'quantity': 'sum'
    }).reset_index()

    ret_grp = df_return.groupby(['state_code', 'pos_name', 'gst_rate']).agg({
        'total_taxable_sale_value': 'sum',
        'tax_amount': 'sum',
        'quantity': 'sum'
    }).reset_index()

    merged_b2cs = pd.merge(
        sales_grp, ret_grp,
        on=['state_code', 'pos_name', 'gst_rate'],
        how='outer',
        suffixes=('_sales', '_ret')
    ).fillna(0)

    merged_b2cs['net_taxable'] = merged_b2cs['total_taxable_sale_value_sales'] - merged_b2cs['total_taxable_sale_value_ret']
    merged_b2cs['net_tax'] = merged_b2cs['tax_amount_sales'] - merged_b2cs['tax_amount_ret']
    merged_b2cs['net_qty'] = merged_b2cs['quantity_sales'] - merged_b2cs['quantity_ret']

    # Filter into B2CS (rate > 0) and Nil/Exempt (rate == 0)
    b2cs_records = []
    nil_intra = 0.0
    nil_inter = 0.0

    eco_net_taxable = 0.0
    eco_igst = 0.0
    eco_cgst = 0.0
    eco_sgst = 0.0

    for _, row in merged_b2cs.iterrows():
        pos_code = row['state_code']
        pos_name = row['pos_name']
        rate = float(row['gst_rate'])
        txval = round(float(row['net_taxable']), 2)
        tax = round(float(row['net_tax']), 2)

        if rate == 0:
            if pos_code == seller_state_code:
                nil_intra += txval
            else:
                nil_inter += txval
            eco_net_taxable += txval
            continue

        is_intra = (pos_code == seller_state_code)
        sply_ty = "INTRA" if is_intra else "INTER"
        
        iamt = 0.0
        camt = 0.0
        samt = 0.0

        if is_intra:
            half_tax = round(tax / 2, 2)
            camt = half_tax
            samt = half_tax
            eco_cgst += camt
            eco_sgst += samt
        else:
            iamt = tax
            eco_igst += iamt

        eco_net_taxable += txval

        b2cs_records.append({
            'sply_ty': sply_ty,
            'pos': pos_code,
            'pos_name': pos_name,
            'rt': int(rate) if rate.is_integer() else rate,
            'typ': 'OE',
            'txval': txval,
            'iamt': iamt,
            'camt': camt,
            'samt': samt,
            'csamt': 0.0
        })

    # 4. HSN Summary (Table 12)
    hsn_sales = df_sales.groupby(['hsn_code', 'gst_rate', 'state_code']).agg({
        'quantity': 'sum',
        'total_taxable_sale_value': 'sum',
        'tax_amount': 'sum'
    }).reset_index()

    hsn_ret = df_return.groupby(['hsn_code', 'gst_rate', 'state_code']).agg({
        'quantity': 'sum',
        'total_taxable_sale_value': 'sum',
        'tax_amount': 'sum'
    }).reset_index()

    hsn_merged = pd.merge(
        hsn_sales, hsn_ret,
        on=['hsn_code', 'gst_rate', 'state_code'],
        how='outer',
        suffixes=('_sales', '_ret')
    ).fillna(0)

    hsn_merged['net_qty'] = hsn_merged['quantity_sales'] - hsn_merged['quantity_ret']
    hsn_merged['net_txval'] = hsn_merged['total_taxable_sale_value_sales'] - hsn_merged['total_taxable_sale_value_ret']
    hsn_merged['net_tax'] = hsn_merged['tax_amount_sales'] - hsn_merged['tax_amount_ret']

    hsn_dict = {}
    for _, row in hsn_merged.iterrows():
        hsn = str(row['hsn_code']).strip()
        if not hsn or hsn.lower() == 'nan':
            continue
        rate = float(row['gst_rate'])
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
        
        is_intra = (row['state_code'] == seller_state_code)
        hsn_dict[key]['qty'] += int(row['net_qty'])
        hsn_dict[key]['txval'] += float(row['net_txval'])
        
        if is_intra:
            half = round(float(row['net_tax']) / 2, 2)
            hsn_dict[key]['camt'] += half
            hsn_dict[key]['samt'] += half
        else:
            hsn_dict[key]['iamt'] += float(row['net_tax'])

    hsn_records = []
    for key, val in hsn_dict.items():
        val['txval'] = round(val['txval'], 2)
        val['iamt'] = round(val['iamt'], 2)
        val['camt'] = round(val['camt'], 2)
        val['samt'] = round(val['samt'], 2)
        val['total_val'] = round(val['txval'] + val['iamt'] + val['camt'] + val['samt'], 2)
        if val['txval'] > 0 or val['qty'] > 0:
            hsn_records.append(val)

    # 5. Documents Issued (Table 13)
    docs_records = []
    if invoice_details_file is not None:
        try:
            df_inv = pd.read_excel(invoice_details_file)
            for doc_type, group in df_inv.groupby('Type'):
                inv_nums = group['Invoice No.'].dropna().astype(str).tolist()
                if not inv_nums:
                    continue
                # Sort to get min and max
                inv_nums.sort()
                from_num = inv_nums[0]
                to_num = inv_nums[-1]
                tot_num = len(inv_nums)
                
                nature = "Invoices for outward supply" if doc_type == "INVOICE" else "Credit Note"
                doc_num = 1 if doc_type == "INVOICE" else 5
                
                docs_records.append({
                    'doc_num': doc_num,
                    'doc_typ': nature,
                    'sub_type': doc_type,
                    'from': from_num,
                    'to': to_num,
                    'totnum': tot_num,
                    'cancel': 0,
                    'net_issue': tot_num
                })
        except Exception as e:
            print(f"Warning parsing Meesho invoice details: {e}")

    return {
        'marketplace': 'Meesho',
        'seller_gstin': seller_gstin,
        'eco_gstin': eco_gstin,
        'eco_name': 'meesho',
        'b2cs': b2cs_records,
        'nil_intra': round(nil_intra, 2),
        'nil_inter': round(nil_inter, 2),
        'hsn': hsn_records,
        'docs': docs_records,
        'eco_summary': {
            'nature': 'Liable to collect tax u/s 52(TCS)',
            'etin': eco_gstin,
            'name': 'meesho',
            'suppval': round(eco_net_taxable, 2),
            'igst': round(eco_igst, 2),
            'cgst': round(eco_cgst, 2),
            'sgst': round(eco_sgst, 2),
            'cess': 0.0
        },
        'raw_counts': {
            'sales_orders': len(df_sales),
            'return_orders': len(df_return)
        }
    }
