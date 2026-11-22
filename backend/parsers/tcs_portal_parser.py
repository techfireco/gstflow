"""
GST Portal TCS Parser & Reconciler:
Parses `07XXXXXXXXXXXXX_MONTH_YEAR.xlsx` or `.zip` downloaded from GST Portal (TDS & TCS Credit Received).
"""
import zipfile
import io
import pandas as pd
from typing import Dict, Any, List, Optional
from backend.config import normalize_state_code, GST_STATE_CODES

def parse_tcs_portal_file(file_content_or_path: Any) -> Dict[str, Any]:
    """
    Parses Table 5 (TCS Credit Received) from the GST Portal Excel/Zip file.
    """
    excel_file = None
    
    # Check if zip file
    if isinstance(file_content_or_path, (str, bytes, io.BytesIO)):
        try:
            with zipfile.ZipFile(file_content_or_path) as z:
                for name in z.namelist():
                    if name.endswith('.xlsx'):
                        excel_file = io.BytesIO(z.read(name))
                        break
        except zipfile.BadZipFile:
            excel_file = file_content_or_path
    else:
        excel_file = file_content_or_path

    xls = pd.ExcelFile(excel_file)
    if '5' not in xls.sheet_names:
        raise ValueError("Sheet '5' (TCS credit received) not found in portal file.")

    df5 = pd.read_excel(xls, sheet_name='5', header=7)
    df5.columns = [
        'etin', 'tax_period', 'gross_value', 'supplies_returned',
        'net_value', 'iamt', 'camt', 'samt', 'pos', 'remarks', 'action'
    ]
    df5 = df5.iloc[1:].dropna(subset=['etin']).copy()

    for col in ['gross_value', 'supplies_returned', 'net_value', 'iamt', 'camt', 'samt']:
        df5[col] = pd.to_numeric(df5[col], errors='coerce').fillna(0)

    df5['state_code'] = df5['pos'].apply(normalize_state_code)
    df5['pos_name'] = df5['state_code'].map(GST_STATE_CODES).fillna('97-Other Territory')

    # Summary by ECO
    eco_summary = {}
    for etin, group in df5.groupby('etin'):
        eco_summary[str(etin).strip().upper()] = {
            'gross_value': round(float(group['gross_value'].sum()), 2),
            'supplies_returned': round(float(group['supplies_returned'].sum()), 2),
            'net_value': round(float(group['net_value'].sum()), 2),
            'iamt': round(float(group['iamt'].sum()), 2),
            'camt': round(float(group['camt'].sum()), 2),
            'samt': round(float(group['samt'].sum()), 2),
            'records_count': len(group)
        }

    # State-wise details
    state_details = []
    for _, row in df5.iterrows():
        state_details.append({
            'etin': str(row['etin']).strip().upper(),
            'state_code': row['state_code'],
            'pos_name': row['pos_name'],
            'gross_value': round(float(row['gross_value']), 2),
            'supplies_returned': round(float(row['supplies_returned']), 2),
            'net_value': round(float(row['net_value']), 2),
            'iamt': round(float(row['iamt']), 2),
            'camt': round(float(row['camt']), 2),
            'samt': round(float(row['samt']), 2)
        })

    return {
        'eco_summary': eco_summary,
        'state_details': state_details,
        'total_net_value': round(float(df5['net_value'].sum()), 2),
        'total_tcs_tax': round(float(df5['iamt'].sum() + df5['camt'].sum() + df5['samt'].sum()), 2)
    }
