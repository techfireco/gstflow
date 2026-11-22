"""
GSTR-1 Generation Engine:
Merges marketplace data and generates official GST Portal GSTR-1 JSON and Excel files.
"""
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from typing import Dict, Any, List, Optional
from backend.config import GST_STATE_CODES

def generate_gstr1(
    marketplace_results: List[Dict[str, Any]],
    seller_gstin: str,
    filing_period: str, # e.g. "102026" (MMYYYY)
    output_excel_path: Optional[str] = None,
    output_json_path: Optional[str] = None
) -> Dict[str, Any]:
    """
    Combines parsed marketplace data and creates official GSTR-1 JSON and Excel.
    """
    seller_state_code = seller_gstin[:2] if len(seller_gstin) >= 2 else "07"

    # 1. Merge B2CS
    b2cs_map = {}
    total_nil_intra = 0.0
    total_nil_inter = 0.0

    for res in marketplace_results:
        # Nil rated
        total_nil_intra += res.get('nil_intra', 0.0)
        total_nil_inter += res.get('nil_inter', 0.0)

        for item in res.get('b2cs', []):
            pos = item['pos']
            rt = item['rt']
            pos_name = item.get('pos_name', GST_STATE_CODES.get(pos, '97-Other Territory'))
            key = (pos, rt)
            if key not in b2cs_map:
                is_intra = (pos == seller_state_code)
                b2cs_map[key] = {
                    'sply_ty': 'INTRA' if is_intra else 'INTER',
                    'pos': pos,
                    'pos_name': pos_name,
                    'rt': rt,
                    'typ': 'OE',
                    'txval': 0.0,
                    'iamt': 0.0,
                    'camt': 0.0,
                    'samt': 0.0,
                    'csamt': 0.0
                }
            b2cs_map[key]['txval'] += item['txval']
            b2cs_map[key]['iamt'] += item.get('iamt', 0.0)
            b2cs_map[key]['camt'] += item.get('camt', 0.0)
            b2cs_map[key]['samt'] += item.get('samt', 0.0)
            b2cs_map[key]['csamt'] += item.get('csamt', 0.0)

    # Sort B2CS records: first by POS state code, then by Rate
    merged_b2cs = []
    for key in sorted(b2cs_map.keys(), key=lambda x: (x[0], x[1])):
        item = b2cs_map[key]
        item['txval'] = round(item['txval'], 2)
        item['iamt'] = round(item['iamt'], 2)
        item['camt'] = round(item['camt'], 2)
        item['samt'] = round(item['samt'], 2)
        item['csamt'] = round(item['csamt'], 2)
        merged_b2cs.append(item)

    # 2. Merge HSN Summary
    hsn_map = {}
    for res in marketplace_results:
        for item in res.get('hsn', []):
            hsn = str(item['hsn_sc']).strip()
            rt = item['rt']
            key = (hsn, rt)
            if key not in hsn_map:
                hsn_map[key] = {
                    'hsn_sc': hsn,
                    'uqc': item.get('uqc', 'PCS'),
                    'uqc_name': item.get('uqc_name', 'PCS-PIECES'),
                    'qty': 0,
                    'rt': rt,
                    'txval': 0.0,
                    'iamt': 0.0,
                    'camt': 0.0,
                    'samt': 0.0,
                    'csamt': 0.0
                }
            hsn_map[key]['qty'] += item['qty']
            hsn_map[key]['txval'] += item['txval']
            hsn_map[key]['iamt'] += item.get('iamt', 0.0)
            hsn_map[key]['camt'] += item.get('camt', 0.0)
            hsn_map[key]['samt'] += item.get('samt', 0.0)
            hsn_map[key]['csamt'] += item.get('csamt', 0.0)

    merged_hsn = []
    hsn_counter = 1
    for key in sorted(hsn_map.keys(), key=lambda x: (x[1], x[0]), reverse=True):
        val = hsn_map[key]
        val['num'] = hsn_counter
        val['txval'] = round(val['txval'], 2)
        val['iamt'] = round(val['iamt'], 2)
        val['camt'] = round(val['camt'], 2)
        val['samt'] = round(val['samt'], 2)
        val['csamt'] = round(val['csamt'], 2)
        val['total_val'] = round(val['txval'] + val['iamt'] + val['camt'] + val['samt'], 2)
        merged_hsn.append(val)
        hsn_counter += 1

    # 3. Merge Documents Issued
    all_docs = []
    for res in marketplace_results:
        for doc in res.get('docs', []):
            all_docs.append(doc)

    # 4. Merge ECO (Table 14)
    all_eco = []
    for res in marketplace_results:
        eco = res.get('eco_summary')
        if eco and eco.get('suppval', 0) > 0:
            all_eco.append({
                'nature': eco.get('nature', 'Liable to collect tax u/s 52(TCS)'),
                'etin': eco.get('etin'),
                'name': eco.get('name'),
                'suppval': round(eco.get('suppval', 0.0), 2),
                'igst': round(eco.get('igst', 0.0), 2),
                'cgst': round(eco.get('cgst', 0.0), 2),
                'sgst': round(eco.get('sgst', 0.0), 2),
                'cess': 0.0,
                'flag': 'N'
            })

    # 5. Grand Totals for Summary Report & Dashboard
    total_taxable_b2cs = sum(i['txval'] for i in merged_b2cs)
    total_igst = sum(i['iamt'] for i in merged_b2cs)
    total_cgst = sum(i['camt'] for i in merged_b2cs)
    total_sgst = sum(i['samt'] for i in merged_b2cs)
    total_gst = total_igst + total_cgst + total_sgst
    total_nil_rated = round(total_nil_intra + total_nil_inter, 2)
    total_invoice_value = total_taxable_b2cs + total_gst + total_nil_rated

    summary_metrics = {
        'seller_gstin': seller_gstin,
        'filing_period': filing_period,
        'taxable_sales': round(total_taxable_b2cs, 2),
        'total_gst': round(total_gst, 2),
        'igst': round(total_igst, 2),
        'cgst': round(total_cgst, 2),
        'sgst': round(total_sgst, 2),
        'nil_rated': total_nil_rated,
        'nil_intra': round(total_nil_intra, 2),
        'nil_inter': round(total_nil_inter, 2),
        'total_invoice_value': round(total_invoice_value, 2),
        'b2cs_records_count': len(merged_b2cs),
        'hsn_records_count': len(merged_hsn),
        'total_docs_count': sum(d['totnum'] for d in all_docs),
        'eco_count': len(all_eco)
    }

    # 6. Build Official JSON Object
    json_b2cs = []
    for item in merged_b2cs:
        record = {
            'sply_ty': item['sply_ty'],
            'rt': item['rt'],
            'typ': item['typ'],
            'pos': str(item['pos']).zfill(2),
            'txval': item['txval'],
            'csamt': 0
        }
        if item['sply_ty'] == 'INTRA':
            record['camt'] = item['camt']
            record['samt'] = item['samt']
        else:
            record['iamt'] = item['iamt']
        json_b2cs.append(record)

    json_hsn_b2c = []
    for item in merged_hsn:
        json_hsn_b2c.append({
            'num': item['num'],
            'hsn_sc': item['hsn_sc'],
            'uqc': item['uqc'],
            'qty': item['qty'],
            'rt': item['rt'],
            'txval': item['txval'],
            'iamt': item['iamt'],
            'samt': item['samt'],
            'camt': item['camt'],
            'csamt': 0
        })

    json_supeco = {
        'clttx': [
            {
                'etin': eco['etin'],
                'suppval': eco['suppval'],
                'igst': eco['igst'],
                'cgst': eco['cgst'],
                'sgst': eco['sgst'],
                'cess': 0,
                'flag': 'N'
            }
            for eco in all_eco
        ]
    }

    # Format doc_issue structure
    doc_det = []
    inv_docs = [d for d in all_docs if d['doc_num'] == 1]
    cn_docs = [d for d in all_docs if d['doc_num'] == 5]

    if inv_docs:
        doc_det.append({
            'doc_num': 1,
            'doc_typ': 'Invoices for outward supply',
            'docs': [
                {
                    'num': i + 1,
                    'from': d['from'],
                    'to': d['to'],
                    'totnum': d['totnum'],
                    'cancel': d['cancel'],
                    'net_issue': d['net_issue']
                }
                for i, d in enumerate(inv_docs)
            ]
        })

    if cn_docs:
        doc_det.append({
            'doc_num': 5,
            'doc_typ': 'Credit Note',
            'docs': [
                {
                    'num': i + 1,
                    'from': d['from'],
                    'to': d['to'],
                    'totnum': d['totnum'],
                    'cancel': d['cancel'],
                    'net_issue': d['net_issue']
                }
                for i, d in enumerate(cn_docs)
            ]
        })

    gstr1_json_obj = {
        'gstin': seller_gstin,
        'fp': filing_period,
        'version': 'GST3.1.6',
        'hash': 'hash',
        'b2cs': json_b2cs,
        'hsn': {
            'hsn_b2c': json_hsn_b2c,
            'hsn_b2b': [
                {
                    'num': 1,
                    'hsn_sc': '490900',
                    'uqc': 'PCS',
                    'qty': 0,
                    'rt': 0,
                    'txval': 0,
                    'iamt': 0,
                    'samt': 0,
                    'camt': 0,
                    'csamt': 0
                }
            ]
        },
        'nil': {
            'inv': [
                {'sply_ty': 'INTRB2B', 'nil_amt': 0, 'expt_amt': 0, 'ngsup_amt': 0},
                {'sply_ty': 'INTRAB2B', 'nil_amt': 0, 'expt_amt': 0, 'ngsup_amt': 0},
                {'sply_ty': 'INTRB2C', 'nil_amt': int(round(total_nil_inter)), 'expt_amt': 0, 'ngsup_amt': 0},
                {'sply_ty': 'INTRAB2C', 'nil_amt': int(round(total_nil_intra)), 'expt_amt': 0, 'ngsup_amt': 0}
            ]
        },
        'supeco': json_supeco,
        'doc_issue': {'doc_det': doc_det}
    }

    if output_json_path:
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(gstr1_json_obj, f, indent=2)

    # 7. Build Official GSTR-1 Excel File
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Sheet definitions matching GSTR-1 offline utility
    _build_excel_b2b(wb)
    _build_excel_b2cl(wb)
    _build_excel_b2cs(wb, merged_b2cs)
    _build_excel_cdnr(wb)
    _build_excel_hsn_b2b(wb)
    _build_excel_hsn_b2c(wb, merged_hsn)
    _build_excel_exemp(wb, total_nil_inter, total_nil_intra)
    _build_excel_eco(wb, all_eco)
    _build_excel_docs(wb, all_docs)

    if output_excel_path:
        wb.save(output_excel_path)

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    return {
        'summary': summary_metrics,
        'json_data': gstr1_json_obj,
        'excel_bytes': excel_buffer.getvalue(),
        'b2cs': merged_b2cs,
        'hsn': merged_hsn,
        'docs': all_docs,
        'eco': all_eco
    }

# Excel Sheet Helper Builders
def _build_excel_b2b(wb):
    ws = wb.create_sheet(title='b2b,sez,de')
    ws.append(['Summary For B2B, SEZ, DE (4A, 4B, 6B, 6C)'])
    ws.append(['No. of Recipients', None, 'No. of Invoices', None, 'Total Invoice Value', None, None, None, None, None, None, 'Total Taxable Value', 'Total Cess'])
    ws.append([0, None, 0, None, 0, None, None, None, None, None, None, 0, 0])
    ws.append(['GSTIN/UIN of Recipient', 'Receiver Name', 'Invoice Number', 'Invoice date', 'Invoice Value', 'Place Of Supply', 'Reverse Charge', 'Applicable % of Tax Rate', 'Invoice Type', 'E-Commerce GSTIN', 'Rate', 'Taxable Value', 'Cess Amount'])

def _build_excel_b2cl(wb):
    ws = wb.create_sheet(title='b2cl')
    ws.append(['Summary For B2CL(5)', None, None, None, None, None, None, None, 'HELP'])
    ws.append(['No. of Invoices', None, None, None, None, None, None, None, None])
    ws.append([0, None, 0, None, None, None, 0, None, None])
    ws.append(['Invoice Number', 'Invoice date', 'Invoice Value', 'Place Of Supply', 'Applicable % of Tax Rate', 'Rate', 'Taxable Value', 'Cess Amount', 'E-Commerce GSTIN'])

def _build_excel_b2cs(wb, b2cs_list):
    ws = wb.create_sheet(title='b2cs')
    ws.append(['Summary For B2CS(7)', None, None, None, None, None, 'HELP'])
    ws.append([None, None, None, None, 'Total Taxable  Value', 'Total Cess', None])
    ws.append([None, None, None, None, 0, 0, None])
    ws.append(['Type', 'Place Of Supply', 'Applicable % of Tax Rate', 'Rate', 'Taxable Value', 'Cess Amount', 'E-Commerce GSTIN'])
    for item in b2cs_list:
        ws.append([
            item['typ'],
            item['pos_name'],
            None,
            item['rt'],
            item['txval'],
            item['csamt'],
            None
        ])

def _build_excel_cdnr(wb):
    ws = wb.create_sheet(title='cdnr')
    ws.append(['Summary For CDNR(9B)'])
    ws.append(['No. of Recipients', None, 'No. of Notes', None, None, None, None, None, 'Total Note Value', None, None, 'Total Taxable Value', 'Total Cess'])
    ws.append([0, None, 0, None, None, None, None, None, 0, None, None, 0, 0])
    ws.append(['GSTIN/UIN of Recipient', 'Receiver Name', 'Note Number', 'Note Date', 'Note Type', 'Place Of Supply', 'Reverse Charge', 'Note Supply Type', 'Note Value', 'Applicable % of Tax Rate', 'Rate', 'Taxable Value', 'Cess Amount'])

def _build_excel_hsn_b2b(wb):
    ws = wb.create_sheet(title='hsn(b2b)')
    ws.append(['Summary For HSN(12)', None, None, None, None, None, None, None, None, None, 'HELP'])
    ws.append(['No. of HSN', None, None, None, 'Total Value', None, 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess'])
    ws.append([0, None, None, None, 0, None, 0, 0, 0, 0, 0])
    ws.append(['HSN', 'Description', 'UQC', 'Total Quantity', 'Total Value', 'Rate', 'Taxable Value', 'Integrated Tax Amount', 'Central Tax Amount', 'State/UT Tax Amount', 'Cess Amount'])

def _build_excel_hsn_b2c(wb, hsn_list):
    ws = wb.create_sheet(title='hsn(b2c)')
    ws.append(['Summary For HSN(12)', None, None, None, None, None, None, None, None, None, 'HELP'])
    ws.append(['No. of HSN', None, None, None, 'Total Value', None, 'Total Taxable Value', 'Total Integrated Tax', 'Total Central Tax', 'Total State/UT Tax', 'Total Cess'])
    ws.append([0, None, None, None, 0, None, 0, 0, 0, 0, 0])
    ws.append(['HSN', 'Description', 'UQC', 'Total Quantity', 'Total Value', 'Rate', 'Taxable Value', 'Integrated Tax Amount', 'Central Tax Amount', 'State/UT Tax Amount', 'Cess Amount'])
    for item in hsn_list:
        ws.append([
            item['hsn_sc'],
            None,
            item['uqc_name'],
            item['qty'],
            item['total_val'],
            item['rt'],
            item['txval'],
            item['iamt'],
            item['camt'],
            item['samt'],
            item['csamt']
        ])

def _build_excel_exemp(wb, nil_inter, nil_intra):
    ws = wb.create_sheet(title='exemp')
    ws.append(['Summary For Nil rated, exempted and non GST outward supplies (8)', None, None, 'HELP'])
    ws.append([None, 'Total Nil Rated Supplies', 'Total Exempted Supplies', 'Total Non-GST Supplies'])
    ws.append([None, 0, 0, 0])
    ws.append(['Description', 'Nil Rated Supplies', 'Exempted(other than nil rated/non GST supply)', 'Non-GST Supplies'])
    ws.append(['Inter-State supplies to registered persons', 0, 0, 0])
    ws.append(['Intra-State supplies to registered persons', 0, 0, 0])
    ws.append(['Inter-State supplies to unregistered persons', int(round(nil_inter)), 0, 0])
    ws.append(['Intra-State supplies to unregistered persons', int(round(nil_intra)), 0, 0])

def _build_excel_eco(wb, eco_list):
    ws = wb.create_sheet(title='eco')
    ws.append(['Summary For Supplies through ECO-14', None, None, None, None, None, None, 'HELP'])
    ws.append([None, 'No. of E-Commerce Operator', None, 'Total Net Value of Supplies', 'Total Integrated Tax', 'Total Central Tax ', 'Total State/UT Tax ', 'Total Cess'])
    ws.append([None, 0, None, 0, 0, 0, 0, 0])
    ws.append(['Nature of Supply', 'GSTIN of E-Commerce Operator', 'E-Commerce Operator Name', 'Net value of supplies', 'Integrated tax', 'Central tax', 'State/UT tax', 'Cess'])
    for item in eco_list:
        ws.append([
            item['nature'],
            item['etin'],
            item['name'],
            item['suppval'],
            item['igst'],
            item['cgst'],
            item['sgst'],
            item['cess']
        ])

def _build_excel_docs(wb, docs_list):
    ws = wb.create_sheet(title='docs')
    ws.append(['Summary of documents issued during the tax period (13)', None, None, None, 'HELP'])
    ws.append([None, None, None, 'Total Number', 'Total Cancelled'])
    ws.append([None, None, None, 4, 2])
    ws.append(['Nature of Document', 'Sr. No. From', 'Sr. No. To', 'Total Number', 'Cancelled'])
    for item in docs_list:
        ws.append([
            item['doc_typ'],
            item['from'],
            item['to'],
            item['totnum'],
            item['cancel']
        ])
