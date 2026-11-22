"""
Marketplace to Tally XML Export Engine:
Generates official Tally Prime / Tally.ERP 9 XML import files for Sales, Credit Notes, and TCS.
"""
import xml.etree.ElementTree as ET
from xml.dom import minidom
from typing import Dict, Any, List, Optional
import datetime

def generate_tally_xml(
    marketplace_data: Dict[str, Any],
    company_name: str = "My E-Commerce Company",
    import_mode: str = "summary", # 'summary' or 'detailed'
    party_ledger: Optional[str] = None,
    sales_ledger_prefix: str = "E-Commerce Sales",
    igst_ledger: str = "Output IGST",
    cgst_ledger: str = "Output CGST",
    sgst_ledger: str = "Output SGST",
    tcs_ledger: str = "TCS Receivable u/s 52"
) -> str:
    """
    Generates Tally XML ready for Import -> Transactions in Tally Prime / Tally.ERP 9.
    """
    marketplace = marketplace_data.get('marketplace', 'E-Commerce')
    party_name = party_ledger or f"{marketplace} Sales Ledger"
    
    # Root Envelope
    envelope = ET.Element("ENVELOPE")
    
    # Header
    header = ET.SubElement(envelope, "HEADER")
    tally_req = ET.SubElement(header, "TALLYREQUEST")
    tally_req.text = "Import Data"

    # Body
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    
    req_desc = ET.SubElement(import_data, "REQUESTDESC")
    rep_name = ET.SubElement(req_desc, "REPORTNAME")
    rep_name.text = "Vouchers"
    
    stat_vars = ET.SubElement(req_desc, "STATICVARIABLES")
    comp_elem = ET.SubElement(stat_vars, "SVCURRENTCOMPANY")
    comp_elem.text = company_name

    req_data = ET.SubElement(import_data, "REQUESTDATA")

    # In summary mode: generate state and rate wise consolidated vouchers
    if import_mode == "summary":
        b2cs = marketplace_data.get('b2cs', [])
        for item in b2cs:
            txval = float(item.get('txval', 0.0))
            if abs(txval) <= 0.01:
                continue

            rate = item.get('rt', 0)
            is_intra = (item.get('sply_ty') == 'INTRA')
            pos_name = item.get('pos_name', 'Domestic')
            vch_type = "Sales" if txval > 0 else "Credit Note"
            
            tally_msg = ET.SubElement(req_data, "TALLYMESSAGE", {"xmlns:UDF": "TallyUDF"})
            vch = ET.SubElement(tally_msg, "VOUCHER", {
                "VCHTYPE": vch_type,
                "ACTION": "Create",
                "OBJVIEW": "Accounting Voucher View"
            })

            # Date format YYYYMMDD
            now_str = datetime.datetime.now().strftime("%Y%m01")
            vch_date = ET.SubElement(vch, "DATE")
            vch_date.text = now_str
            
            vch_type_elem = ET.SubElement(vch, "VOUCHERTYPENAME")
            vch_type_elem.text = vch_type

            party_elem = ET.SubElement(vch, "PARTYLEDGERNAME")
            party_elem.text = party_name

            vch_narration = ET.SubElement(vch, "NARRATION")
            vch_narration.text = f"{marketplace} {vch_type} for POS: {pos_name} @ {rate}% GST"

            # Ledgers
            # 1. Party Ledger (Debit for Sales, Credit for Credit Note)
            total_tax = float(item.get('iamt', 0.0)) + float(item.get('camt', 0.0)) + float(item.get('samt', 0.0))
            total_inv_val = round(abs(txval) + total_tax, 2)
            
            party_entry = ET.SubElement(vch, "ALLLEDGERENTRIES.LIST")
            p_led = ET.SubElement(party_entry, "LEDGERNAME")
            p_led.text = party_name
            p_is_debit = ET.SubElement(party_entry, "ISDEEMEDPOSITIVE")
            p_is_debit.text = "Yes" if vch_type == "Sales" else "No"
            p_amt = ET.SubElement(party_entry, "AMOUNT")
            p_amt.text = f"{-total_inv_val if vch_type == 'Sales' else total_inv_val}"

            # 2. Sales Ledger (Credit for Sales)
            sales_entry = ET.SubElement(vch, "ALLLEDGERENTRIES.LIST")
            s_led = ET.SubElement(sales_entry, "LEDGERNAME")
            s_led.text = f"{sales_ledger_prefix} {rate}%"
            s_is_debit = ET.SubElement(sales_entry, "ISDEEMEDPOSITIVE")
            s_is_debit.text = "No" if vch_type == "Sales" else "Yes"
            s_amt = ET.SubElement(sales_entry, "AMOUNT")
            s_amt.text = f"{abs(txval) if vch_type == 'Sales' else -abs(txval)}"

            # 3. Tax Ledgers
            if is_intra:
                for tax_name, tax_val in [ (cgst_ledger, item.get('camt', 0.0)), (sgst_ledger, item.get('samt', 0.0)) ]:
                    if tax_val > 0:
                        tax_entry = ET.SubElement(vch, "ALLLEDGERENTRIES.LIST")
                        t_led = ET.SubElement(tax_entry, "LEDGERNAME")
                        t_led.text = tax_name
                        t_is_debit = ET.SubElement(tax_entry, "ISDEEMEDPOSITIVE")
                        t_is_debit.text = "No" if vch_type == "Sales" else "Yes"
                        t_amt = ET.SubElement(tax_entry, "AMOUNT")
                        t_amt.text = f"{tax_val if vch_type == 'Sales' else -tax_val}"
            else:
                igst_val = float(item.get('iamt', 0.0))
                if igst_val > 0:
                    tax_entry = ET.SubElement(vch, "ALLLEDGERENTRIES.LIST")
                    t_led = ET.SubElement(tax_entry, "LEDGERNAME")
                    t_led.text = igst_ledger
                    t_is_debit = ET.SubElement(tax_entry, "ISDEEMEDPOSITIVE")
                    t_is_debit.text = "No" if vch_type == "Sales" else "Yes"
                    t_amt = ET.SubElement(tax_entry, "AMOUNT")
                    t_amt.text = f"{igst_val if vch_type == 'Sales' else -igst_val}"

    # Return formatted XML string
    rough_string = ET.tostring(envelope, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")
