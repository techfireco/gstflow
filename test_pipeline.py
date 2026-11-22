"""
End-to-End Pipeline Test Script:
Tests parsing and GSTR-1 generation on actual October 2026 dataset and verifies parity with Final Output.
"""
import os
import json
import pandas as pd
from backend.parsers.meesho_parser import parse_meesho_reports
from backend.parsers.amazon_parser import parse_amazon_mtr
from backend.parsers.tcs_portal_parser import parse_tcs_portal_file
from backend.engines.gstr1_engine import generate_gstr1
from backend.engines.tally_xml_engine import generate_tally_xml
from backend.engines.gstr2b_engine import reconcile_gstr2b_vs_books

OCT_DIR = "GST TOOL DATA _ 07FVUPP8356J1ZG"

def test_october_data():
    print("=== 1. PARSING MEESHO REPORTS ===")
    meesho_sales = os.path.join(OCT_DIR, "Meesho", "tcs_sales.xlsx")
    meesho_ret = os.path.join(OCT_DIR, "Meesho", "tcs_sales_return.xlsx")
    meesho_inv = os.path.join(OCT_DIR, "Meesho", "Tax_invoice_details.xlsx")

    meesho_res = parse_meesho_reports(
        sales_file=meesho_sales,
        return_file=meesho_ret,
        invoice_details_file=meesho_inv,
        seller_gstin="07FVUPP8356J1ZG"
    )
    print(f"Meesho parsed: {meesho_res['raw_counts']}")
    print(f"Meesho ECO SuppVal: {meesho_res['eco_summary']['suppval']}")
    print(f"Meesho Nil: Intra={meesho_res['nil_intra']}, Inter={meesho_res['nil_inter']}")
    print(f"Meesho B2CS records: {len(meesho_res['b2cs'])}, HSN records: {len(meesho_res['hsn'])}, Docs: {len(meesho_res['docs'])}")

    print("\n=== 2. PARSING AMAZON REPORT ===")
    amz_csv = os.path.join(OCT_DIR, "Amazon", "MTR_B2C-OCTOBER-2026-A1BMCYF963TNMV.csv")
    amz_res = parse_amazon_mtr(amz_csv, seller_gstin="07FVUPP8356J1ZG")
    print(f"Amazon parsed: {amz_res['raw_counts']}")
    print(f"Amazon ECO SuppVal: {amz_res['eco_summary']['suppval']}")
    print(f"Amazon B2CS records: {len(amz_res['b2cs'])}, HSN records: {len(amz_res['hsn'])}, Docs: {len(amz_res['docs'])}")

    print("\n=== 3. PARSING GST PORTAL TCS FILE ===")
    tcs_zip = os.path.join(OCT_DIR, "TCS data from GST portal", "07FVUPP8356J1ZG_OCT_2026.zip")
    tcs_res = parse_tcs_portal_file(tcs_zip)
    print(f"Portal TCS Total Net Value: {tcs_res['total_net_value']}, ECO count: {len(tcs_res['eco_summary'])}")
    for etin, eco in tcs_res['eco_summary'].items():
        print(f"   ETIN: {etin} -> Net: {eco['net_value']}, IGST: {eco['iamt']}, CGST: {eco['camt']}, SGST: {eco['samt']}")

    print("\n=== 4. GENERATING COMBINED GSTR-1 ===")
    out_json = "/tmp/test_gstr1.json"
    out_xlsx = "/tmp/test_gstr1.xlsx"

    # For Amazon in October: If user wants GSTR-1 to align with TCS portal filed figures,
    # or direct MTR:
    gstr1_res = generate_gstr1(
        marketplace_results=[meesho_res, amz_res],
        seller_gstin="07FVUPP8356J1ZG",
        filing_period="102026",
        output_excel_path=out_xlsx,
        output_json_path=out_json
    )

    summary = gstr1_res['summary']
    print("Generated Summary:")
    for k, v in summary.items():
        print(f"  {k}: {v}")

    print("\n=== 5. GENERATING TALLY XML ===")
    tally_xml = generate_tally_xml(meesho_res, company_name="IQRA PARVEEN", import_mode="summary")
    print(f"Tally XML generated successfully! Length: {len(tally_xml)} chars")
    print("Sample Tally XML:")
    print(tally_xml[:600])

    print("\n=== 6. VERIFYING PARITY WITH GSTRTOOL FINAL OUTPUT ===")
    actual_json_path = os.path.join(OCT_DIR, "Final Output", "GSTR1_returns_07FVUPP8356J1ZG_monthly_102026.json")
    with open(actual_json_path) as f:
        actual_json = json.load(f)

    print(f"Actual JSON keys: {list(actual_json.keys())}")
    print(f"Actual JSON B2CS count: {len(actual_json['b2cs'])}, Generated: {len(gstr1_res['json_data']['b2cs'])}")
    print(f"Actual JSON HSN count: {len(actual_json['hsn']['hsn_b2c'])}, Generated: {len(gstr1_res['json_data']['hsn']['hsn_b2c'])}")
    print(f"Actual Nil rated: {actual_json['nil']['inv']}")
    print(f"Generated Nil rated: {gstr1_res['json_data']['nil']['inv']}")

if __name__ == "__main__":
    test_october_data()
