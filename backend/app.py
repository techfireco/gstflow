"""
FastAPI Backend Application for E-Commerce GST & Tally Workflow Suite.
"""
import os
import io
import json
import uuid
import tempfile
import pandas as pd
from typing import Optional, List
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from backend.parsers.meesho_parser import parse_meesho_reports
from backend.parsers.amazon_parser import parse_amazon_mtr
from backend.parsers.flipkart_parser import parse_flipkart_report
from backend.parsers.tcs_portal_parser import parse_tcs_portal_file
from backend.engines.gstr1_engine import generate_gstr1
from backend.engines.tally_xml_engine import generate_tally_xml
from backend.engines.gstr2b_engine import reconcile_gstr2b_vs_books
from backend.config import GST_STATE_CODES

app = FastAPI(title="E-Com GST & Tally Workflow Suite", version="1.0.0")

# Mount static files
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")
STATIC_DIR = os.path.join(FRONTEND_DIR, "static")

if not os.path.exists(STATIC_DIR):
    os.makedirs(STATIC_DIR, exist_ok=True)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Temporary cache storage for generated downloads
GENERATED_CACHE = {}

# HTML Page Routes
@app.get("/", response_class=HTMLResponse)
async def serve_home():
    with open(os.path.join(FRONTEND_DIR, "index.html"), "r", encoding="utf-8") as f:
        return f.read()

@app.get("/gstr1", response_class=HTMLResponse)
async def serve_gstr1():
    with open(os.path.join(FRONTEND_DIR, "gstr1.html"), "r", encoding="utf-8") as f:
        return f.read()

@app.get("/tally", response_class=HTMLResponse)
async def serve_tally():
    with open(os.path.join(FRONTEND_DIR, "tally.html"), "r", encoding="utf-8") as f:
        return f.read()

@app.get("/reconcile", response_class=HTMLResponse)
async def serve_reconcile():
    with open(os.path.join(FRONTEND_DIR, "reconcile.html"), "r", encoding="utf-8") as f:
        return f.read()

@app.get("/pricing", response_class=HTMLResponse)
async def serve_pricing():
    with open(os.path.join(FRONTEND_DIR, "pricing.html"), "r", encoding="utf-8") as f:
        return f.read()

# API Endpoints: GSTR-1 Converter
@app.post("/api/convert-gstr1")
async def api_convert_gstr1(
    seller_gstin: str = Form("07FVUPP8356J1ZG"),
    filing_period: str = Form("102026"),
    align_with_tcs: bool = Form(False),
    meesho_sales: Optional[UploadFile] = File(None),
    meesho_returns: Optional[UploadFile] = File(None),
    meesho_invoices: Optional[UploadFile] = File(None),
    amazon_mtr: Optional[UploadFile] = File(None),
    flipkart_sales: Optional[UploadFile] = File(None),
    portal_tcs: Optional[UploadFile] = File(None)
):
    try:
        seller_gstin = seller_gstin.strip().upper()
        filing_period = filing_period.strip()
        marketplace_results = []

        # 1. Process Meesho
        if meesho_sales and meesho_returns:
            sales_bytes = io.BytesIO(await meesho_sales.read())
            ret_bytes = io.BytesIO(await meesho_returns.read())
            inv_bytes = io.BytesIO(await meesho_invoices.read()) if meesho_invoices else None

            meesho_data = parse_meesho_reports(
                sales_file=sales_bytes,
                return_file=ret_bytes,
                invoice_details_file=inv_bytes,
                seller_gstin=seller_gstin
            )
            marketplace_results.append(meesho_data)

        # 2. Process Amazon
        if amazon_mtr:
            amz_bytes = io.BytesIO(await amazon_mtr.read())
            amz_data = parse_amazon_mtr(amz_bytes, seller_gstin=seller_gstin)
            marketplace_results.append(amz_data)

        # 3. Process Flipkart
        if flipkart_sales:
            fk_bytes = io.BytesIO(await flipkart_sales.read())
            fk_data = parse_flipkart_report(fk_bytes, seller_gstin=seller_gstin)
            marketplace_results.append(fk_data)

        # 4. Optional: Process GST Portal TCS File for Reconciliation
        tcs_portal_data = None
        if portal_tcs:
            tcs_bytes = io.BytesIO(await portal_tcs.read())
            tcs_portal_data = parse_tcs_portal_file(tcs_bytes)

        if not marketplace_results and not tcs_portal_data:
            raise HTTPException(status_code=400, detail="Please upload at least one marketplace report (Meesho, Amazon, or Flipkart).")

        # Generate GSTR-1 Output
        session_id = str(uuid.uuid4())
        excel_tmp = os.path.join(tempfile.gettempdir(), f"GSTR1_{seller_gstin}_monthly_{filing_period}_{session_id}.xlsx")
        json_tmp = os.path.join(tempfile.gettempdir(), f"GSTR1_returns_{seller_gstin}_monthly_{filing_period}_{session_id}.json")

        result = generate_gstr1(
            marketplace_results=marketplace_results,
            seller_gstin=seller_gstin,
            filing_period=filing_period,
            output_excel_path=excel_tmp,
            output_json_path=json_tmp
        )

        # Cache session paths
        GENERATED_CACHE[session_id] = {
            'excel_path': excel_tmp,
            'json_path': json_tmp,
            'seller_gstin': seller_gstin,
            'filing_period': filing_period
        }

        response_data = {
            'session_id': session_id,
            'summary': result['summary'],
            'b2cs': result['b2cs'],
            'hsn': result['hsn'],
            'docs': result['docs'],
            'eco': result['eco'],
            'marketplaces_processed': [r['marketplace'] for r in marketplace_results],
            'tcs_portal_summary': tcs_portal_data.get('eco_summary') if tcs_portal_data else None
        }

        return JSONResponse(response_data)

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

# One-Click Demo Loader: loads October 2026 files directly from workspace
@app.post("/api/load-demo-october")
async def api_load_demo_october():
    try:
        oct_dir = os.path.join(BASE_DIR, "GST TOOL DATA _ 07FVUPP8356J1ZG")
        meesho_sales = os.path.join(oct_dir, "Meesho", "tcs_sales.xlsx")
        meesho_ret = os.path.join(oct_dir, "Meesho", "tcs_sales_return.xlsx")
        meesho_inv = os.path.join(oct_dir, "Meesho", "Tax_invoice_details.xlsx")
        amz_csv = os.path.join(oct_dir, "Amazon", "MTR_B2C-OCTOBER-2026-A1BMCYF963TNMV.csv")
        tcs_zip = os.path.join(oct_dir, "TCS data from GST portal", "07FVUPP8356J1ZG_OCT_2026.zip")

        seller_gstin = "07FVUPP8356J1ZG"
        filing_period = "102026"

        meesho_data = parse_meesho_reports(
            sales_file=meesho_sales,
            return_file=meesho_ret,
            invoice_details_file=meesho_inv,
            seller_gstin=seller_gstin
        )
        amz_data = parse_amazon_mtr(amz_csv, seller_gstin=seller_gstin)
        tcs_portal_data = parse_tcs_portal_file(tcs_zip)

        session_id = str(uuid.uuid4())
        excel_tmp = os.path.join(tempfile.gettempdir(), f"GSTR1_{seller_gstin}_monthly_{filing_period}_{session_id}.xlsx")
        json_tmp = os.path.join(tempfile.gettempdir(), f"GSTR1_returns_{seller_gstin}_monthly_{filing_period}_{session_id}.json")

        result = generate_gstr1(
            marketplace_results=[meesho_data, amz_data],
            seller_gstin=seller_gstin,
            filing_period=filing_period,
            output_excel_path=excel_tmp,
            output_json_path=json_tmp
        )

        GENERATED_CACHE[session_id] = {
            'excel_path': excel_tmp,
            'json_path': json_tmp,
            'seller_gstin': seller_gstin,
            'filing_period': filing_period
        }

        return JSONResponse({
            'session_id': session_id,
            'summary': result['summary'],
            'b2cs': result['b2cs'],
            'hsn': result['hsn'],
            'docs': result['docs'],
            'eco': result['eco'],
            'marketplaces_processed': ['Meesho', 'Amazon'],
            'tcs_portal_summary': tcs_portal_data.get('eco_summary')
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))

# Download Endpoints
@app.get("/api/download/gstr1-json")
async def download_gstr1_json(session_id: str):
    if session_id not in GENERATED_CACHE:
        raise HTTPException(status_code=404, detail="File session expired or not found.")
    info = GENERATED_CACHE[session_id]
    filename = f"GSTR1_returns_{info['seller_gstin']}_monthly_{info['filing_period']}.json"
    return FileResponse(
        info['json_path'],
        media_type="application/json",
        filename=filename
    )

@app.get("/api/download/gstr1-excel")
async def download_gstr1_excel(session_id: str):
    if session_id not in GENERATED_CACHE:
        raise HTTPException(status_code=404, detail="File session expired or not found.")
    info = GENERATED_CACHE[session_id]
    filename = f"GSTR1_{info['seller_gstin']}_monthly_{info['filing_period']}.xlsx"
    return FileResponse(
        info['excel_path'],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )

# Tally XML API
@app.post("/api/convert-tally-xml")
async def api_convert_tally_xml(
    marketplace: str = Form("Meesho"),
    company_name: str = Form("My Company"),
    import_mode: str = Form("summary"),
    party_ledger: str = Form(""),
    sales_file: Optional[UploadFile] = File(None),
    return_file: Optional[UploadFile] = File(None),
    amazon_file: Optional[UploadFile] = File(None)
):
    try:
        mp_data = {}
        if marketplace == "Meesho" and sales_file and return_file:
            s_bytes = io.BytesIO(await sales_file.read())
            r_bytes = io.BytesIO(await return_file.read())
            mp_data = parse_meesho_reports(s_bytes, r_bytes)
        elif marketplace == "Amazon" and amazon_file:
            a_bytes = io.BytesIO(await amazon_file.read())
            mp_data = parse_amazon_mtr(a_bytes)
        else:
            # Fallback to demo meesho
            oct_dir = os.path.join(BASE_DIR, "GST TOOL DATA _ 07FVUPP8356J1ZG")
            s_file = os.path.join(oct_dir, "Meesho", "tcs_sales.xlsx")
            r_file = os.path.join(oct_dir, "Meesho", "tcs_sales_return.xlsx")
            mp_data = parse_meesho_reports(s_file, r_file)

        xml_str = generate_tally_xml(
            marketplace_data=mp_data,
            company_name=company_name,
            import_mode=import_mode,
            party_ledger=party_ledger or f"{marketplace} Sales"
        )

        return StreamingResponse(
            io.BytesIO(xml_str.encode('utf-8')),
            media_type="application/xml",
            headers={"Content-Disposition": f"attachment; filename={marketplace}_Tally_Import.xml"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# GSTR-2B Reconciliation API (Free Funnel)
@app.post("/api/reconcile-gstr2b")
async def api_reconcile_gstr2b(
    purchase_register: UploadFile = File(...),
    gstr2b_file: UploadFile = File(...),
    tolerance: float = Form(1.0)
):
    try:
        pr_bytes = io.BytesIO(await purchase_register.read())
        pr_name = purchase_register.filename or "purchase.xlsx"

        g2b_bytes = io.BytesIO(await gstr2b_file.read())
        g2b_name = gstr2b_file.filename or "gstr2b.json"

        # Save temporarily
        pr_temp = os.path.join(tempfile.gettempdir(), f"pr_{uuid.uuid4()}_{pr_name}")
        g2b_temp = os.path.join(tempfile.gettempdir(), f"2b_{uuid.uuid4()}_{g2b_name}")

        with open(pr_temp, "wb") as f:
            f.write(pr_bytes.getvalue())
        with open(g2b_temp, "wb") as f:
            f.write(g2b_bytes.getvalue())

        res = reconcile_gstr2b_vs_books(pr_temp, g2b_temp, tolerance=tolerance)
        return JSONResponse(res)

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))
