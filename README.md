# GSTFlow — E-Commerce GST & Tally Workflow Suite

An automated compliance and accounting platform for Indian e-commerce sellers (Meesho, Amazon, Flipkart). Converts raw marketplace settlement and tax reports directly into official **GST Portal-ready GSTR-1 Excel & JSON**, generates **Tally Prime XML vouchers**, and provides **GSTR-2B vs Books reconciliation** as a high-converting free acquisition funnel.

Inspired by [gsttool.in](https://gsttool.in/).

---

## 🚀 Key Features & Product Architecture

```
                             ┌──────────────────────────────────────────────┐
                             │           E-COMMERCE RAW REPORTS             │
                             │   Meesho (Sales, Returns, Invoice Details)   │
                             │   Amazon (MTR B2C CSV)                       │
                             │   Flipkart (Sales & Section 52 Tax CSV)      │
                             └──────────────────────┬───────────────────────┘
                                                    │
                      ┌─────────────────────────────┼─────────────────────────────┐
                      ▼                             ▼                             ▼
       ┌────────────────────────────┐┌────────────────────────────┐┌────────────────────────────┐
       │     TOOL 1 (₹149/mo)       ││     TOOL 2 (₹249/mo)       ││    FREE FUNNEL (₹0)        │
       │   GSTR-1 Auto-Generator    ││  Marketplace to Tally XML  ││   GSTR-2B vs Books Match   │
       ├────────────────────────────┤├────────────────────────────┤├────────────────────────────┤
       │• Table 7 B2CS State POS    ││• Tally Prime / ERP 9 XML   ││• Rule 36(4) Risk Detection │
       │• Table 8 Nil/Exempted      ││• Consolidated state vouchers││• Supplier GSTR-1 Defaulters│
       │• Table 12 HSN (PCS, rate)  ││• Credit Notes for returns  ││• Fuzzy invoice normalization│
       │• Table 13 Document Series  ││• TCS u/s 52 Asset Ledger   ││• Multi-sheet Excel Audit   │
       │• Table 14 ECO u/s 52       ││• Zero database bloating    ││• Converts users to Tool 1/2│
       │• Official JSON (GST3.1.6)  ││• Instant voucher import    ││                            │
       │• Official 9-Sheet Excel    ││                            ││                            │
       └────────────────────────────┘└────────────────────────────┘└────────────────────────────┘
```

---

## 📊 How Raw Data Transforms to GSTR-1 Output (The Exact Math)

Our reverse-engineering of the provided October 2026 dataset (`07FVUPP8356J1ZG`) reveals the exact data engineering pipeline:

### 1. Meesho Raw Reports
* **Inputs:** `tcs_sales.xlsx`, `tcs_sales_return.xlsx`, `Tax_invoice_details.xlsx`
* **Net Taxable Sales:** Sum of `total_taxable_sale_value` in sales minus returns:
  $$\text{Meesho Net Taxable} = ₹3,27,888.57 - ₹80,852.18 = ₹2,47,036.39$$
  *(Matches the ₹2,47,036.33 in Table 14 ECO Section 52).*
* **Table 7 (B2CS):** Grouped by `(end_customer_state_new, gst_rate)`.
  * Deliveries to Home State (Delhi `07`) $\rightarrow$ Intra-state (`CGST` + `SGST`).
  * Deliveries to other States $\rightarrow$ Inter-state (`IGST`).
* **Table 8 (Nil/Exempted Supplies):**
  * Meesho orders with `gst_rate = 0.0` are isolated:
  * Intra-state (Delhi): ₹168.00 $\rightarrow$ `Intra-State supplies to unregistered persons`
  * Inter-state (Non-Delhi): ₹263.00 $\rightarrow$ `Inter-State supplies to unregistered persons`
  * Total Nil-Rated = ₹431.00.
* **Table 13 (Document Series):**
  * Invoices: `rw7a3274387` to `rw7a3276488` (Total: 2,063)
  * Credit Notes: `rw7a327C1116` to `rw7a327C1736` (Total: 557)
  * Credit Discounts: `rw7a327CM96` to `rw7a327CM118` (Total: 23)

### 2. Amazon Merchant Tax Report (MTR B2C CSV)
* **Inputs:** `MTR_B2C-OCTOBER-2026-A1BMCYF963TNMV.csv`
* **Filtering:** Drops `Cancel` and `FreeReplacement` rows (0 value).
* **Net Sales:** Shipments ($+₹91,547.53$) minus Refunds ($-₹1,598.05$) = $₹89,949.48$.
* **Table 13 (Document Series):**
  * Invoices: `IN-595` to `IN-768` (Total: 150)
  * Credit Notes: `CN-67` to `CN-71` (Total: 5)
* **HSN Summary (Table 12):**
  * HSN 4802 (3%): Qty $112 - 3 = 109$, Taxable: $₹38,821.40 - ₹773.78 = ₹38,047.62$
  * HSN 57050090 (3%): Qty $25 - 1 = 24$, Taxable: $₹16,795.17 - ₹631.07 = ₹16,164.10$
  * HSN 4820 (3%): Qty 120, Taxable: $₹35,562.04$
  * HSN 49090090 (3%): Qty 4, Taxable: $₹1,165.04$
  * HSN 4820 (0% Nil): Odisha order = $₹4,299.00$

### 3. Combined GSTR-1 Grand Totals
* **Taxable Sales:** ₹3,41,843.13
* **Total GST:** ₹15,058.60 (IGST: ₹14,257.28, CGST: ₹400.66, SGST: ₹400.66)
* **Exempt / Nil Sales:** ₹431.00
* **Total Invoiced Value:** ₹3,57,332.70
* **Documents Issued:** 2,798 total documents

---

## ⚡ Quick Start

### 1. Requirements
* Python 3.10+
* `uv` or standard Python `pip`

### 2. Run the Server
Using `uv`:
```bash
uv run --with fastapi,uvicorn,python-multipart,pandas,openpyxl python run_server.py
```

Or using standard `pip`:
```bash
pip install fastapi uvicorn python-multipart pandas openpyxl
python run_server.py
```

### 3. Access the Web Dashboard
* **Dashboard:** [http://127.0.0.1:8000](http://127.0.0.1:8000)
* **GSTR-1 Auto-Filer:** [http://127.0.0.1:8000/gstr1](http://127.0.0.1:8000/gstr1)
* **Tally XML Sync:** [http://127.0.0.1:8000/tally](http://127.0.0.1:8000/tally)
* **Free GSTR-2B Reconciler:** [http://127.0.0.1:8000/reconcile](http://127.0.0.1:8000/reconcile)
* **Pricing Plans:** [http://127.0.0.1:8000/pricing](http://127.0.0.1:8000/pricing)

> **Pro Tip:** On the GSTR-1 page, click **"⚡ Load October Demo Data (1-Click)"** to immediately see the engine transform the raw Meesho & Amazon files into the complete GSTR-1 dashboard and download the official JSON and Excel files.

---

## 💼 SaaS Monetization & Funnel Strategy

| Tier | Price | Features Included |
| :--- | :--- | :--- |
| **Free Funnel** | **₹0 / Free** | **GSTR-2B vs Books Reconciler**: Unlimited invoice matching, Rule 36(4) blocked ITC risk identification, supplier defaulter alerts, and multi-sheet Excel export. *(High-volume organic traffic magnet)* |
| **Starter** | **₹149 / mo** | **GSTR-1 Auto-Filer**: Unlimited orders, Meesho + Amazon + Flipkart support, Table 7 B2CS, Table 8 Nil, Table 12 HSN, Table 13 Docs, Table 14 ECO u/s 52, official GST3.1.6 JSON & 9-sheet Excel. |
| **Tally Sync** | **₹249 / mo** | **Marketplace to Tally XML**: 1-click import into Tally Prime & ERP 9, consolidated state POS vouchers, automated return credit notes, and TCS deduction ledger mapping. |
| **Pro Suite** | **₹399 / mo** | **All-in-One**: GSTR-1 + Tally XML + GST Portal GSTR-8 TCS cross-reconciler (protects against Rule 88C/88D mismatch notices) + Multi-channel Profitability & P&L audit. |

---

## 🛠️ Project Structure

```
GST/
├── backend/
│   ├── app.py                     # FastAPI REST API & Web Server
│   ├── config.py                  # 2-Digit GST State Map & ECO GSTINs
│   ├── parsers/
│   │   ├── meesho_parser.py       # Meesho tcs_sales & returns parser
│   │   ├── amazon_parser.py       # Amazon MTR B2C CSV parser
│   │   ├── flipkart_parser.py     # Flipkart sales/tax report parser
│   │   └── tcs_portal_parser.py   # GST Portal TCS (GSTR-8) parser
│   └── engines/
│       ├── gstr1_engine.py        # Official GSTR-1 JSON (GST3.1.6) & Excel builder
│       ├── gstr2b_engine.py       # GSTR-2B vs Purchase Register matching engine
│       └── tally_xml_engine.py    # Tally Prime / ERP 9 XML generator
├── frontend/
│   ├── index.html                 # Homepage & Product Overview
│   ├── gstr1.html                 # GSTR-1 Upload & Visual Summary Report
│   ├── tally.html                 # Tally XML Generator & Guide
│   ├── reconcile.html             # Free 2B Reconciler Funnel
│   └── pricing.html               # SaaS Pricing & Packaging
├── run_server.py                  # One-click server launcher
├── test_pipeline.py               # Data pipeline verification against actual files
└── test_api.py                    # API integration test suite
```
