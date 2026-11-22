"""
Server Launcher Script for E-Commerce GST & Tally Workflow Suite.
"""
import uvicorn

if __name__ == "__main__":
    print("================================================================")
    print("🚀 Starting GSTFlow E-Commerce GST & Tally Workflow Server...")
    print("👉 Dashboard:           http://127.0.0.1:8000")
    print("👉 GSTR-1 Generator:    http://127.0.0.1:8000/gstr1")
    print("👉 Tally XML Importer:  http://127.0.0.1:8000/tally")
    print("👉 Free 2B Reconciler:  http://127.0.0.1:8000/reconcile")
    print("👉 SaaS Pricing Plans:  http://127.0.0.1:8000/pricing")
    print("================================================================")
    uvicorn.run("backend.app:app", host="127.0.0.1", port=8000, reload=True)
