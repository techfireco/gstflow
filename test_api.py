"""
API Integration Test Script:
Verifies FastAPI endpoints, demo loader, and file downloads.
"""
from fastapi.testclient import TestClient
from backend.app import app
import json

client = TestClient(app)

def test_routes():
    print("1. Testing GET / ...")
    r = client.get("/")
    assert r.status_code == 200
    assert "GSTFlow" in r.text
    print("   -> OK!")

    print("2. Testing GET /gstr1 ...")
    r = client.get("/gstr1")
    assert r.status_code == 200
    assert "GSTR-1" in r.text
    print("   -> OK!")

    print("3. Testing POST /api/load-demo-october ...")
    r = client.post("/api/load-demo-october")
    assert r.status_code == 200, r.text
    data = r.json()
    assert "summary" in data
    assert "session_id" in data
    session_id = data["session_id"]
    summary = data["summary"]
    print(f"   -> OK! Taxable Sales: {summary['taxable_sales']}, Total GST: {summary['total_gst']}")
    print(f"   -> B2CS count: {summary['b2cs_records_count']}, HSN count: {summary['hsn_records_count']}")

    print("4. Testing GET /api/download/gstr1-json ...")
    r = client.get(f"/api/download/gstr1-json?session_id={session_id}")
    assert r.status_code == 200
    json_content = r.json()
    assert json_content["gstin"] == "07FVUPP8356J1ZG"
    assert "b2cs" in json_content
    assert "hsn" in json_content
    print("   -> OK! Downloaded valid GSTR-1 JSON.")

    print("5. Testing GET /api/download/gstr1-excel ...")
    r = client.get(f"/api/download/gstr1-excel?session_id={session_id}")
    assert r.status_code == 200
    assert len(r.content) > 5000
    print("   -> OK! Downloaded valid GSTR-1 Excel.")

    print("6. Testing POST /api/convert-tally-xml ...")
    r = client.post("/api/convert-tally-xml", data={"marketplace": "Meesho", "company_name": "IQRA PARVEEN", "import_mode": "summary"})
    assert r.status_code == 200
    assert "<ENVELOPE>" in r.text
    assert "TallyUDF" in r.text
    print("   -> OK! Generated valid Tally XML.")

    print("\n🎉 ALL API TESTS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    test_routes()
