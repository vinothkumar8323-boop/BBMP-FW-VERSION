import requests
import json

BASE_URL = "https://schnelliot.in"
USERNAME = "vinoth.joel@schnellenergy.com"
PASSWORD = "vinoth555"

session = requests.Session()
session.headers.update({"Content-Type": "application/json"})

# Login
login_url = f"{BASE_URL}/api/auth/login"
res = session.post(login_url, json={"username": USERNAME, "password": PASSWORD})
token = res.json()["token"]
session.headers.update({"X-Authorization": f"Bearer {token}"})

# 1. Fetch all DEVICE entity groups
eg_res = session.get(f"{BASE_URL}/api/entityGroups/DEVICE")
if eg_res.status_code == 200:
    groups = eg_res.json()
    print(f"--- DEVICE Entity Groups ({len(groups)}) ---")
    for g in groups:
        g_id = g["id"]["id"]
        g_name = g.get("name")
        print(f"Group ID: {g_id} | Name: {g_name}")

# 2. Check ASSET Entity Groups
eg_asset_res = session.get(f"{BASE_URL}/api/entityGroups/ASSET")
if eg_asset_res.status_code == 200:
    asset_groups = eg_asset_res.json()
    print(f"\n--- ASSET Entity Groups ({len(asset_groups)}) ---")
    for g in asset_groups:
        g_id = g["id"]["id"]
        g_name = g.get("name")
        print(f"Asset Group ID: {g_id} | Name: {g_name}")

# 3. Check CUSTOMER Entity Groups
eg_cust_res = session.get(f"{BASE_URL}/api/entityGroups/CUSTOMER")
if eg_cust_res.status_code == 200:
    cust_groups = eg_cust_res.json()
    print(f"\n--- CUSTOMER Entity Groups ({len(cust_groups)}) ---")
    for g in cust_groups:
        g_id = g["id"]["id"]
        g_name = g.get("name")
        print(f"Customer Group ID: {g_id} | Name: {g_name}")

# 4. Search Customers
cust_res = session.get(f"{BASE_URL}/api/customers?pageSize=100&page=0")
if cust_res.status_code == 200:
    cust_data = cust_res.json()
    print(f"\n--- Customers ({cust_data.get('totalElements', len(cust_data.get('data', [])))}) ---")
    for c in cust_data.get("data", []):
        c_id = c["id"]["id"]
        c_title = c.get("title")
        print(f"Customer ID: {c_id} | Title: {c_title}")
