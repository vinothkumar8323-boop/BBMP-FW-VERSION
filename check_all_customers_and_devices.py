import requests
import json

BASE_URL = "https://schnelliot.in"
USERNAME = "vinoth.joel@schnellenergy.com"
PASSWORD = "vinoth555"

session = requests.Session()
session.headers.update({"Content-Type": "application/json"})

# Login
res = session.post(f"{BASE_URL}/api/auth/login", json={"username": USERNAME, "password": PASSWORD})
token = res.json()["token"]
session.headers.update({"X-Authorization": f"Bearer {token}"})

# Search all customers with 'bbmp' or 'bangalore' in title (case-insensitive)
c_res = session.get(f"{BASE_URL}/api/customers?pageSize=1000&page=0")
if c_res.status_code == 200:
    customers = c_res.json().get("data", [])
    print(f"Total customers found: {len(customers)}")
    for c in customers:
        title = c.get("title", "")
        if "bbmp" in title.lower() or "bangalore" in title.lower() or "bgr" in title.lower():
            c_id = c["id"]["id"]
            # get device count
            dev_cnt_res = session.get(f"{BASE_URL}/api/customer/{c_id}/deviceInfos?pageSize=1&page=0")
            cnt = dev_cnt_res.json().get("totalElements", 0) if dev_cnt_res.status_code == 200 else 0
            print(f"Customer: '{title}' | ID: {c_id} | Device Count: {cnt}")

# Search device groups
eg_res = session.get(f"{BASE_URL}/api/entityGroups/DEVICE")
if eg_res.status_code == 200:
    groups = eg_res.json()
    print(f"\nDevice Entity Groups containing 'bbmp' or 'bangalore' or 'bgr':")
    for g in groups:
        g_name = g.get("name", "")
        if "bbmp" in g_name.lower() or "bangalore" in g_name.lower() or "bgr" in g_name.lower():
            print(f"  Group ID: {g['id']['id']} | Name: {g_name}")
