import requests
import json

BASE_URL = "https://schnelliot.in"
USERNAME = "vinoth.joel@schnellenergy.com"
PASSWORD = "vinoth555"
BBMP_CUSTOMER_ID = "e2119df0-45c3-11f0-94dc-77130b2f47e9"

session = requests.Session()
session.headers.update({"Content-Type": "application/json"})

# Login
res = session.post(f"{BASE_URL}/api/auth/login", json={"username": USERNAME, "password": PASSWORD})
token = res.json()["token"]
session.headers.update({"X-Authorization": f"Bearer {token}"})

print(f"--- Customer Info ---")
c_info = session.get(f"{BASE_URL}/api/customer/{BBMP_CUSTOMER_ID}").json()
print(json.dumps(c_info, indent=2))

endpoints = [
    f"/api/customer/{BBMP_CUSTOMER_ID}/deviceInfos?pageSize=10&page=0",
    f"/api/customer/{BBMP_CUSTOMER_ID}/devices?pageSize=10&page=0",
    f"/api/customer/{BBMP_CUSTOMER_ID}/entityGroups/DEVICE",
    f"/api/customer/{BBMP_CUSTOMER_ID}/entityGroups/ASSET",
    f"/api/customer/{BBMP_CUSTOMER_ID}/assets?pageSize=10&page=0",
    f"/api/relations/info?fromId={BBMP_CUSTOMER_ID}&fromType=CUSTOMER",
]

for ep in endpoints:
    r = session.get(f"{BASE_URL}{ep}")
    print(f"\nGET {ep} => {r.status_code}")
    if r.status_code == 200:
        data = r.json()
        print("Response summary:")
        if isinstance(data, dict):
            print("  Keys:", list(data.keys()))
            if "totalElements" in data:
                print("  totalElements:", data["totalElements"])
            if "data" in data and len(data["data"]) > 0:
                print("  Sample element 0:", json.dumps(data["data"][0], indent=2)[:400])
        elif isinstance(data, list):
            print(f"  List length: {len(data)}")
            if len(data) > 0:
                print("  Sample element 0:", json.dumps(data[0], indent=2)[:400])
