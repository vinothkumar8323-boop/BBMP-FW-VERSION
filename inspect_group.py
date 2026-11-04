import requests
import json

BASE_URL = "https://schnelliot.in"
USERNAME = "vinoth.joel@schnellenergy.com"
PASSWORD = "vinoth555"
GROUP_ID = "b62b9310-9574-11f1-b375-919c5dbde10c"

session = requests.Session()
session.headers.update({"Content-Type": "application/json"})

res = session.post(f"{BASE_URL}/api/auth/login", json={"username": USERNAME, "password": PASSWORD})
token = res.json()["token"]
session.headers.update({"X-Authorization": f"Bearer {token}"})

endpoints = [
    f"/api/entityGroup/{GROUP_ID}",
    f"/api/entityGroup/{GROUP_ID}/entities?pageSize=10&page=0",
    f"/api/entityGroup/{GROUP_ID}/devices?pageSize=10&page=0",
]

for ep in endpoints:
    r = session.get(f"{BASE_URL}{ep}")
    print(f"GET {ep} => {r.status_code}")
    if r.status_code == 200:
        d = r.json()
        if isinstance(d, dict) and "totalElements" in d:
            print("  TotalElements:", d["totalElements"])
            if d.get("data"):
                print("  Sample 0:", json.dumps(d["data"][0], indent=2)[:300])
        else:
            print("  Response:", json.dumps(d, indent=2)[:300])
