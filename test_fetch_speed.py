import requests
import json
import time
from concurrent.futures import ThreadPoolExecutor

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

# Fetch first 200 devices
r = session.get(f"{BASE_URL}/api/customer/{BBMP_CUSTOMER_ID}/deviceInfos?pageSize=200&page=0")
devices = r.json().get("data", [])
print(f"Testing speed on {len(devices)} devices...")

def fetch_device_data(dev):
    dev_id = dev["id"]["id"]
    attr_url = f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/attributes"
    try:
        res = session.get(attr_url, timeout=10)
        if res.status_code == 200:
            attrs = {a["key"]: a["value"] for a in res.json()}
            return dev_id, attrs, None
        return dev_id, {}, f"Status {res.status_code}"
    except Exception as e:
        return dev_id, {}, str(e)

start_time = time.time()
with ThreadPoolExecutor(max_workers=30) as executor:
    results = list(executor.map(fetch_device_data, devices))
elapsed = time.time() - start_time

success_count = sum(1 for r in results if r[1])
print(f"Fetched {success_count}/{len(devices)} device attributes in {elapsed:.2f} seconds!")
print(f"Rate: {len(devices)/elapsed:.1f} devices/sec")
