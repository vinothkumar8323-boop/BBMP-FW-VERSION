import requests
import json
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

# Fetch 200 devices
r = session.get(f"{BASE_URL}/api/customer/{BBMP_CUSTOMER_ID}/deviceInfos?pageSize=200&page=0")
devices = r.json().get("data", [])

def check_dev(dev):
    dev_id = dev["id"]["id"]
    # Attr
    attr_res = session.get(f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/attributes", timeout=10)
    attrs = {a["key"]: a["value"] for a in attr_res.json()} if attr_res.status_code == 200 else {}
    
    # Telemetry keys request (specifically ver, cs_ver, cs_gsmVersion, gsmVersion)
    keys = "ver,cs_ver,cs_gsmVersion,gsmVersion"
    ts_res = session.get(f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/timeseries?keys={keys}", timeout=10)
    ts = ts_res.json() if ts_res.status_code == 200 else {}
    ts_latest = {k: ts[k][0]["value"] for k in ts if ts[k]}
    
    attr_ver = attrs.get("ver")
    ts_ver = ts_latest.get("ver") or ts_latest.get("cs_ver")
    attr_gsm = attrs.get("gsmVersion")
    ts_gsm = ts_latest.get("cs_gsmVersion") or ts_latest.get("gsmVersion")
    
    return {
        "id": dev_id,
        "name": dev.get("name"),
        "attr_ver": attr_ver,
        "ts_ver": ts_ver,
        "attr_gsm": attr_gsm,
        "ts_gsm": ts_gsm
    }

with ThreadPoolExecutor(max_workers=30) as executor:
    results = list(executor.map(check_dev, devices))

diff_ver = 0
missing_ver = 0
for r in results:
    v1 = r["attr_ver"]
    v2 = r["ts_ver"]
    if not v1 and not v2:
        missing_ver += 1
    elif v1 != v2:
        diff_ver += 1

print(f"Total checked: {len(results)}")
print(f"Devices with missing firmware version: {missing_ver}")
print(f"Devices where attr_ver != ts_ver: {diff_ver}")
print("Sample comparisons:")
for r in results[:10]:
    print(r)
