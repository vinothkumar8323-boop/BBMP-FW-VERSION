import requests
import json
from collections import Counter

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

# 1. Fetch all devices list metadata (without telemetry yet) to check count and device types
print("Fetching device list pages...")
all_devices = []
page = 0
page_size = 1000

while True:
    r = session.get(f"{BASE_URL}/api/customer/{BBMP_CUSTOMER_ID}/deviceInfos?pageSize={page_size}&page={page}")
    if r.status_code != 200:
        print(f"Error fetching page {page}: {r.status_code}")
        break
    data = r.json()
    items = data.get("data", [])
    all_devices.extend(items)
    print(f"Page {page}: fetched {len(items)} items. Total so far: {len(all_devices)}")
    if not data.get("hasNext", False):
        break
    page += 1

print(f"\nTotal BBMP Devices Fetched: {len(all_devices)}")

types_counter = Counter(d.get("type", "UNKNOWN") for d in all_devices)
print("Device Types Distribution:")
for t, cnt in types_counter.most_common():
    print(f"  {t}: {cnt}")

# 2. Sample devices from each type to inspect attribute and telemetry keys
print("\nSampling attributes and telemetry across device types...")
for d_type in types_counter.keys():
    samples = [d for d in all_devices if d.get("type") == d_type][:5]
    print(f"\n================ Device Type: {d_type} (Sample size: {len(samples)}) ================")
    for dev in samples:
        dev_id = dev["id"]["id"]
        dev_name = dev.get("name")
        dev_label = dev.get("label")
        
        # Attributes
        attr_r = session.get(f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/attributes")
        attrs = attr_r.json() if attr_r.status_code == 200 else []
        attr_dict = {a["key"]: a["value"] for a in attrs}
        
        # Telemetry
        ts_r = session.get(f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/timeseries")
        ts = ts_r.json() if ts_r.status_code == 200 else {}
        ts_latest = {k: ts[k][0]["value"] for k in ts if ts[k]}
        
        # Find potential firmware keys in attributes and telemetry
        fw_attrs = {k: v for k, v in attr_dict.items() if 'ver' in k.lower() or 'fw' in k.lower() or 'sw' in k.lower() or 'soft' in k.lower()}
        fw_ts = {k: v for k, v in ts_latest.items() if 'ver' in k.lower() or 'fw' in k.lower() or 'sw' in k.lower() or 'soft' in k.lower()}
        
        print(f"Name: {dev_name} | Label: {dev_label}")
        print(f"  FW in Attributes: {fw_attrs}")
        print(f"  FW in Telemetry: {fw_ts}")
        print(f"  Sample Attr Keys: {list(attr_dict.keys())[:10]}")
        print(f"  Sample TS Keys: {list(ts_latest.keys())[:10]}")
