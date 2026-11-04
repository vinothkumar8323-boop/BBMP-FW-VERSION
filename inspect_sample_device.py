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

# 1. Fetch first 10 devices
d_res = session.get(f"{BASE_URL}/api/customer/{BBMP_CUSTOMER_ID}/deviceInfos?pageSize=10&page=0")
devices = d_res.json().get("data", [])

print(f"Fetched {len(devices)} sample devices")

for i, dev in enumerate(devices[:3]):
    dev_id = dev["id"]["id"]
    dev_name = dev.get("name")
    dev_type = dev.get("type")
    dev_label = dev.get("label")
    print(f"\n--- Device {i+1}: ID={dev_id}, Name={dev_name}, Label={dev_label}, Type={dev_type} ---")
    print("Full device object:", json.dumps(dev, indent=2))
    
    # Check attributes (CLIENT_SCOPE, SERVER_SCOPE, SHARED_SCOPE)
    attr_res = session.get(f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/attributes")
    if attr_res.status_code == 200:
        attrs = attr_res.json()
        print("Attributes:", json.dumps(attrs, indent=2))
    else:
        print(f"Attributes fetch status: {attr_res.status_code}")
        
    # Check latest telemetry
    ts_res = session.get(f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/timeseries")
    if ts_res.status_code == 200:
        ts = ts_res.json()
        print("Latest Telemetry Keys:", list(ts.keys()))
        # Print firmware related keys if any
        fw_keys = [k for k in ts.keys() if 'fw' in k.lower() or 'ver' in k.lower() or 'sw' in k.lower() or 'soft' in k.lower() or 'v' in k.lower()]
        print("Telemetry Keys containing fw/ver/sw/v:", fw_keys)
        for k in fw_keys[:10]:
            print(f"  {k}: {ts[k]}")
    else:
        print(f"Telemetry fetch status: {ts_res.status_code}")
