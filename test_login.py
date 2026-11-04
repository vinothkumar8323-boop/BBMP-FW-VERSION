import requests
import json
import sys

BASE_URL = "https://schnelliot.in"
USERNAME = "vinoth.joel@schnellenergy.com"
PASSWORD = "vinoth555"

session = requests.Session()
session.headers.update({"Content-Type": "application/json"})

# 1. Login
login_url = f"{BASE_URL}/api/auth/login"
payload = {"username": USERNAME, "password": PASSWORD}

print(f"Logging in to {login_url}...")
res = session.post(login_url, json=payload)
print(f"Login status: {res.status_code}")

if res.status_code != 200:
    print("Login failed:", res.text)
    sys.exit(1)

data = res.json()
token = data.get("token")
print("Token retrieved successfully.")

session.headers.update({
    "X-Authorization": f"Bearer {token}",
    "Authorization": f"Bearer {token}"
})

# 2. Get User Info
user_res = session.get(f"{BASE_URL}/api/auth/user")
print("User Info Status:", user_res.status_code)
if user_res.status_code == 200:
    user_info = user_res.json()
    print("User Info:", json.dumps(user_info, indent=2))

# 3. Explore endpoints for devices / assets / device groups
endpoints_to_test = [
    "/api/user/devices",
    "/api/tenant/deviceInfos?pageSize=10&page=0",
    "/api/customer/deviceInfos?pageSize=10&page=0",
    "/api/deviceGroups/all",
    "/api/entityGroups/all",
    "/api/entityGroups/DEVICE",
    "/api/tenant/assets?pageSize=10&page=0",
    "/api/customer/assets?pageSize=10&page=0",
    "/api/dashboard/all"
]

for ep in endpoints_to_test:
    r = session.get(f"{BASE_URL}{ep}")
    print(f"GET {ep} => {r.status_code}")
    if r.status_code == 200:
        try:
            body = r.json()
            if isinstance(body, dict) and "data" in body:
                print(f"  Count: {len(body['data'])}, TotalElements: {body.get('totalElements')}")
                if body['data']:
                    print(f"  Sample 0: {json.dumps(body['data'][0], indent=2)[:300]}")
            elif isinstance(body, list):
                print(f"  Count: {len(body)}")
                if body:
                    print(f"  Sample 0: {json.dumps(body[0], indent=2)[:300]}")
            else:
                print(f"  Response: {json.dumps(body, indent=2)[:300]}")
        except Exception as e:
            print("  Parsing error:", e)
