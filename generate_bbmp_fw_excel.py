import requests
import json
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_URL = "https://schnelliot.in"
USERNAME = "vinoth.joel@schnellenergy.com"
PASSWORD = "vinoth555"
BBMP_CUSTOMER_ID = "e2119df0-45c3-11f0-94dc-77130b2f47e9"
OUTPUT_FILE = "BBMP_Panels_Firmware_Versions.xlsx"

print("==================================================")
print("Schnell IoT - BBMP Panels Firmware Version Extractor")
print("==================================================")

session = requests.Session()
session.headers.update({"Content-Type": "application/json"})

# 1. Login
print("\n[1/5] Authenticating with Schnell IoT platform...")
login_res = session.post(f"{BASE_URL}/api/auth/login", json={"username": USERNAME, "password": PASSWORD})
if login_res.status_code != 200:
    print(f"Login failed: {login_res.status_code} - {login_res.text}")
    exit(1)

token = login_res.json()["token"]
session.headers.update({"X-Authorization": f"Bearer {token}"})
print("Authentication successful!")

# 2. Fetch all devices metadata for BBMP customer
print("\n[2/5] Fetching all BBMP devices metadata...")
all_devices = []
page = 0
page_size = 1000

while True:
    res = session.get(f"{BASE_URL}/api/customer/{BBMP_CUSTOMER_ID}/deviceInfos?pageSize={page_size}&page={page}")
    if res.status_code != 200:
        print(f"Failed to fetch devices page {page}: {res.status_code}")
        break
    data = res.json()
    items = data.get("data", [])
    all_devices.extend(items)
    print(f"  Page {page}: fetched {len(items)} devices. Total devices so far: {len(all_devices)}")
    if not data.get("hasNext", False):
        break
    page += 1

total_devices = len(all_devices)
print(f"Total BBMP devices identified: {total_devices}")

# 3. Fetch attributes for each device concurrently
print(f"\n[3/5] Fetching attributes & firmware details for {total_devices} devices in parallel...")

def fetch_attributes(dev):
    dev_id = dev["id"]["id"]
    attr_url = f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/attributes"
    try:
        r = session.get(attr_url, timeout=12)
        if r.status_code == 200:
            attrs = {a["key"]: a["value"] for a in r.json()}
            return dev_id, attrs
    except Exception as e:
        pass
    return dev_id, {}

start_time = time.time()
dev_attrs_map = {}

# Use 40 worker threads for fast parallel fetching
with ThreadPoolExecutor(max_workers=40) as executor:
    results = list(executor.map(fetch_attributes, all_devices))
    for dev_id, attrs in results:
        dev_attrs_map[dev_id] = attrs

elapsed = time.time() - start_time
print(f"Attributes fetched for {len(dev_attrs_map)} devices in {elapsed:.2f} seconds ({len(all_devices)/elapsed:.1f} devices/sec)!")

# Helper function to convert epoch ms to readable datetime string
def format_ts(ts):
    if not ts:
        return "N/A"
    try:
        # Check if epoch is in ms or sec
        val = int(ts)
        if val > 1e11:
            val = val / 1000.0
        return datetime.fromtimestamp(val).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return "N/A"

now_dt = datetime.now()

# Helper function to calculate communication pending days
def calc_pending_days(ts):
    if not ts:
        return "N/A"
    try:
        val = int(ts)
        if val > 1e11:
            val = val / 1000.0
        dt = datetime.fromtimestamp(val)
        days = (now_dt.date() - dt.date()).days
        return max(0, int(days))
    except Exception:
        return "N/A"

# Load resolved locations cache if available (for devices with blank ThingsBoard location string)
resolved_locations = {}
try:
    with open("resolved_locations.json", "r") as f:
        resolved_locations = json.load(f)
    print(f"Loaded {len(resolved_locations)} pre-resolved location addresses from cache.")
except Exception:
    pass

# 4. Construct Data Matrix
print("\n[4/5] Processing and formatting data...")
records = []

for idx, dev in enumerate(all_devices, 1):
    dev_id = dev["id"]["id"]
    dev_name = dev.get("name", "N/A")
    dev_label = dev.get("label", "N/A")
    dev_type = dev.get("type", "N/A")
    created_ts = format_ts(dev.get("createdTime"))
    
    attrs = dev_attrs_map.get(dev_id, {})
    
    fw_ver = attrs.get("ver", "N/A")
    gsm_ver = attrs.get("gsmVersion", "N/A")
    zone_name = attrs.get("zoneName", "N/A")
    ward_name = attrs.get("wardName", "N/A")
    region = attrs.get("region", "N/A")
    meter_sl_no = attrs.get("meterSerialNumber", "N/A")
    sim_name = attrs.get("simName") or attrs.get("gsmProvider") or "N/A"
    imei = attrs.get("imeiNumber", "N/A")
    
    is_active = attrs.get("active")
    if is_active is True:
        status_str = "Active"
    elif is_active is False:
        status_str = "Inactive"
    else:
        status_str = "N/A"
        
    state_str = attrs.get("state", "N/A")
    installed_by = attrs.get("installedBy", "N/A")
    installed_on = format_ts(attrs.get("installedOn"))
    
    # Prioritize lastUpdateTs (actual telemetry/meter data reception timestamp as shown in Schnell IoT)
    # Fallback to lastActivityTime if lastUpdateTs is not available
    last_rx_ts = attrs.get("lastUpdateTs") or attrs.get("lastActivityTime")
    last_communicated = format_ts(last_rx_ts)
    pending_days = calc_pending_days(last_rx_ts)
    
    last_connect = format_ts(attrs.get("lastConnectTime"))
    last_disconnect = format_ts(attrs.get("lastDisconnectTime"))
    
    lat = attrs.get("latitude", "N/A")
    lng = attrs.get("longitude", "N/A")
    
    # 1. Resolve exact physical location/address
    loc_val = attrs.get("location") or resolved_locations.get(dev_name) or ""
    loc_val = str(loc_val).strip()
    if not loc_val or loc_val == "N/A":
        if lat not in ["N/A", None, ""] and lng not in ["N/A", None, ""]:
            loc_val = f"Ward: {ward_name}, Zone: {zone_name}, Bengaluru (Lat: {lat}, Lng: {lng})"
        else:
            loc_val = "N/A"
            
    # 2. Formulate Google Maps link
    maps_link = "N/A"
    if lat not in ["N/A", None, ""] and lng not in ["N/A", None, ""]:
        maps_link = f"https://www.google.com/maps?q={lat},{lng}"
    
    # Organize columns logically: Panel ID -> Location & Maps Link -> Firmware -> Status & Pending Days -> Dates -> HW Specs
    records.append({
        "Sl. No.": idx,
        "Panel Name": dev_name,
        "Panel Label / Serial": dev_label,
        "Location / Address": loc_val,
        "Zone": zone_name,
        "Ward": ward_name,
        "Region": region,
        "Google Maps Link": maps_link,
        "Latitude": lat,
        "Longitude": lng,
        "Firmware Version": fw_ver,
        "GSM Firmware Version": gsm_ver,
        "Status": status_str,
        "Communication Pending Days": pending_days,
        "Received / Last Communicated Date & Time": last_communicated,
        "Last Connect Date & Time": last_connect,
        "Last Disconnect Date & Time": last_disconnect,
        "Installation State": state_str,
        "Installed By": installed_by,
        "Installed Date": installed_on,
        "Created Date": created_ts,
        "Meter Serial Number": meter_sl_no,
        "SIM / Network": sim_name,
        "IMEI": imei,
        "Device Type": dev_type,
        "Device ID": dev_id
    })

df_main = pd.DataFrame(records)

# Create Summary Statistics DataFrames
df_fw_summary = df_main["Firmware Version"].value_counts().reset_index()
df_fw_summary.columns = ["Firmware Version", "Panel Count"]
df_fw_summary["Percentage (%)"] = (df_fw_summary["Panel Count"] / len(df_main) * 100).round(2)

df_gsm_summary = df_main["GSM Firmware Version"].value_counts().reset_index()
df_gsm_summary.columns = ["GSM Firmware Version", "Panel Count"]
df_gsm_summary["Percentage (%)"] = (df_gsm_summary["Panel Count"] / len(df_main) * 100).round(2)

df_region_summary = df_main["Region"].value_counts().reset_index()
df_region_summary.columns = ["Region", "Panel Count"]
df_region_summary["Percentage (%)"] = (df_region_summary["Panel Count"] / len(df_main) * 100).round(2)

df_status_summary = df_main["Status"].value_counts().reset_index()
df_status_summary.columns = ["Status", "Panel Count"]
df_status_summary["Percentage (%)"] = (df_status_summary["Panel Count"] / len(df_main) * 100).round(2)

df_zone_summary = df_main["Zone"].value_counts().reset_index()
df_zone_summary.columns = ["Zone", "Panel Count"]
df_zone_summary["Percentage (%)"] = (df_zone_summary["Panel Count"] / len(df_main) * 100).round(2)

# Pending Days Summary
def categorize_pending(d):
    if d == "N/A" or pd.isna(d):
        return "Unknown / N/A"
    try:
        val = int(d)
        if val == 0:
            return "0 Days (Communicated Today)"
        elif 1 <= val <= 3:
            return "1 - 3 Days"
        elif 4 <= val <= 7:
            return "4 - 7 Days (Within 1 Week)"
        elif 8 <= val <= 15:
            return "8 - 15 Days (1 - 2 Weeks)"
        elif 16 <= val <= 30:
            return "16 - 30 Days (2 - 4 Weeks)"
        else:
            return "> 30 Days (Over 1 Month)"
    except Exception:
        return "Unknown / N/A"

pending_categories = [
    "0 Days (Communicated Today)",
    "1 - 3 Days",
    "4 - 7 Days (Within 1 Week)",
    "8 - 15 Days (1 - 2 Weeks)",
    "16 - 30 Days (2 - 4 Weeks)",
    "> 30 Days (Over 1 Month)",
    "Unknown / N/A"
]

cat_series = df_main["Communication Pending Days"].apply(categorize_pending)
df_pending_summary = cat_series.value_counts().reindex(pending_categories).dropna().reset_index()
df_pending_summary.columns = ["Communication Pending Range", "Panel Count"]
df_pending_summary["Panel Count"] = df_pending_summary["Panel Count"].astype(int)
df_pending_summary = df_pending_summary[df_pending_summary["Panel Count"] > 0]
df_pending_summary["Percentage (%)"] = (df_pending_summary["Panel Count"] / len(df_main) * 100).round(2)



# 5. Build Styled Excel Workbook
print(f"\n[5/5] Generating professional Excel file: {OUTPUT_FILE}...")

wb = openpyxl.Workbook()

# Style definitions
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Dark navy blue
header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Light ice blue
section_font = Font(name="Segoe UI", size=12, bold=True, color="1F4E78")

title_font = Font(name="Segoe UI", size=16, bold=True, color="1F4E78")
subtitle_font = Font(name="Segoe UI", size=10, italic=True, color="595959")
bold_font = Font(name="Segoe UI", size=10, bold=True)
regular_font = Font(name="Segoe UI", size=10)
link_font = Font(name="Segoe UI", size=10, color="0563C1", underline="single")

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

# --- Sheet 1: Main Panel Firmware & Location Details ---
ws_main = wb.active
ws_main.title = "BBMP Panel FW & Locations"
ws_main.views.sheetView[0].showGridLines = True

# Title block
ws_main.cell(row=1, column=1, value="BBMP PANELS FIRMWARE & LOCATION DETAILS").font = title_font
ws_main.cell(row=2, column=1, value=f"Schnell IoT Platform (schnelliot.in) | Total Panels: {len(df_main)} | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}").font = subtitle_font

# Freeze panes at D5: Columns A, B, C (Sl. No., Panel Name, Panel Label) remain fixed while scrolling horizontally
ws_main.freeze_panes = "D5"

# Header row start at row 4
start_row = 4
cols = list(df_main.columns)
ws_main.row_dimensions[start_row].height = 28

for c_idx, col_name in enumerate(cols, 1):
    cell = ws_main.cell(row=start_row, column=c_idx, value=col_name)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Data rows
for r_idx, row_data in enumerate(df_main.values, start_row + 1):
    ws_main.row_dimensions[r_idx].height = 20
    for c_idx, val in enumerate(row_data, 1):
        cell = ws_main.cell(row=r_idx, column=c_idx, value=val)
        cell.font = regular_font
        cell.border = thin_border
        
        # Zebra striping
        if r_idx % 2 == 0:
            cell.fill = zebra_fill
            
        col_name = cols[c_idx - 1]
        
        # Clickable Hyperlink for Google Maps
        if col_name == "Google Maps Link" and val and str(val).startswith("http"):
            cell.value = "View on Google Maps"
            cell.hyperlink = str(val)
            cell.font = link_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_name in [
            "Sl. No.", "Status", "Communication Pending Days", "Firmware Version", "GSM Firmware Version", 
            "Device Type", "Region", "Zone", "Ward", "Installed Date", 
            "Received / Last Communicated Date & Time", 
            "Last Connect Date & Time", "Last Disconnect Date & Time", "Created Date",
            "Google Maps Link"
        ]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if col_name == "Communication Pending Days" and isinstance(val, (int, float)):
                cell.number_format = '#,##0'
        elif col_name in ["Latitude", "Longitude"]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
        elif col_name == "Location / Address":
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

# Enable AutoFilter
ws_main.auto_filter.ref = f"A{start_row}:{openpyxl.utils.get_column_letter(len(cols))}{len(df_main) + start_row}"

# Column widths
for col in ws_main.columns:
    col_letter = openpyxl.utils.get_column_letter(col[0].column)
    max_len = 0
    for cell in col:
        if cell.row < (start_row - 1):
            continue
        val_str = str(cell.value or "")
        if len(val_str) > max_len:
            max_len = len(val_str)
    ws_main.column_dimensions[col_letter].width = max(max_len + 4, 12)

# Specific custom widths for high-priority columns
ws_main.column_dimensions['A'].width = 8   # Sl. No.
ws_main.column_dimensions['B'].width = 16  # Panel Name
ws_main.column_dimensions['C'].width = 18  # Panel Label / Serial
ws_main.column_dimensions['D'].width = 52  # Location / Address (prominent wide column)
ws_main.column_dimensions['E'].width = 18  # Zone
ws_main.column_dimensions['F'].width = 16  # Ward
ws_main.column_dimensions['G'].width = 14  # Region
ws_main.column_dimensions['H'].width = 22  # Google Maps Link
ws_main.column_dimensions['I'].width = 14  # Latitude
ws_main.column_dimensions['J'].width = 14  # Longitude
ws_main.column_dimensions['K'].width = 18  # Firmware Version
ws_main.column_dimensions['L'].width = 22  # GSM Firmware Version
ws_main.column_dimensions['M'].width = 14  # Status
ws_main.column_dimensions['N'].width = 26  # Communication Pending Days
ws_main.column_dimensions['O'].width = 28  # Received / Last Communicated Date & Time


# --- Sheet 2: Summary Dashboard ---
ws_sum = wb.create_sheet(title="Firmware & Zone Dashboard")
ws_sum.views.sheetView[0].showGridLines = True

ws_sum.cell(row=1, column=1, value="BBMP PANELS FIRMWARE & LOCATION SUMMARY").font = title_font
ws_sum.cell(row=2, column=1, value=f"Total Panels Audited: {len(df_main)}").font = subtitle_font

def write_summary_table(ws, start_r, start_c, title, df_sum):
    ws.cell(row=start_r, column=start_c, value=title).font = section_font
    ws.merge_cells(start_row=start_r, start_column=start_c, end_row=start_r, end_column=start_c + len(df_sum.columns) - 1)
    ws.cell(row=start_r, column=start_c).fill = section_fill
    
    # Headers
    h_row = start_r + 1
    for c_i, col in enumerate(df_sum.columns, start_c):
        cell = ws.cell(row=h_row, column=c_i, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    # Data
    d_row = h_row + 1
    for r_i, r_val in enumerate(df_sum.values, d_row):
        for c_i, val in enumerate(r_val, start_c):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            cell.font = regular_font
            cell.border = thin_border
            if c_i > start_c:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
    # Total row
    t_row = d_row + len(df_sum)
    cell_label = ws.cell(row=t_row, column=start_c, value="Total")
    cell_label.font = bold_font
    cell_label.border = thin_border
    
    cell_cnt = ws.cell(row=t_row, column=start_c + 1, value=df_sum.iloc[:, 1].sum())
    cell_cnt.font = bold_font
    cell_cnt.border = thin_border
    cell_cnt.alignment = Alignment(horizontal="right", vertical="center")
    
    cell_pct = ws.cell(row=t_row, column=start_c + 2, value=100.0)
    cell_pct.font = bold_font
    cell_pct.border = thin_border
    cell_pct.alignment = Alignment(horizontal="right", vertical="center")
    
    return t_row + 2

r1 = write_summary_table(ws_sum, 4, 1, "Firmware Version Breakdown (ver)", df_fw_summary)
r2 = write_summary_table(ws_sum, r1, 1, "GSM Module Firmware Breakdown (gsmVersion)", df_gsm_summary)

r3 = write_summary_table(ws_sum, 4, 5, "Region-wise Panel Distribution", df_region_summary)
r4 = write_summary_table(ws_sum, r3, 5, "Panel Online/Active Status", df_status_summary)
r5 = write_summary_table(ws_sum, r4, 5, "Communication Pending Days Breakdown", df_pending_summary)

r6 = write_summary_table(ws_sum, 4, 9, "Zone-wise Panel Distribution", df_zone_summary)

# Adjust widths for summary sheet
for col in ws_sum.columns:
    col_letter = openpyxl.utils.get_column_letter(col[0].column)
    ws_sum.column_dimensions[col_letter].width = 30

# --- Sheet 3: Offline > 7 Days Monitoring (145 Panels) ---
print("\n[5b] Updating 'Offline > 7 Days' monitoring sheet (145 panels)...")
try:
    from update_offline_7_days import update_offline_sheet_in_workbook
    update_offline_sheet_in_workbook(wb, session=session, now_dt=now_dt)
except Exception as e:
    print(f"[Notice] Could not update 'Offline > 7 Days' sheet: {e}")

saved_excel_file = OUTPUT_FILE
try:
    wb.save(OUTPUT_FILE)
    print(f"\nSUCCESS! Excel workbook saved as: {OUTPUT_FILE}")
except PermissionError:
    saved_excel_file = "BBMP_Panels_Firmware_Versions_With_Locations.xlsx"
    wb.save(saved_excel_file)
    print(f"\nNOTICE: '{OUTPUT_FILE}' is currently open in Excel or another program.")
    print(f"SUCCESS! Workbook saved as: {saved_excel_file}")

# 6. Automatic Sync to Google Sheets
print("\n[6/6] Syncing generated data to Google Sheets...")
try:
    from upload_to_gsheet import upload_excel_to_gsheet
    upload_excel_to_gsheet(excel_path=saved_excel_file)
except Exception as e:
    print(f"[Google Sheets Sync Notice] Could not auto-upload: {e}")
    print("Run 'python upload_to_gsheet.py' manually after setting up credentials.")

print("==================================================")


