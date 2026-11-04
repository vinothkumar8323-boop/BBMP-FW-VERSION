import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import json
import os

SOURCE_FILE = "BBMP_Panels_Firmware_Versions.xlsx"
OUTPUT_FILE = "BBMP_Panels_Firmware_Versions.xlsx"
FALLBACK_FILE = "BBMP_Panels_Firmware_Versions_With_Locations.xlsx"

print("==================================================")
print("Applying Enhanced Panel Locations & Pending Days to Excel")
print("==================================================")

# 1. Load resolved locations cache
resolved_locations = {}
if os.path.exists("resolved_locations.json"):
    with open("resolved_locations.json", "r") as f:
        resolved_locations = json.load(f)
    print(f"Loaded {len(resolved_locations)} pre-resolved addresses.")

# 2. Read existing data
print(f"Reading existing panel dataset from: {SOURCE_FILE}...")
df_raw = pd.read_excel(SOURCE_FILE, sheet_name=0, skiprows=3)
print(f"Total panels loaded: {len(df_raw)}")

now_dt = datetime.now()

def calc_pending_from_str(date_val):
    if date_val is None or pd.isna(date_val) or str(date_val).strip() in ["", "N/A", "nan"]:
        return "N/A"
    try:
        dt = pd.to_datetime(date_val)
        days = (now_dt.date() - dt.date()).days
        return max(0, int(days))
    except Exception:
        return "N/A"

# 3. Clean and enhance records
records = []
for idx, row in df_raw.iterrows():
    dev_name = str(row.get("Panel Name") or "N/A")
    dev_label = str(row.get("Panel Label / Serial") or "N/A")
    lat = row.get("Latitude")
    lng = row.get("Longitude")
    zone_name = str(row.get("Zone") or "N/A")
    ward_name = str(row.get("Ward") or "N/A")
    region = str(row.get("Region") or "N/A")
    
    # Location resolution
    loc_val = row.get("Location / Address")
    if pd.isna(loc_val) or str(loc_val).strip() in ["", "N/A", "nan"]:
        loc_val = resolved_locations.get(dev_name, "")
    loc_val = str(loc_val).strip()
    if not loc_val or loc_val in ["N/A", "nan"]:
        if pd.notna(lat) and pd.notna(lng) and str(lat) != "N/A":
            loc_val = f"Ward: {ward_name}, Zone: {zone_name}, Bengaluru (Lat: {lat}, Lng: {lng})"
        else:
            loc_val = "N/A"
            
    # Google Maps URL
    maps_link = "N/A"
    if pd.notna(lat) and pd.notna(lng) and str(lat) not in ["N/A", "nan", ""]:
        maps_link = f"https://www.google.com/maps?q={lat},{lng}"
        
    # Communication pending days calculation
    last_comm_val = row.get("Received / Last Communicated Date & Time")
    pending_days = calc_pending_from_str(last_comm_val)
        
    records.append({
        "Sl. No.": idx + 1,
        "Panel Name": dev_name,
        "Panel Label / Serial": dev_label,
        "Location / Address": loc_val,
        "Zone": zone_name,
        "Ward": ward_name,
        "Region": region,
        "Google Maps Link": maps_link,
        "Latitude": lat if pd.notna(lat) else "N/A",
        "Longitude": lng if pd.notna(lng) else "N/A",
        "Firmware Version": row.get("Firmware Version", "N/A"),
        "GSM Firmware Version": row.get("GSM Firmware Version", "N/A"),
        "Status": row.get("Status", "N/A"),
        "Communication Pending Days": pending_days,
        "Received / Last Communicated Date & Time": last_comm_val if pd.notna(last_comm_val) else "N/A",
        "Last Connect Date & Time": row.get("Last Connect Date & Time", "N/A"),
        "Last Disconnect Date & Time": row.get("Last Disconnect Date & Time", "N/A"),
        "Installation State": row.get("Installation State", "N/A"),
        "Installed By": row.get("Installed By", "N/A"),
        "Installed Date": row.get("Installed Date", "N/A"),
        "Created Date": row.get("Created Date", "N/A"),
        "Meter Serial Number": row.get("Meter Serial Number", "N/A"),
        "SIM / Network": row.get("SIM / Network", "N/A"),
        "IMEI": row.get("IMEI", "N/A"),
        "Device Type": row.get("Device Type", "N/A"),
        "Device ID": row.get("Device ID", "N/A")
    })

df_main = pd.DataFrame(records)

# Summary tables
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

# Build workbook
wb = openpyxl.Workbook()

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
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

# Sheet 1: Main Panel FW & Locations
ws_main = wb.active
ws_main.title = "BBMP Panel FW & Locations"
ws_main.views.sheetView[0].showGridLines = True

ws_main.cell(row=1, column=1, value="BBMP PANELS FIRMWARE & LOCATION DETAILS").font = title_font
ws_main.cell(row=2, column=1, value=f"Total Panels: {len(df_main)} | Location & Pending Days Added | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}").font = subtitle_font

# Freeze panes at D5: Columns A-C remain fixed while scrolling horizontally
ws_main.freeze_panes = "D5"

start_row = 4
cols = list(df_main.columns)
ws_main.row_dimensions[start_row].height = 28

for c_idx, col_name in enumerate(cols, 1):
    cell = ws_main.cell(row=start_row, column=c_idx, value=col_name)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for r_idx, row_data in enumerate(df_main.values, start_row + 1):
    ws_main.row_dimensions[r_idx].height = 20
    for c_idx, val in enumerate(row_data, 1):
        cell = ws_main.cell(row=r_idx, column=c_idx, value=val)
        cell.font = regular_font
        cell.border = thin_border
        
        if r_idx % 2 == 0:
            cell.fill = zebra_fill
            
        col_name = cols[c_idx - 1]
        
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

ws_main.auto_filter.ref = f"A{start_row}:{openpyxl.utils.get_column_letter(len(cols))}{len(df_main) + start_row}"

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

ws_main.column_dimensions['A'].width = 8   # Sl. No.
ws_main.column_dimensions['B'].width = 16  # Panel Name
ws_main.column_dimensions['C'].width = 18  # Panel Label / Serial
ws_main.column_dimensions['D'].width = 52  # Location / Address
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

# Sheet 2: Summary Dashboard
ws_sum = wb.create_sheet(title="Firmware & Zone Dashboard")
ws_sum.views.sheetView[0].showGridLines = True

ws_sum.cell(row=1, column=1, value="BBMP PANELS FIRMWARE & LOCATION SUMMARY").font = title_font
ws_sum.cell(row=2, column=1, value=f"Total Panels Audited: {len(df_main)}").font = subtitle_font

def write_summary_table(ws, start_r, start_c, title, df_sum):
    ws.cell(row=start_r, column=start_c, value=title).font = section_font
    ws.merge_cells(start_row=start_r, start_column=start_c, end_row=start_r, end_column=start_c + len(df_sum.columns) - 1)
    ws.cell(row=start_r, column=start_c).fill = section_fill
    
    h_row = start_r + 1
    for c_i, col in enumerate(df_sum.columns, start_c):
        cell = ws.cell(row=h_row, column=c_i, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
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

for col in ws_sum.columns:
    col_letter = openpyxl.utils.get_column_letter(col[0].column)
    ws_sum.column_dimensions[col_letter].width = 30

saved_path = OUTPUT_FILE
try:
    wb.save(OUTPUT_FILE)
    print(f"\nSUCCESS! Enhanced Excel workbook saved as: {OUTPUT_FILE}")
except PermissionError:
    saved_path = FALLBACK_FILE
    wb.save(FALLBACK_FILE)
    print(f"\nNOTICE: '{OUTPUT_FILE}' is currently open in Microsoft Excel.")
    print(f"SUCCESS! Enhanced Excel workbook saved as: {FALLBACK_FILE}")

# Also update the fallback file so both are always in sync
try:
    wb.save(FALLBACK_FILE)
    print(f"Also synced to: {FALLBACK_FILE}")
except Exception:
    pass

print(f"Total Rows: {len(df_main)} with full Location, Zone, Ward, Region, Google Maps Link, Lat, Lng, Communication Pending Days.")
print("==================================================")
