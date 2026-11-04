import pandas as pd
import openpyxl
import os

file_path = "BBMP_Panels_Firmware_Versions_With_Locations.xlsx"
if not os.path.exists(file_path):
    file_path = "BBMP_Panels_Firmware_Versions.xlsx"

print(f"Loading Excel file: {file_path}")
wb = openpyxl.load_workbook(file_path)
print(f"Sheet names in Excel: {wb.sheetnames}")

main_sheet = "BBMP Panel FW & Locations" if "BBMP Panel FW & Locations" in wb.sheetnames else "BBMP Panel FW Details"
df_main = pd.read_excel(file_path, sheet_name=main_sheet, skiprows=3)
print(f"Main Data Rows: {len(df_main)}")
print("Columns:", list(df_main.columns))

print("\n--- Panel Location Completeness ---")
print(f"Total Panels: {len(df_main)}")
valid_locs = df_main['Location / Address'].dropna()
print(f"Panels with Valid Address: {len(valid_locs[~valid_locs.isin(['N/A', ''])])} / {len(df_main)}")
print(f"Panels with Latitude & Longitude: {df_main['Latitude'].notna().sum()} / {len(df_main)}")
print(f"Total Unique BBMP Zones: {df_main['Zone'].nunique()}")
print(f"Total Unique BBMP Wards: {df_main['Ward'].nunique()}")

print("\n--- Top 10 Zones by Panel Count ---")
print(df_main["Zone"].value_counts().head(10))

print("\n--- Firmware Version Breakdown ---")
print(df_main["Firmware Version"].value_counts())

print("\n--- GSM Firmware Breakdown ---")
print(df_main["GSM Firmware Version"].value_counts())

print("\n--- Region Breakdown ---")
print(df_main["Region"].value_counts())

print("\n--- Status Breakdown ---")
print(df_main["Status"].value_counts())

if "Communication Pending Days" in df_main.columns:
    print("\n--- Communication Pending Days Breakdown ---")
    p_days = pd.to_numeric(df_main["Communication Pending Days"], errors="coerce")
    bins = [-1, 0, 3, 7, 15, 30, 9999]
    labels = ["0 Days (Communicated Today)", "1 - 3 Days", "4 - 7 Days (Within 1 Week)", "8 - 15 Days (1 - 2 Weeks)", "16 - 30 Days (2 - 4 Weeks)", "> 30 Days (Over 1 Month)"]
    print(pd.cut(p_days, bins=bins, labels=labels).value_counts().to_string())
    print("\nMax Pending Days:", p_days.max(), "| Avg Pending Days:", round(p_days.mean(), 2))


