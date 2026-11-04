"""
update_offline_7_days.py
========================
Manages the "Offline > 7 Days" sheet in BBMP_Panels_Firmware_Versions.xlsx:
- Retains EXACTLY the 145 monitored panels (no additions, no removals).
- Updates:
    * Status (Online / Offline)
    * Received / Last Communicated Date & Time
    * Communication Pending Days (calculated up to current date)
    * Signal Strength (dBm)
    * RSSI value
    * Column W: Signal Strength (Good/Medium/Low)
- Preserves Column X and beyond: DOES NOT erase or overwrite user text/changes after Column X.
- Light Green Highlight: Applied ONLY for panels with Communication Pending Days == 0
  (i.e. latest communicate / communicated today).
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import requests
import json
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

BASE_URL = "https://schnelliot.in"
USERNAME = "vinoth.joel@schnellenergy.com"
PASSWORD = "vinoth555"
EXCEL_FILE = "BBMP_Panels_Firmware_Versions.xlsx"
FALLBACK_FILE = "BBMP_Panels_Firmware_Versions_With_Locations.xlsx"
CACHE_FILE = "offline_145_panels.json"
SHEET_NAME = "Offline > 7 Days"

# Fixed system columns: Columns A (1) to W (23)
# Column X (24) and beyond are user columns that are NEVER erased or overwritten!
COLUMNS = [
    ("Sl. No.", 8, "center"),                                    # Col 1 (A)
    ("Panel Name", 16, "center"),                                 # Col 2 (B)
    ("Panel Label / Serial", 18, "center"),                       # Col 3 (C)
    ("Location / Address", 48, "left"),                           # Col 4 (D)
    ("Zone", 18, "center"),                                       # Col 5 (E)
    ("Ward", 16, "center"),                                       # Col 6 (F)
    ("Region", 14, "center"),                                     # Col 7 (G)
    ("Latitude", 14, "right"),                                    # Col 8 (H)
    ("Longitude", 14, "right"),                                   # Col 9 (I)
    ("Firmware Version", 18, "center"),                           # Col 10 (J)
    ("GSM Firmware Version", 22, "center"),                       # Col 11 (K)
    ("Status", 14, "center"),                                     # Col 12 (L)
    ("Communication Pending Days", 26, "center"),                 # Col 13 (M)
    ("Received / Last Communicated Date & Time", 30, "center"),   # Col 14 (N)
    ("Signal Strength (dBm)", 22, "center"),                      # Col 15 (O)
    ("RSSI", 12, "center"),                                       # Col 16 (P)
    ("Installation State", 18, "center"),                         # Col 17 (Q)
    ("Installed Date", 22, "center"),                             # Col 18 (R)
    ("Created Date", 22, "center"),                               # Col 19 (S)
    ("Meter Serial Number", 20, "center"),                        # Col 20 (T)
    ("SIM / Network", 16, "center"),                              # Col 21 (U)
    ("IMEI", 18, "center"),                                       # Col 22 (V)
    ("Signal Strength (Good/Medium/Low)", 26, "center"),          # Col 23 (W)
]


def classify_signal_quality(rssi, dbm):
    """
    Classifies signal strength into Good, Medium, or Low:
    - Good:   RSSI >= 15 (or dBm >= -85 dBm)
    - Medium: RSSI 10 to 14 (or dBm -95 to -86 dBm)
    - Low:    RSSI 0 to 9, 99 (or dBm <= -96 dBm)
    """
    if rssi is not None and str(rssi).strip() not in ["", "N/A", "None"]:
        try:
            r = int(str(rssi).strip())
            if r == 99:
                return "Low"
            elif r >= 15:
                return "Good"
            elif 10 <= r <= 14:
                return "Medium"
            elif 0 <= r <= 9:
                return "Low"
        except Exception:
            pass

    if dbm is not None and str(dbm).strip() not in ["", "N/A", "None"]:
        try:
            clean_dbm = str(dbm).lower().replace("dbm", "").strip()
            d = float(clean_dbm)
            if d >= -85:
                return "Good"
            elif -95 <= d <= -86:
                return "Medium"
            else:
                return "Low"
        except Exception:
            pass

    return "N/A"


def load_145_panels_base():
    """Load the base 145 panels list from JSON cache or existing Excel sheet."""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r") as f:
                data = json.load(f)
                if len(data) == 145:
                    return data
        except Exception:
            pass

    for fpath in [EXCEL_FILE, FALLBACK_FILE]:
        if os.path.exists(fpath):
            try:
                wb = openpyxl.load_workbook(fpath, data_only=True)
                if SHEET_NAME in wb.sheetnames:
                    ws = wb[SHEET_NAME]
                    panels = []
                    for r in range(5, ws.max_row + 1):
                        name = ws.cell(r, 2).value
                        if not name:
                            continue
                        panels.append({
                            "sl_no": ws.cell(r, 1).value,
                            "panel_name": str(name).strip(),
                            "panel_label": str(ws.cell(r, 3).value or ""),
                            "location": str(ws.cell(r, 4).value or ""),
                            "zone": str(ws.cell(r, 5).value or ""),
                            "ward": str(ws.cell(r, 6).value or ""),
                            "region": str(ws.cell(r, 7).value or ""),
                            "lat": ws.cell(r, 8).value,
                            "lng": ws.cell(r, 9).value,
                            "fw_ver": str(ws.cell(r, 10).value or "N/A"),
                            "gsm_ver": str(ws.cell(r, 11).value or "N/A"),
                            "status": str(ws.cell(r, 12).value or "Offline"),
                            "pdays": ws.cell(r, 13).value,
                            "comm_date": str(ws.cell(r, 14).value or ""),
                            "dbm": str(ws.cell(r, 15).value or "N/A"),
                            "rssi": str(ws.cell(r, 16).value or "N/A"),
                            "install_state": str(ws.cell(r, 17).value or ""),
                            "install_date": str(ws.cell(r, 18).value or ""),
                            "created_date": str(ws.cell(r, 19).value or ""),
                            "meter_sl_no": str(ws.cell(r, 20).value or ""),
                            "sim_network": str(ws.cell(r, 21).value or ""),
                            "imei": str(ws.cell(r, 22).value or ""),
                        })
                    if panels:
                        return panels
            except Exception:
                pass
    return []


def update_offline_sheet_in_workbook(wb, session=None, now_dt=None):
    """
    Populates or refreshes the 'Offline > 7 Days' worksheet in an openpyxl Workbook.
    - Updates Columns A to W.
    - Preserves Column X (24) and beyond intact!
    - Highlights in light green ONLY when Communication Pending Days == 0.
    """
    if now_dt is None:
        now_dt = datetime.now()

    if session is None:
        session = requests.Session()
        session.headers.update({"Content-Type": "application/json"})
        login_res = session.post(f"{BASE_URL}/api/auth/login", json={"username": USERNAME, "password": PASSWORD})
        if login_res.status_code != 200:
            print(f"[Offline Sheet] Login failed: {login_res.status_code}")
            return False
        token = login_res.json()["token"]
        session.headers.update({"X-Authorization": f"Bearer {token}"})

    base_panels = load_145_panels_base()
    if not base_panels:
        print("[Offline Sheet] No 145 panels base data found.")
        return False

    print(f"\n[Offline Sheet] Refreshing live data for {len(base_panels)} panels...")

    def fetch_live(p):
        name = p["panel_name"]
        item = dict(p)
        try:
            r = session.get(f"{BASE_URL}/api/tenant/devices?deviceName={name}", timeout=10)
            if r.status_code == 200:
                dev = r.json()
                dev_id = dev["id"]["id"]

                # Attributes
                ar = session.get(f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/attributes", timeout=10)
                attrs = {a["key"]: a["value"] for a in ar.json()} if ar.status_code == 200 else {}

                # Timeseries
                tr = session.get(f"{BASE_URL}/api/plugins/telemetry/DEVICE/{dev_id}/values/timeseries?keys=rssi,dBm", timeout=10)
                ts = tr.json() if tr.status_code == 200 else {}

                last_rx = attrs.get("lastUpdateTs") or attrs.get("lastActivityTime")
                if last_rx:
                    val = int(last_rx) / 1000.0 if int(last_rx) > 1e11 else int(last_rx)
                    dt = datetime.fromtimestamp(val)
                    dt_str = dt.strftime("%Y-%m-%d %H:%M:%S")
                    days_pending = max(0, (now_dt.date() - dt.date()).days)
                    item["comm_date"] = dt_str
                    item["pdays"] = days_pending

                rssi_val = ts.get("rssi", [{}])[0].get("value")
                dbm_val = ts.get("dBm", [{}])[0].get("value")
                if rssi_val is not None:
                    item["rssi"] = str(rssi_val)
                    if dbm_val is None:
                        try:
                            r_int = int(rssi_val)
                            if 0 <= r_int <= 31:
                                dbm_val = str(2 * r_int - 113)
                        except Exception:
                            pass
                if dbm_val is not None:
                    item["dbm"] = f"{dbm_val} dBm" if "dBm" not in str(dbm_val) else str(dbm_val)

                is_active = attrs.get("active")
                if is_active is True or item.get("pdays") == 0:
                    item["status"] = "Online"
                else:
                    item["status"] = "Offline"

                if attrs.get("ver"):
                    item["fw_ver"] = attrs.get("ver")
                if attrs.get("gsmVersion"):
                    item["gsm_ver"] = attrs.get("gsmVersion")

        except Exception:
            pass

        # Classify Signal Quality: Good / Medium / Low (Column W)
        item["signal_quality"] = classify_signal_quality(item.get("rssi"), item.get("dbm"))

        # HIGHLIGHT RULE: ONLY for pending days == 0 (latest communicate / communicated today)
        item["highlight_green"] = (item.get("pdays") == 0)
        return item

    with ThreadPoolExecutor(max_workers=30) as ex:
        updated = list(ex.map(fetch_live, base_panels))

    comm_today_count = sum(1 for p in updated if p.get("highlight_green"))
    print(f"[Offline Sheet] Monitored: {len(updated)} panels | Communicated Today (0 days): {comm_today_count} | Offline: {len(updated) - comm_today_count}")

    # Get or create worksheet in workbook
    if SHEET_NAME in wb.sheetnames:
        ws = wb[SHEET_NAME]
        # SAFE CLEAR: ONLY clear columns 1 to 23 (A to W).
        # NEVER erase or touch columns 24+ (Column X onwards) so user text is preserved!
        for r in range(1, ws.max_row + 1):
            for c in range(1, 24):
                cell = ws.cell(r, c)
                cell.value = None
                cell.fill = PatternFill(fill_type=None)
    else:
        ws = wb.create_sheet(title=SHEET_NAME)

    # Styles
    navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    light_green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Soft light green
    green_font = Font(name="Segoe UI", size=10, bold=True, color="276A3C")
    orange_font = Font(name="Segoe UI", size=10, bold=True, color="B25900")
    red_font = Font(name="Segoe UI", size=10, bold=True, color="C00000")
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    regular_font = Font(name="Segoe UI", size=10)
    bold_font = Font(name="Segoe UI", size=10, bold=True)
    title_font = Font(name="Segoe UI", size=15, bold=True, color="1F4E78")
    subtitle_font = Font(name="Segoe UI", size=10, italic=True, color="595959")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = "D5"

    # Title & Subtitle (placed in columns 1 to 23)
    ws.cell(1, 1, "BBMP OFFLINE PANELS MONITORING (> 7 DAYS)").font = title_font
    ws.cell(2, 1, (
        f"Total Monitored Panels: {len(updated)} | "
        f"Communicated Today (Pending 0 Days): {comm_today_count} (Light Green Highlighted) | "
        f"Offline (> 7 Days): {len(updated) - comm_today_count} | "
        f"Updated: {now_dt.strftime('%Y-%m-%d %H:%M:%S')}"
    )).font = subtitle_font

    # Headers (Row 4, Columns 1 to 23)
    header_row = 4
    ws.row_dimensions[header_row].height = 28
    for c_idx, (col_name, col_width, _) in enumerate(COLUMNS, 1):
        cell = ws.cell(header_row, c_idx, col_name)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        col_letter = openpyxl.utils.get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = col_width

    # Data Rows (Row 5 to 149, Columns 1 to 23)
    for idx, p in enumerate(updated, 1):
        curr_row = header_row + idx
        ws.row_dimensions[curr_row].height = 20
        is_green = p.get("highlight_green", False)
        sig_qual = p.get("signal_quality", "N/A")

        row_values = [
            idx,
            p["panel_name"],
            p["panel_label"],
            p["location"],
            p["zone"],
            p["ward"],
            p["region"],
            p["lat"],
            p["lng"],
            p["fw_ver"],
            p["gsm_ver"],
            p["status"],
            p["pdays"],
            p["comm_date"],
            p["dbm"],
            p["rssi"],
            p["install_state"],
            p["install_date"],
            p["created_date"],
            p["meter_sl_no"],
            p["sim_network"],
            p["imei"],
            sig_qual                                                      # Col 23 (W)
        ]

        # Write only to Columns 1 through 23 (A to W)
        for c_idx, val in enumerate(row_values, 1):
            cell = ws.cell(curr_row, c_idx, val)
            cell.border = thin_border
            align_type = COLUMNS[c_idx - 1][2]
            col_name = COLUMNS[c_idx - 1][0]

            if col_name == "Location / Address":
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal=align_type, vertical="center")

            # Highlighting: ONLY when Communication Pending Days == 0 (latest communicate / today)
            if is_green:
                cell.fill = light_green_fill
                if col_name == "Status":
                    cell.font = green_font
                elif col_name == "Signal Strength (Good/Medium/Low)":
                    cell.font = green_font if sig_qual == "Good" else (orange_font if sig_qual == "Medium" else red_font)
                elif col_name in ["Panel Name", "Communication Pending Days"]:
                    cell.font = bold_font
                else:
                    cell.font = regular_font
            else:
                if idx % 2 == 0:
                    cell.fill = zebra_fill
                if col_name == "Status":
                    cell.font = red_font
                elif col_name == "Signal Strength (Good/Medium/Low)":
                    cell.font = green_font if sig_qual == "Good" else (orange_font if sig_qual == "Medium" else red_font)
                else:
                    cell.font = regular_font

        # Note: Columns 24+ (Column X onwards) are completely untouched!

    # AutoFilter covers up to max_column (includes user columns after X if present)
    final_max_col = max(len(COLUMNS), ws.max_column)
    ws.auto_filter.ref = f"A{header_row}:{openpyxl.utils.get_column_letter(final_max_col)}{header_row + len(updated)}"
    return True


def update_offline_7_days_standalone():
    """Standalone runner for update_offline_7_days.py."""
    print("=" * 60)
    print("  Updating 'Offline > 7 Days' Sheet (145 Panels)")
    print("  * Column W: Signal Strength (Good/Medium/Low)")
    print("  * Preserving Column X and beyond intact")
    print("=" * 60)

    target_file = EXCEL_FILE
    if not os.path.exists(target_file) and os.path.exists(FALLBACK_FILE):
        target_file = FALLBACK_FILE

    if not os.path.exists(target_file):
        print(f"[ERROR] Excel file not found: {EXCEL_FILE}")
        return False

    wb = openpyxl.load_workbook(target_file)
    success = update_offline_sheet_in_workbook(wb)
    if success:
        try:
            wb.save(target_file)
            print(f"\nSUCCESS! File saved to: {target_file}")
        except PermissionError:
            wb.save(FALLBACK_FILE)
            print(f"\nNOTICE: '{target_file}' is currently open in Excel.")
            print(f"SUCCESS! Saved to fallback file: {FALLBACK_FILE}")
    print("=" * 60)
    return success


if __name__ == "__main__":
    update_offline_7_days_standalone()
