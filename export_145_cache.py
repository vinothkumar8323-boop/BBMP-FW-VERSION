import openpyxl
import json

wb = openpyxl.load_workbook('BBMP_Panels_Firmware_Versions.xlsx', data_only=True)
ws = wb['Offline > 7 Days']

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

with open('offline_145_panels.json', 'w') as f:
    json.dump(panels, f, indent=2)

print(f"Saved {len(panels)} panels to offline_145_panels.json")

