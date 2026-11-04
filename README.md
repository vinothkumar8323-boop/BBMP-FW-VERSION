# BBMP Panels Firmware & Location Tracking System

Automated reporting, telemetry extraction, and live dashboard sync for BBMP (Bruhat Bengaluru Mahanagara Palike) smart street light panels using the **Schnell IoT Platform** and **Google Sheets API**.

---

## 📌 Features

- **Automated Data Extraction**: Queries Schnell IoT ThingsBoard API in parallel (multi-threaded) across 5,600+ BBMP panels.
- **Firmware & Telemetry Auditing**: Extracts controller firmware version (`ver`), GSM firmware version (`gsmVersion`), meter serial numbers, SIM/IMEI, and network providers.
- **Location & Geotagging**: Resolves physical panel addresses (Zone, Ward, Region) and generates direct Google Maps coordinates.
- **Offline Panel Monitoring**: Special monitoring sheet for panels offline > 7 days with live attributes, signal strength (Good / Medium / Low), RSSI, and dBm values.
- **Smart Highlighting**: Automatically highlights panels in soft light green when they communicate today (0 pending days).
- **Google Sheets Live Sync**: Automatically updates existing Google Spreadsheets with frozen panes, custom styling, and protection for user notes (Column X and beyond).

---

## 📂 Project Structure

| File | Description |
|---|---|
| `generate_bbmp_fw_excel.py` | Full pipeline: fetches all devices from Schnell IoT, builds formatted Excel report, updates offline monitoring, and syncs to Google Sheets. |
| `update_offline_7_days.py` | Live updater for the 145 monitored offline panels sheet; calculates pending days, RSSI/dBm signal levels, and preserves user remarks after Column X. |
| `upload_to_gsheet.py` | Standalone Google Sheets API synchronization utility using Service Account or OAuth 2.0. |
| `apply_locations_to_excel.py` | Applies geocoded location cache and addresses to existing Excel workbooks. |
| `verify_excel.py` | Audits generated Excel files for completeness, zones, and communication statistics. |
| `offline_145_panels.json` | Baseline dataset for the 145 monitored offline panels. |
| `resolved_locations.json` | Cache of pre-resolved addresses and locations. |

---

## 🚀 Setup & Usage

### 1. Requirements
Install required Python libraries:
```bash
pip install requests pandas openpyxl gspread google-auth-oauthlib
```

### 2. Configuration
- Place your Google Cloud service account key as `service_account.json` in the project root.
- Ensure the service account email has **Editor** permissions on your target Google Sheet.

### 3. Running the Pipeline
- **Run the full extraction & Google Sheet sync:**
  ```bash
  python generate_bbmp_fw_excel.py
  ```

- **Quickly refresh only the monitored Offline panels sheet:**
  ```bash
  python update_offline_7_days.py
  ```

- **Sync existing Excel to Google Sheets:**
  ```bash
  python upload_to_gsheet.py
  ```
