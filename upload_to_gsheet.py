"""
upload_to_gsheet.py
===================
Uploads the BBMP Panels Firmware data from the generated Excel file
to the target Google Spreadsheet.

Target Spreadsheet:
  https://docs.google.com/spreadsheets/d/1fmALRdb7C9itdenS6t-8sBsFw8volLPko_RgcIrhTwg

Sheets Handled:
  1. "BBMP Panel FW & Locations"  -> Full panel telemetry & firmware data
  2. "Firmware & Zone Dashboard"  -> Summary statistics
  3. "Offline > 7 Days"           -> 145 monitored panels with live updates

PROTECTION RULE FOR "Offline > 7 Days":
  - Only Columns A to W (Columns 1 to 23) are updated and styled.
  - Columns X (24) and BEYOND ARE NEVER CLEARED, OVERWRITTEN, OR FORMATTED!
  - User notes, remarks, extra columns typed in Google Sheets remain 100% intact.
  - Light Green highlighting is applied in Google Sheets ONLY for panels where
    Communication Pending Days == 0 (latest communicate / today).
"""

import os
import sys
import json
import time
import openpyxl
import gspread
from datetime import datetime

# Ensure utf-8 output encoding on Windows console
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ─── CONFIGURATION ────────────────────────────────────────────────────────────
SPREADSHEET_ID   = "1fmALRdb7C9itdenS6t-8sBsFw8volLPko_RgcIrhTwg"
EXCEL_FILE       = "BBMP_Panels_Firmware_Versions.xlsx"
FALLBACK_EXCEL   = "BBMP_Panels_Firmware_Versions_With_Locations.xlsx"

MAIN_SHEET_NAME    = "BBMP Panel FW & Locations"
DASH_SHEET_NAME    = "Firmware & Zone Dashboard"
OFFLINE_SHEET_NAME = "Offline > 7 Days"

TOKEN_FILE       = "token_gspread.json"
OAUTH_CREDS_FILES = ["credentials.json", "client_secret.json", "oauth_credentials.json"]
SERVICE_ACCOUNT_FILES = ["service_account.json", "sa_credentials.json", "bbmo-513109-8352846b357a.json"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
# ──────────────────────────────────────────────────────────────────────────────


def get_gspread_client():
    """
    Attempts to authenticate with Google Sheets using:
    1. Service Account JSON file (if present)
    2. Cached OAuth token
    3. OAuth client credentials file (browser popup)
    """
    for sa_name in SERVICE_ACCOUNT_FILES:
        if os.path.exists(sa_name):
            try:
                with open(sa_name, "r") as f:
                    data = json.load(f)
                    if data.get("type") == "service_account":
                        print(f"   [Auth] Using Service Account: {sa_name}")
                        print(f"   [Auth] Service Account Email: {data.get('client_email')}")
                        return gspread.service_account(filename=sa_name, scopes=SCOPES)
            except Exception as e:
                print(f"   [Auth Warning] Failed to load {sa_name}: {e}")

    for fname in os.listdir("."):
        if fname.endswith(".json") and fname not in [TOKEN_FILE, "resolved_locations.json", "offline_145_panels.json"]:
            try:
                with open(fname, "r") as f:
                    data = json.load(f)
                    if data.get("type") == "service_account":
                        print(f"   [Auth] Detected Service Account in: {fname}")
                        print(f"   [Auth] Email: {data.get('client_email')}")
                        return gspread.service_account(filename=fname, scopes=SCOPES)
            except Exception:
                pass

    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            print("   [Auth] Using cached OAuth credentials (token_gspread.json)")
            return gspread.authorize(creds)
        except Exception as e:
            print(f"   [Auth] Cached token refresh failed: {e}")
            creds = None

    oauth_file = None
    for fname in OAUTH_CREDS_FILES:
        if os.path.exists(fname):
            oauth_file = fname
            break

    if oauth_file:
        print(f"   [Auth] Opening browser for Google sign-in using: {oauth_file}...")
        flow = InstalledAppFlow.from_client_secrets_file(oauth_file, SCOPES)
        creds = flow.run_local_server(port=0, open_browser=True)
        with open(TOKEN_FILE, "w") as tf:
            tf.write(creds.to_json())
        print(f"   [Auth] Credentials saved to {TOKEN_FILE} for future automatic runs.")
        return gspread.authorize(creds)

    return None


def read_sheet_to_rows(ws, max_cols=None):
    """Read all cells in an openpyxl worksheet as a list of lists of strings."""
    rows = []
    for row in ws.iter_rows(values_only=True):
        if max_cols:
            row = row[:max_cols]
        rows.append([str(v) if v is not None else "" for v in row])
    return rows


def get_or_create_worksheet(sh, title, rows=2000, cols=30):
    """Retrieve existing worksheet or create a new one."""
    try:
        ws = sh.worksheet(title)
        return ws
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)


def upload_worksheet_data(sh, sheet_title, rows_data, chunk_size=500):
    """Batch upload rows to Google Sheets worksheet (used for main data and summary)."""
    if not rows_data:
        return None

    num_rows = len(rows_data)
    num_cols = max(len(r) for r in rows_data)

    ws = get_or_create_worksheet(sh, sheet_title, rows=max(num_rows + 20, 100), cols=max(num_cols + 2, 26))

    print(f"   Clearing existing data in '{sheet_title}'...")
    ws.clear()
    time.sleep(1)

    print(f"   Uploading {num_rows} rows x {num_cols} columns to '{sheet_title}'...")
    for start in range(0, num_rows, chunk_size):
        chunk = rows_data[start : start + chunk_size]
        end = min(start + chunk_size, num_rows)
        ws.append_rows(chunk, value_input_option="USER_ENTERED")
        print(f"     Uploaded rows {start+1} to {end} of {num_rows}")
        if end < num_rows:
            time.sleep(0.8)

    print(f"   [OK] Sheet '{sheet_title}' uploaded successfully.")
    return ws


def upload_offline_sheet_preserve_x(sh, sheet_title, rows_data):
    """
    Updates 'Offline > 7 Days' in Google Sheets:
    - NEVER clears the worksheet (ws.clear() is NEVER called).
    - Writes ONLY to Columns A through W (range A1:W{num_rows}).
    - Columns X (24) and beyond are 100% untouched and preserved!
    - Applies formatting:
        * Freeze top 4 rows and first 3 columns
        * Navy Blue bold header on Row 4 (Cols A-W)
        * Soft Light Green (#E2EFDA) on entire row (Cols A-W) ONLY if Pending Days == 0
        * Bold Red status for Offline panels (> 0 pending days)
        * Bold Green/Orange/Red text for Column W (Signal Strength)
        * Auto-column sizing for Columns A-W
    """
    if not rows_data:
        return None

    num_rows = len(rows_data)
    # Exactly 23 columns (Columns A through W)
    num_cols = 23
    trimmed_rows = [r[:num_cols] for r in rows_data]

    # Ensure worksheet exists and has at least 30 columns so user has space for Column X and beyond
    try:
        ws = sh.worksheet(sheet_title)
        if ws.col_count < 30:
            ws.add_cols(30 - ws.col_count)
        if ws.row_count < num_rows + 10:
            ws.add_rows(num_rows + 10 - ws.row_count)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_title, rows=max(num_rows + 20, 160), cols=30)

    print(f"   Updating range A1:W{num_rows} in '{sheet_title}' (Columns X+ preserved)...")
    target_range = f"A1:W{num_rows}"

    # Update ONLY range A1:W{num_rows}
    ws.update(values=trimmed_rows, range_name=target_range, value_input_option="USER_ENTERED")
    print(f"   [OK] Columns A-W updated. User columns (X and beyond) remain intact!")

    # Format 'Offline > 7 Days' in Google Sheets
    print(f"   Applying formatting and green highlights (Pending Days = 0) to '{sheet_title}'...")
    try:
        ws.freeze(rows=4, cols=3)

        requests_list = [
            # 1. Navy Blue Header (Row 4, Columns A to W / indices 0 to 23)
            {
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": 3,
                        "endRowIndex": 4,
                        "startColumnIndex": 0,
                        "endColumnIndex": num_cols,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.122, "green": 0.306, "blue": 0.471},
                            "textFormat": {
                                "bold": True,
                                "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                                "fontSize": 10
                            },
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
                }
            },
            # 2. Title Block (Row 1, Column A)
            {
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "bold": True,
                                "fontSize": 14,
                                "foregroundColor": {"red": 0.122, "green": 0.306, "blue": 0.471}
                            }
                        }
                    },
                    "fields": "userEnteredFormat(textFormat)",
                }
            },
        ]

        # 3. Format Each Data Row (Row 5 to num_rows / indices 4 to num_rows)
        # ONLY across Columns A to W (startColumnIndex: 0, endColumnIndex: 23)!
        for r_idx in range(4, num_rows):
            row_data = trimmed_rows[r_idx]
            pdays_val = row_data[12] if len(row_data) > 12 else ""
            status_val = row_data[11] if len(row_data) > 11 else ""
            sig_qual = row_data[22] if len(row_data) > 22 else ""

            is_pending_zero = (str(pdays_val).strip() in ["0", "0.0"])

            if is_pending_zero:
                # LIGHT GREEN HIGHLIGHT across Columns A through W
                requests_list.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": r_idx,
                            "endRowIndex": r_idx + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": num_cols,  # Cols A-W only!
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.886, "green": 0.937, "blue": 0.855}  # #E2EFDA Soft Light Green
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })
                # Bold Green Status (Col L / index 11)
                requests_list.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": r_idx,
                            "endRowIndex": r_idx + 1,
                            "startColumnIndex": 11,
                            "endColumnIndex": 12,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {
                                    "bold": True,
                                    "foregroundColor": {"red": 0.153, "green": 0.416, "blue": 0.235}
                                }
                            }
                        },
                        "fields": "userEnteredFormat.textFormat"
                    }
                })
            else:
                # Normal zebra / white for offline rows
                bg_color = {"red": 0.976, "green": 0.980, "blue": 0.984} if (r_idx % 2 == 0) else {"red": 1.0, "green": 1.0, "blue": 1.0}
                requests_list.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": r_idx,
                            "endRowIndex": r_idx + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": num_cols,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": bg_color
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })
                # Bold Red Status (Col L / index 11)
                requests_list.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": r_idx,
                            "endRowIndex": r_idx + 1,
                            "startColumnIndex": 11,
                            "endColumnIndex": 12,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {
                                    "bold": True,
                                    "foregroundColor": {"red": 0.753, "green": 0.0, "blue": 0.0}
                                }
                            }
                        },
                        "fields": "userEnteredFormat.textFormat"
                    }
                })

            # Column W (Signal Strength Good/Medium/Low, index 22) text color
            if sig_qual == "Good":
                fg_color = {"red": 0.153, "green": 0.416, "blue": 0.235} # Green
            elif sig_qual == "Medium":
                fg_color = {"red": 0.698, "green": 0.349, "blue": 0.0}   # Orange
            else:
                fg_color = {"red": 0.753, "green": 0.0, "blue": 0.0}     # Red

            requests_list.append({
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": r_idx,
                        "endRowIndex": r_idx + 1,
                        "startColumnIndex": 22,
                        "endColumnIndex": 23,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "bold": True,
                                "foregroundColor": fg_color
                            }
                        }
                    },
                    "fields": "userEnteredFormat.textFormat"
                }
            })

        # 4. Auto-Resize Columns A to W (indices 0 to 23)
        requests_list.append({
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": ws.id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": num_cols,
                }
            }
        })

        # Batch update formatting in chunks of 100 requests to avoid quota limit
        for i in range(0, len(requests_list), 100):
            chunk = requests_list[i : i + 100]
            sh.batch_update({"requests": chunk})
            time.sleep(0.5)

        print("   [OK] Applied formatting & light green highlights for communicated panels!")
    except Exception as e:
        print(f"   [Notice] Offline sheet formatting details skipped: {e}")

    return ws


def apply_main_sheet_formatting(sh, ws_title, num_cols):
    """Apply styling: frozen panes, navy bold header, auto column resize to main sheet."""
    try:
        ws = sh.worksheet(ws_title)
        ws.freeze(rows=4, cols=3)

        body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 3,
                            "endRowIndex": 4,
                            "startColumnIndex": 0,
                            "endColumnIndex": num_cols,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.122, "green": 0.306, "blue": 0.471},
                                "textFormat": {
                                    "bold": True,
                                    "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                                    "fontSize": 10
                                },
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
                    }
                },
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": {
                                    "bold": True,
                                    "fontSize": 14,
                                    "foregroundColor": {"red": 0.122, "green": 0.306, "blue": 0.471}
                                }
                            }
                        },
                        "fields": "userEnteredFormat(textFormat)",
                    }
                },
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": ws.id,
                            "dimension": "COLUMNS",
                            "startIndex": 0,
                            "endIndex": num_cols,
                        }
                    }
                }
            ]
        }
        sh.batch_update(body)
        print("   [OK] Applied formatting to main sheet: Navy header, frozen panes, auto widths.")
    except Exception as e:
        print(f"   [Notice] Main sheet formatting details skipped: {e}")


def upload_excel_to_gsheet(excel_path=None):
    """
    Main entry point to upload the generated Excel to Google Sheets.
    Can be called standalone or imported by generate_bbmp_fw_excel.py.
    """
    print("\n" + "=" * 60)
    print("  BBMP Firmware -> Google Sheets Sync")
    print("=" * 60)

    # 1. Determine Excel file path
    if not excel_path:
        if os.path.exists(EXCEL_FILE):
            excel_path = EXCEL_FILE
        elif os.path.exists(FALLBACK_EXCEL):
            excel_path = FALLBACK_EXCEL
        else:
            print(f"\n[ERROR] No Excel file found to upload ({EXCEL_FILE}).")
            return False

    print(f"\n[1/3] Reading data from Excel: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    main_ws_name = MAIN_SHEET_NAME if MAIN_SHEET_NAME in wb.sheetnames else wb.sheetnames[0]
    dash_ws_name = DASH_SHEET_NAME if DASH_SHEET_NAME in wb.sheetnames else None
    offline_ws_name = OFFLINE_SHEET_NAME if OFFLINE_SHEET_NAME in wb.sheetnames else None

    main_rows = read_sheet_to_rows(wb[main_ws_name])
    dash_rows = read_sheet_to_rows(wb[dash_ws_name]) if dash_ws_name else []
    # Read Columns A through W (23 columns) for Offline sheet
    offline_rows = read_sheet_to_rows(wb[offline_ws_name], max_cols=23) if offline_ws_name else []

    print(f"   Main sheet '{main_ws_name}': {len(main_rows)} rows x {len(main_rows[0]) if main_rows else 0} columns")
    if dash_ws_name:
        print(f"   Dashboard sheet '{dash_ws_name}': {len(dash_rows)} rows x {len(dash_rows[0]) if dash_rows else 0} columns")
    if offline_ws_name:
        print(f"   Offline sheet '{offline_ws_name}': {len(offline_rows)} rows x {len(offline_rows[0]) if offline_rows else 0} columns")

    # 2. Authenticate
    print("\n[2/3] Authenticating with Google...")
    gc = get_gspread_client()
    if not gc:
        print("\n" + "!" * 60)
        print("  GOOGLE AUTHENTICATION NOT CONFIGURED YET!")
        print("!" * 60)
        return False

    # 3. Open Spreadsheet & Upload
    print(f"\n[3/3] Opening Google Spreadsheet (ID: {SPREADSHEET_ID})...")
    try:
        sh = gc.open_by_key(SPREADSHEET_ID)
        print(f"   Connected to: '{sh.title}'")
    except Exception as e:
        actual_err = getattr(e, "__cause__", None) or e
        print(f"\n[ERROR] Unable to open spreadsheet {SPREADSHEET_ID}: {actual_err}")
        return False

    # Upload Main Sheet
    upload_worksheet_data(sh, MAIN_SHEET_NAME, main_rows)

    # Upload Dashboard Sheet
    if dash_rows:
        upload_worksheet_data(sh, DASH_SHEET_NAME, dash_rows)

    # Upload Offline > 7 Days Sheet (PRESERVING COLUMN X AND BEYOND)
    if offline_rows:
        upload_offline_sheet_preserve_x(sh, OFFLINE_SHEET_NAME, offline_rows)

    # Apply Formatting to Main Sheet
    if main_rows:
        apply_main_sheet_formatting(sh, MAIN_SHEET_NAME, len(main_rows[0]))

    print("\n" + "=" * 60)
    print("  [SUCCESS] GOOGLE SHEET SYNC COMPLETE!")
    print(f"  Live Sheet: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
    print("=" * 60)
    return True


if __name__ == "__main__":
    success = upload_excel_to_gsheet()
    if not success:
        sys.exit(1)
