"""
Excel LIVE Watcher (bina save kiye)
-----------------------------------
Ye version file-save ka wait nahi karta. Ye seedha already-open Excel
application se (COM automation ke zariye) live memory se data padhta hai —
bilkul waise hi jaise tumhari aankhein Excel screen par dekhti hain.

ZAROORI: Excel me tumhari file already khuli honi chahiye jab ye script chalao.

Install karo pehle:
    pip install xlwings pywin32
"""

import time
import json
from datetime import datetime

import requests
import xlwings as xw

# ============ CONFIG — YAHAN APNI VALUES DAALO ============
WORKBOOK_PATH = r"Z:\PANKAJ\pankaj_codes\Strategy\ramesh12.xlsx"  # Excel me jo file khuli hai, uska pura path
SHEET_NAME = None          # None = active sheet, ya "Sheet1" jaisa naam do
SERVER_URL = "SERVER_URL = "https://exceltrade.onrender.com/update""
API_KEY = "mysecretkey123"     # server.py wali API_KEY se match honi chahiye
POLL_SECONDS = 1.5         # kitni baar (seconds me) check karega ki data badla ya nahi
# ============================================================


def find_open_workbook():
    """Already khuli Excel application me se ye workbook dhundta hai."""
    target = WORKBOOK_PATH.lower().replace("/", "\\")
    for app in xw.apps:
        for book in app.books:
            if book.fullname.lower().replace("/", "\\") == target:
                return book
    raise RuntimeError(
        f"'{WORKBOOK_PATH}' Excel me khuli hui nahi mili. "
        "Pehle is file ko Excel me khol lo, phir ye script chalao."
    )


def read_live_data(wb):
    """Live in-memory data padhta hai — koi save zaroori nahi."""
    ws = wb.sheets[SHEET_NAME] if SHEET_NAME else wb.sheets.active
    used = ws.used_range
    values = used.value  # list of lists (live values, formulas ka calculated result)

    if not values or len(values) < 1:
        return {"sheet_name": ws.name, "headers": [], "rows": [], "updated_at": None}

    # Agar sirf ek row hai to values ek flat list hogi, use normalize karo
    if not isinstance(values[0], (list, tuple)):
        values = [values]

    headers = ["" if h is None else str(h) for h in values[0]]
    rows = [["" if c is None else c for c in row] for row in values[1:]]

    return {
        "sheet_name": ws.name,
        "headers": headers,
        "rows": rows,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def send_to_server(data):
    try:
        resp = requests.post(
            SERVER_URL,
            headers={"X-API-Key": API_KEY, "Content-Type": "application/json"},
            data=json.dumps(data, default=str),
            timeout=10,
        )
        if resp.status_code == 200:
            print(f"[OK] Sent {len(data['rows'])} rows @ {data['updated_at']}")
        else:
            print(f"[ERROR] Server responded {resp.status_code}: {resp.text}")
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Could not reach server: {e}")


def main():
    print(f"Connecting to live Excel: {WORKBOOK_PATH}")
    wb = find_open_workbook()
    print(f"Connected! Watching sheet live (har {POLL_SECONDS}s check karega)...")
    print(f"Sending updates to: {SERVER_URL}")
    print("(Excel save karne ki zaroorat NAHI hai)\n")

    last_snapshot = None

    while True:
        try:
            data = read_live_data(wb)
            snapshot = json.dumps(data, default=str)

            # Sirf tabhi bhejo jab data waaqai badla ho (bekar traffic na ho)
            if snapshot != last_snapshot and data["headers"]:
                send_to_server(data)
                last_snapshot = snapshot

        except Exception as e:
            print(f"[ERROR] {e}")
            # Agar Excel band ho gaya ho to dobara connect try karo
            try:
                wb = find_open_workbook()
            except Exception:
                pass

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nRuka diya gaya.")
