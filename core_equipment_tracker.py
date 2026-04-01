import win32evtlog
import pandas as pd
import os
from datetime import datetime, timedelta
import sys

# ---------------------------
# Config
# ---------------------------
import sys

if getattr(sys, 'frozen', False):
    # Running as EXE
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Running as script
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, 'config.csv')
os.makedirs(BASE_DIR, exist_ok=True)

IGNORE_ACCOUNTS = ['SYSTEM', 'LOCAL SERVICE', 'NETWORK SERVICE', 'ANONYMOUS LOGON']
LOGON_TYPES_TO_BILL = ['2']  # Add '10' if you want RDP sessions


# ---------------------------
# Load Config
# ---------------------------
def load_config():
    """Loads instrument mapping and minimum billing hours."""
    if not os.path.exists(CONFIG_FILE):
        print(f"Warning: {CONFIG_FILE} not found. Using defaults.")
        return {}, {}

    df = pd.read_csv(CONFIG_FILE)

    # Normalize column names
    df.columns = df.columns.str.strip()

    required_cols = {'ComputerName', 'InstrumentName', 'MinHours'}
    if not required_cols.issubset(df.columns):
        raise ValueError(f"{CONFIG_FILE} must contain columns: {required_cols}")

    # Clean values
    df['ComputerName'] = df['ComputerName'].astype(str).str.strip().str.upper()
    df['InstrumentName'] = df['InstrumentName'].astype(str).str.strip()
    df['MinHours'] = pd.to_numeric(df['MinHours'], errors='coerce').fillna(1.0)

    instrument_map = dict(zip(df['ComputerName'], df['InstrumentName']))
    min_map = dict(zip(df['ComputerName'], df['MinHours']))

    return instrument_map, min_map


# ---------------------------
# User Filter
# ---------------------------
def is_human_local_user(user, l_type):
    """Ensure it's a real interactive user."""
    if l_type not in LOGON_TYPES_TO_BILL:
        return False

    if user.upper() in IGNORE_ACCOUNTS or user.endswith('$'):
        return False

    return True


# ---------------------------
# Main Logic
# ---------------------------
def get_billing_report(days_back=31):
    instrument_map, instrument_mins = load_config()

    server = 'localhost'
    hand = win32evtlog.OpenEventLog(server, 'Security')
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ

    start_time = datetime.now() - timedelta(days=days_back)

    sessions = {}
    final_data = []

    try:
        while True:
            events = win32evtlog.ReadEventLog(hand, flags, 0)
            if not events:
                break

            for event in events:
                if event.TimeCreated < start_time:
                    break

                e_id = event.EventID & 0xFFFF
                data = event.StringInserts

                if not data:
                    continue

                # ---------------------------
                # LOGIN EVENTS
                # ---------------------------
                if e_id in [4624, 4672]:
                    if len(data) < 9:
                        continue

                    try:
                        if e_id == 4624:
                            logon_id = data[7]
                            user = data[5]
                            l_type = data[8]
                        else:
                            logon_id = data[1]
                            user = data[5] if len(data) > 5 else "UNKNOWN"
                            l_type = "2"
                    except Exception:
                        continue

                    if is_human_local_user(user, l_type):
                        sessions[logon_id] = {
                            'User': user,
                            'Start': event.TimeCreated,
                            'PC': event.ComputerName.upper()
                        }

                # ---------------------------
                # LOGOFF EVENT
                # ---------------------------
                elif e_id == 4647:
                    if len(data) < 4:
                        continue

                    logon_id = data[3]

                    if logon_id in sessions:
                        s = sessions[logon_id]
                        end_time = event.TimeCreated

                        actual_hrs = (end_time - s['Start']).total_seconds() / 3600

                        pc_name = s['PC']
                        instrument_name = instrument_map.get(pc_name, pc_name)
                        min_val = instrument_mins.get(pc_name, 1.0)

                        is_overnight = s['Start'].date() != end_time.date()
                        billed_hrs = min_val if is_overnight else actual_hrs

                        final_data.append({
                            'User': s['User'],
                            'Instrument': instrument_name,
                            'ComputerName': pc_name,
                            'Login': s['Start'].strftime('%Y-%m-%d %H:%M'),
                            'Logoff': end_time.strftime('%Y-%m-%d %H:%M'),
                            'Actual_Hours': round(actual_hrs, 2),
                            'Billed_Hours': round(billed_hrs, 2),
                            'Billing_Type': 'Overnight Min' if is_overnight else 'Standard'
                        })

                        del sessions[logon_id]

            # If we hit older events, stop outer loop too
            if event.TimeCreated < start_time:
                break

    finally:
        win32evtlog.CloseEventLog(hand)

    # ---------------------------
    # Handle Open Sessions
    # ---------------------------
    now = datetime.now()
    for logon_id, s in sessions.items():
        actual_hrs = (now - s['Start']).total_seconds() / 3600

        pc_name = s['PC']
        instrument_name = instrument_map.get(pc_name, pc_name)

        final_data.append({
            'User': s['User'],
            'Instrument': instrument_name,
            'ComputerName': pc_name,
            'Login': s['Start'].strftime('%Y-%m-%d %H:%M'),
            'Logoff': 'OPEN',
            'Actual_Hours': round(actual_hrs, 2),
            'Billed_Hours': round(actual_hrs, 2),
            'Billing_Type': 'Open Session'
        })

    return pd.DataFrame(final_data)


# ---------------------------
# Run Script
# ---------------------------
if __name__ == "__main__":
    df = get_billing_report()

    if not df.empty:
        report_date = (datetime.now() - timedelta(days=5)).strftime('%B_%Y')
        filename = os.path.join(BASE_DIR, f"Core_Billing_{report_date}.csv")

        df.to_csv(filename, index=False)
        print(f"Success! Report generated: {filename}")
    else:
        print("No billable sessions found.")