# CoreBilling – Instrument Usage Tracker

CoreBilling is a lightweight Windows tool that calculates user login durations on instrument PCs using Windows Security Event Logs and generates a CSV report for billing.

It is designed to run as a standalone executable with a simple configuration file that can be maintained by non-technical staff.

---

## 📌 Features

* Tracks user login/logoff sessions from Windows Event Logs
* Filters out system/service accounts
* Calculates actual usage time
* Applies minimum billing rules (e.g., overnight usage)
* Maps computer names to instrument names via a config file
* Outputs a clean CSV report for billing

---

## 📁 File Structure

```
CoreBilling/
│
├── CoreBilling.exe        # Generated executable
├── config.csv            # Editable configuration file
├── README.md
```

---

## ⚙️ Configuration

Create a `config.csv` file in the same folder as the executable.

### Example:

```
ComputerName,InstrumentName,MinHours
LAB-PC-01,LCMS,2
LAB-PC-02,Flow Cytometer,1
```

### Columns

| Column         | Description                                        |
| -------------- | -------------------------------------------------- |
| ComputerName   | Exact Windows machine name                         |
| InstrumentName | Friendly name used in reports                      |
| MinHours       | Minimum billed hours (used for overnight sessions) |

---

## ▶️ Running the Tool

### Manual Run

Double-click:

```
CoreBilling.exe
```

Or run via command line:

```
CoreBilling.exe
```

---

### Output

A CSV file will be generated in the same folder:

```
Core_Billing_<Month_Year>.csv
```

Example:

```
Core_Billing_March_2026.csv
```

---

## 🗓️ Recommended Usage

Run **once per month** to generate billing reports.

### Suggested Automation (optional)

Use Windows Task Scheduler:

* Trigger: Monthly
* Action: Run `CoreBilling.exe`
* Enable: “Run with highest privileges”

---

## 🔐 Permissions

The tool reads the Windows Security Event Log.

### Required:

* Run as Administrator
  **OR**
* User must be in **Event Log Readers** group

---

## 🧪 Testing

For testing:

* Temporarily reduce the lookback window in the script (e.g., 2 days instead of 31)
* Run manually and verify output

---

## ⚠️ Notes

* If a computer is not listed in `config.csv`, the computer name will be used as the instrument name
* Missing or invalid `MinHours` defaults to `1.0`
* Open sessions (no logoff event) are included as `"OPEN"`

---

## 🛠️ Building the Executable

### Prerequisites

* Python 3.10+
* PyInstaller

Install PyInstaller:

```
pip install pyinstaller
```

---

### Build Command

From the script directory:

```
pyinstaller --onefile --noconsole --name CoreBilling your_script.py
```

### Output

The executable will be located in:

```
dist/CoreBilling.exe
```

---

### Development Tip

For debugging, build **without** `--noconsole`:

```
pyinstaller --onefile --name CoreBilling your_script.py
```

---

## 🚀 Deployment

1. Copy the following to the instrument PC:

```
CoreBilling.exe
config.csv
```

2. Place in a folder such as:

```
C:\CoreBilling\
```

3. Run manually or schedule monthly

---

## 📬 Support / Maintenance

* Update `config.csv` to modify instrument names or billing minimums
* No code changes required for normal operation

---

## 🧠 Future Enhancements (Optional)

* Daily automated runs with aggregation
* Real-time background logging
* Centralized reporting dashboard
* Error logging to file

---
