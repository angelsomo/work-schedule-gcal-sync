# Work Schedule → Google Calendar Sync

A Python automation tool that extracts an employee's monthly work schedule from a master CSV file and syncs it directly to Google Calendar.

---

## 🚀 What It Does

- Reads a master schedule CSV (headers on row 3)
- Extracts a specific employee’s shifts
- Detects work shifts via time-range pattern (e.g. `09:00 - 17:00`)
- Logs off-days (R, Sick, blank, etc.)
- Pushes events directly to Google Calendar via API
- Prevents duplicate event creation
- Handles timezone correctly (Europe/Athens, DST-aware)

---

## 📂 Expected CSV Format

- Actual headers start on **row 3**
- Employee names are in column `AGENTS` (or second column fallback)
- Date columns formatted like: `Feb-2`, `Mar-15`, `Nov-30`
- Work shifts contain time ranges like: 09:00 - 17:00. Everything else is treated as an off-day

---

## 🛠 Requirements

Python 3.9+

Install dependencies:

```bash
pip install -r requirements.txt
```