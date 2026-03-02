# Work Schedule → Google Calendar Automation

A Python automation tool that transforms structured workforce scheduling data from CSV format into Google Calendar events using the Google Calendar API.

Built as a practical automation project demonstrating data parsing, API integration, timezone-aware datetime handling, and idempotent execution.

---

## 📌 Project Overview

This tool reads a master schedule CSV file, extracts an individual employee’s work shifts, and synchronizes them directly to Google Calendar.

It is designed to be:
- **Safe to re-run** (duplicate events are prevented)
- **Timezone-aware** (DST-safe)
- **Portable across machines**
- **Secure** (credentials excluded from version control)

---

## ✨ Features

- Parses structured CSV schedules using **pandas**
- Detects work shifts via **regex-based time range extraction**
- Handles off-days (rest days, holidays, sick leave, blanks)
- Converts tabular schedule data into calendar events
- Integrates with **Google Calendar API** via OAuth2
- Prevents duplicate event creation through deterministic keys
- Supports reusable, isolated environments via **virtualenv**
- Runs via command line or one-click Windows batch file

---

## 📂 Expected CSV Format

- Actual column headers begin on **row 3**
- Employee names appear in column `AGENTS` (fallback: second column)
- Date columns formatted as: `Feb-2, Mar-15, Nov-30`
- Work shifts contain time ranges such as: `09:00 - 17:00`. Any non-matching value (e.g. `R`, `Sick`, empty cell) is treated as an off-day

---

## 🛠️ Tech Stack

- **Python 3**
- **pandas**
- **Google Calendar API**
- **OAuth2 authentication**
- **Regex parsing**
- **Timezone-aware datetime handling**
- **Virtual environments (venv)**

---

## 🔐 Authentication & Security

This project uses OAuth2 (Desktop App flow).

Local-only files (excluded via `.gitignore`):
- `credentials.json`
- `token.json`
- `local_config.py`
- Real schedule CSV files

Secrets and personal identifiers are never committed to the repository.

---

## ▶️ Running the Tool

### 1. Install dependencies
```bash
pip install -r requirements.txt