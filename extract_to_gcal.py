"""
Function: Push work shifts to Google Calendar.

- Reads master schedule CSV (headers on 3rd row => header=2)
- Finds TARGET_EMPLOYEE row
- Parses date columns like 'Feb-2'
- Extracts time ranges '09:00 - 17:00' via regex
- Treats everything else as Off Day
- Pushes events to Google Calendar with dedupe (safe to re-run)

Expected files in same folder as this script:
- credentials.json  (OAuth Desktop app credentials)
- CSVs/master_schedule.csv

Creates:
- token.json  (OAuth token cache)
"""

from __future__ import annotations

import calendar
import os
import re
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from zoneinfo import ZoneInfo


# ==============================
# USER CONFIG
# ==============================
TARGET_EMPLOYEE = "Άγγελος Σώμογλου (Θεσσαλονίκη)"
YEAR = 2026
TIMEZONE = "Europe/Athens"

CALENDAR_ID = "f5ea6d8d87e2fdcaa267ebdc5b0114f4b96425b11b4ee472548cd5a3daf17d94@group.calendar.google.com"      # or a specific calendarId
EVENT_TITLE = "Work Shift (Angel)"

DRY_RUN = False              # first run True. When correct, set to False.


# ==============================
# PATHS
# ==============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_CSV_PATH = os.path.join(BASE_DIR, "CSVs", "master_schedule.csv")
CREDENTIALS_PATH = os.path.join(BASE_DIR, "credentials.json")
TOKEN_PATH = os.path.join(BASE_DIR, "token.json")


# ==============================
# RULES
# ==============================
TIME_RANGE_RE = re.compile(r"(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})")

MONTH_MAP = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
    "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
    "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}

SCOPES = ["https://www.googleapis.com/auth/calendar.events"]


# ==============================
# Helpers
# ==============================
def detect_employee_column(df: pd.DataFrame) -> str:
    cols = list(df.columns)
    for c in cols:
        if isinstance(c, str) and c.strip().lower() == "agents":
            return c
    if len(cols) < 2:
        raise ValueError("CSV must have at least 2 columns to infer employee column.")
    return cols[1]


def looks_like_date_header(col_name: object) -> bool:
    if not isinstance(col_name, str):
        return False
    return bool(re.fullmatch(r"[A-Za-z]{3}-\d{1,2}", col_name.strip()))


def parse_date_header(header: str, year: int) -> date:
    mon_abbrev, day_str = header.strip().split("-")
    mon_abbrev = mon_abbrev.strip().title()
    if mon_abbrev not in MONTH_MAP:
        raise ValueError(f"Unknown month abbreviation '{mon_abbrev}' in header '{header}'.")
    return date(year, MONTH_MAP[mon_abbrev], int(day_str))


def month_label(d: date) -> str:
    return d.strftime("%b")


def iso_utc_z(dt_local: datetime) -> str:
    """
    Google Calendar API accepts RFC3339 timestamps.
    We provide local naive dt + timezone separately in event body, so ISO is fine here.
    """
    return dt_local.isoformat()


def event_key(title: str, start_dt: datetime, end_dt: datetime) -> str:
    """Unique key for dedupe based on title + exact start/end."""
    return f"{title}|{start_dt.date().isoformat()}|{start_dt.strftime('%H:%M')}|{end_dt.strftime('%H:%M')}"


# ==============================
# Extract shifts
# ==============================
def extract_shifts_and_offdays(
    master_csv_path: str,
    employee_name: str,
    year: int,
) -> Tuple[List[Tuple[datetime, datetime]], Dict[str, List[str]]]:
    """
    Returns:
      shifts: list of (start_dt, end_dt) as naive datetimes (local time)
      off_days_by_month: dict of month -> list of MM/DD/YYYY strings
    """
    if not os.path.exists(master_csv_path):
        raise FileNotFoundError(f"Master CSV not found at: {master_csv_path}")

    df = pd.read_csv(master_csv_path, header=2)

    emp_col = detect_employee_column(df)
    matches = df[df[emp_col].astype(str).str.strip().eq(employee_name)]
    if matches.empty:
        sample = df[emp_col].dropna().astype(str).str.strip().unique().tolist()[:20]
        raise ValueError(
            f"Employee '{employee_name}' not found in column '{emp_col}'. Sample names: {sample}"
        )

    emp_row = matches.iloc[0]
    date_cols = [c for c in df.columns if looks_like_date_header(c)]
    if not date_cols:
        raise ValueError("No date-like columns found (expected headers like 'Feb-2').")

    shifts: List[Tuple[datetime, datetime]] = []
    off_days_by_month: Dict[str, List[str]] = {}

    for col in date_cols:
        d = parse_date_header(str(col), year)
        cell = emp_row[col]
        cell_str = "" if pd.isna(cell) else str(cell).strip()

        m = TIME_RANGE_RE.search(cell_str)
        if m:
            start_24, end_24 = m.group(1), m.group(2)
            tz = ZoneInfo(TIMEZONE)
            start_dt = datetime.strptime(f"{d.isoformat()} {start_24}", "%Y-%m-%d %H:%M").replace(tzinfo=tz)
            end_dt   = datetime.strptime(f"{d.isoformat()} {end_24}", "%Y-%m-%d %H:%M").replace(tzinfo=tz)
            shifts.append((start_dt, end_dt))
        else:
            mon = month_label(d)
            off_days_by_month.setdefault(mon, []).append(d.strftime("%m/%d/%Y"))

    # sort shifts
    shifts.sort(key=lambda x: x[0])
    return shifts, off_days_by_month


def off_day_summary(off_days_by_month: Dict[str, List[str]]) -> None:
    if not off_days_by_month:
        print("\nOff Days Summary: (none detected)")
        return

    order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    idx = {m:i for i,m in enumerate(order)}
    print("\nOff Days Summary (grouped by month):")
    for mon in sorted(off_days_by_month.keys(), key=lambda m: idx.get(m, 99)):
        print(f"  {mon}: {', '.join(off_days_by_month[mon])}")

def dedupe_shifts(shifts: List[Tuple[datetime, datetime]]) -> List[Tuple[datetime, datetime]]:
    seen = set()
    unique = []
    for start_dt, end_dt in shifts:
        k = (start_dt.isoformat(), end_dt.isoformat())
        if k in seen:
            continue
        seen.add(k)
        unique.append((start_dt, end_dt))
    return unique
# ==============================
# Google Calendar API
# ==============================
def get_calendar_service():
    if not os.path.exists(CREDENTIALS_PATH):
        raise FileNotFoundError(
            f"Missing credentials.json at: {CREDENTIALS_PATH}\n"
            "You need OAuth Client ID (Desktop App) credentials saved as credentials.json"
        )

    creds: Optional[Credentials] = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)


def month_window_from_shifts(
    shifts: List[Tuple[datetime, datetime]],
    tz_name: str,
) -> Tuple[datetime, datetime]:
    """
    Compute a month window for fetching existing events.
    Returns timezone-aware datetimes in tz_name.
    """
    tz = ZoneInfo(tz_name)

    if not shifts:
        now = datetime.now(tz)
        start = datetime(now.year, now.month, 1, 0, 0, 0, tzinfo=tz)
        end = start + timedelta(days=40)
        return start, end

    min_dt = min(s[0] for s in shifts)
    max_dt = max(s[1] for s in shifts)

    # ensure aware
    if min_dt.tzinfo is None:
        min_dt = min_dt.replace(tzinfo=tz)
    if max_dt.tzinfo is None:
        max_dt = max_dt.replace(tzinfo=tz)

    start = datetime(min_dt.year, min_dt.month, 1, 0, 0, 0, tzinfo=tz)
    _, last_day = calendar.monthrange(max_dt.year, max_dt.month)
    end = datetime(max_dt.year, max_dt.month, last_day, 23, 59, 59, tzinfo=tz)
    return start, end

def fetch_existing_keys(service, calendar_id: str, title: str, time_min: datetime, time_max: datetime) -> Set[str]:
    """
    Fetch existing events in [time_min, time_max] and build keys for those matching title.
    """
    existing: Set[str] = set()
    page_token = None

    while True:
        events = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min.isoformat(),
            timeMax=time_max.isoformat(),
            singleEvents=True,
            orderBy="startTime",
            q=title,                 # helps narrow; we still verify summary
            pageToken=page_token,
            maxResults=2500,
        ).execute()

        for ev in events.get("items", []):
            if ev.get("summary") != title:
                continue

            start = ev.get("start", {}).get("dateTime")
            end = ev.get("end", {}).get("dateTime")
            if not start or not end:
                continue

            # Parse RFC3339; remove timezone offset by using fromisoformat
            # (Python can parse offsets like '+02:00' fine)
            start_dt = datetime.fromisoformat(start)
            end_dt = datetime.fromisoformat(end)

            # Normalize to clock time key in local date; keep as-is
            k = f"{title}|{start_dt.date().isoformat()}|{start_dt.strftime('%H:%M')}|{end_dt.strftime('%H:%M')}"
            existing.add(k)

        page_token = events.get("nextPageToken")
        if not page_token:
            break

    return existing


def push_shifts_with_dedupe(
    shifts: List[Tuple[datetime, datetime]],
    calendar_id: str,
    timezone: str,
    title: str,
    dry_run: bool,
) -> int:
    if not shifts:
        print("\nNo shifts detected. Nothing to push.")
        return 0

    service = get_calendar_service()
    win_start, win_end = month_window_from_shifts(shifts, TIMEZONE)
    existing_keys = fetch_existing_keys(service, calendar_id, title, win_start, win_end)

    created = 0
    skipped = 0

    for start_dt, end_dt in shifts:
        k = event_key(title, start_dt, end_dt)
        if k in existing_keys:
            skipped += 1
            continue

        event = {
            "summary": title,
            "start": {"dateTime": start_dt.isoformat(), "timeZone": timezone},
            "end": {"dateTime": end_dt.isoformat(), "timeZone": timezone},
            "description": f"source=master_schedule_csv; key={k}",
        }

        if dry_run:
            print(f"[DRY RUN] Would create: {title} | {start_dt} -> {end_dt}")
        else:
            service.events().insert(calendarId=calendar_id, body=event).execute()
            existing_keys.add(k)

        created += 1

    print(f"\nDone. Created: {created}, Skipped (already existed): {skipped}")
    return created


# ==============================
# Main
# ==============================
def main() -> None:
    shifts, off_days = extract_shifts_and_offdays(
        master_csv_path=MASTER_CSV_PATH,
        employee_name=TARGET_EMPLOYEE,
        year=YEAR,
    )
    shifts = dedupe_shifts(shifts)
    print(f"Unique shifts after dedupe: {len(shifts)}")
    print(f"Detected shifts: {len(shifts)}")
    off_day_summary(off_days)

    print("\nPushing to Google Calendar...")
    push_shifts_with_dedupe(
        shifts=shifts,
        calendar_id=CALENDAR_ID,
        timezone=TIMEZONE,
        title=EVENT_TITLE,
        dry_run=DRY_RUN,
    )

    if DRY_RUN:
        print("\nSet DRY_RUN = False to actually create the events.")
        print(f"After first OAuth login, token will be saved to: {TOKEN_PATH}")


if __name__ == "__main__":
    main()