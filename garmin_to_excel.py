"""
Garmin Connect → Excel Daily Exporter (with email)
----------------------------------------------------
Pulls steps, distance, sleep, heart rate, and activities
from Garmin Connect, saves to Excel, and emails it to you.

Requirements:
    pip install garminconnect openpyxl

Setup — add these to your ~/.bashrc on PythonAnywhere:
    GARMIN_EMAIL        your Garmin login email
    GARMIN_PASSWORD     your Garmin login password
    EMAIL_FROM          Gmail address to send FROM
    EMAIL_APP_PASSWORD  Gmail App Password (see instructions below)
    EMAIL_TO            address to receive the file (can be same as above)

How to get a Gmail App Password:
    1. Go to myaccount.google.com → Security
    2. Enable 2-Step Verification if not already on
    3. Search "App passwords" → create one named "Garmin Script"
    4. Copy the 16-character password — use that as EMAIL_APP_PASSWORD
"""

import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from getpass import getpass

try:
    import garminconnect
except ImportError:
    raise SystemExit("Please run:  pip install garminconnect openpyxl")

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("Please run:  pip install garminconnect openpyxl")


# ── CONFIG ───────────────────────────────────────────────────────────────────
EXCEL_FILE    = "garmin_data.xlsx"  # Created in same folder as this script
DAYS_TO_FETCH = 7                   # How many past days to pull on first run
# ─────────────────────────────────────────────────────────────────────────────


def login():
    email    = os.environ.get("GARMIN_EMAIL")    or input("Garmin email: ")
    password = os.environ.get("GARMIN_PASSWORD") or getpass("Garmin password: ")
    client = garminconnect.Garmin(email, password)
    client.login()
    print("✅ Logged in to Garmin\n")
    return client


def fetch_day(client, date: datetime.date) -> dict:
    d = date.isoformat()
    data = {"date": d}

    try:
        steps_data = client.get_steps_data(d)
        data["steps"] = sum(s.get("steps", 0) for s in steps_data) if steps_data else 0
    except Exception:
        data["steps"] = ""

    try:
        daily = client.get_stats(d)
        data["distance_km"]         = round(daily.get("totalDistanceMeters", 0) / 1000, 2)
        data["active_calories"]     = daily.get("activeKilocalories", "")
        data["total_calories"]      = daily.get("totalKilocalories", "")
        data["floors_climbed"]      = daily.get("floorsAscended", "")
        data["active_minutes"]      = daily.get("highlyActiveSeconds", 0) // 60
        data["resting_heart_rate"]  = daily.get("restingHeartRate", "")
        data["avg_stress"]          = daily.get("averageStressLevel", "")
        data["body_battery_high"]   = daily.get("bodyBatteryHighestValue", "")
        data["body_battery_low"]    = daily.get("bodyBatteryLowestValue", "")
    except Exception:
        for k in ["distance_km","active_calories","total_calories","floors_climbed",
                  "active_minutes","resting_heart_rate","avg_stress","body_battery_high","body_battery_low"]:
            data.setdefault(k, "")

    try:
        hr = client.get_heart_rates(d)
        data["max_heart_rate"] = hr.get("maxHeartRate", "")
        data["min_heart_rate"] = hr.get("minHeartRate", "")
    except Exception:
        data["max_heart_rate"] = ""
        data["min_heart_rate"] = ""

    try:
        sleep   = client.get_sleep_data(d)
        summary = sleep.get("dailySleepDTO", {})
        sleep_seconds = summary.get("sleepTimeSeconds", 0) or 0
        data["sleep_hours"]     = round(sleep_seconds / 3600, 2)
        data["sleep_score"]     = summary.get("sleepScores", {}).get("overall", {}).get("value", "") \
                                  if isinstance(summary.get("sleepScores"), dict) else ""
        data["deep_sleep_min"]  = round((summary.get("deepSleepSeconds",  0) or 0) / 60)
        data["light_sleep_min"] = round((summary.get("lightSleepSeconds", 0) or 0) / 60)
        data["rem_sleep_min"]   = round((summary.get("remSleepSeconds",   0) or 0) / 60)
        data["awake_min"]       = round((summary.get("awakeSleepSeconds", 0) or 0) / 60)
    except Exception:
        for k in ["sleep_hours","sleep_score","deep_sleep_min","light_sleep_min","rem_sleep_min","awake_min"]:
            data.setdefault(k, "")

    try:
        activities = client.get_activities_by_date(d, d)
        if activities:
            act = activities[0]
            data["activity_type"]     = act.get("activityType", {}).get("typeKey", "")
            data["activity_name"]     = act.get("activityName", "")
            data["activity_duration"] = round((act.get("duration", 0) or 0) / 60, 1)
            data["activity_distance"] = round((act.get("distance", 0) or 0) / 1000, 2)
            data["activity_avg_hr"]   = act.get("averageHR", "")
            data["activity_max_hr"]   = act.get("maxHR", "")
            data["activity_calories"] = act.get("calories", "")
        else:
            for k in ["activity_type","activity_name","activity_duration",
                      "activity_distance","activity_avg_hr","activity_max_hr","activity_calories"]:
                data[k] = ""
    except Exception:
        for k in ["activity_type","activity_name","activity_duration",
                  "activity_distance","activity_avg_hr","activity_max_hr","activity_calories"]:
            data.setdefault(k, "")

    return data


COLUMNS = [
    ("date",               "Date"),
    ("steps",              "Steps"),
    ("distance_km",        "Distance (km)"),
    ("active_calories",    "Active Calories"),
    ("total_calories",     "Total Calories"),
    ("floors_climbed",     "Floors Climbed"),
    ("active_minutes",     "Active Minutes"),
    ("resting_heart_rate", "Resting HR"),
    ("min_heart_rate",     "Min HR"),
    ("max_heart_rate",     "Max HR"),
    ("avg_stress",         "Avg Stress"),
    ("body_battery_high",  "Body Battery High"),
    ("body_battery_low",   "Body Battery Low"),
    ("sleep_hours",        "Sleep (hrs)"),
    ("sleep_score",        "Sleep Score"),
    ("deep_sleep_min",     "Deep Sleep (min)"),
    ("light_sleep_min",    "Light Sleep (min)"),
    ("rem_sleep_min",      "REM Sleep (min)"),
    ("awake_min",          "Awake (min)"),
    ("activity_type",      "Activity Type"),
    ("activity_name",      "Activity Name"),
    ("activity_duration",  "Activity Duration (min)"),
    ("activity_distance",  "Activity Distance (km)"),
    ("activity_avg_hr",    "Activity Avg HR"),
    ("activity_max_hr",    "Activity Max HR"),
    ("activity_calories",  "Activity Calories"),
]


def style_header(ws):
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for col_idx, (_, label) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(label) + 4, 14)


def get_existing_dates(ws) -> set:
    return {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)}


def append_row(ws, data: dict):
    row = [data.get(key, "") for key, _ in COLUMNS]
    ws.append(row)
    last_row = ws.max_row
    if last_row % 2 == 0:
        fill = PatternFill("solid", fgColor="D6E4F0")
        for col_idx in range(1, len(COLUMNS) + 1):
            ws.cell(row=last_row, column=col_idx).fill = fill


def save_to_excel(all_data: list) -> int:
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Garmin Data"
        style_header(ws)

    existing = get_existing_dates(ws)
    added = 0
    for row_data in sorted(all_data, key=lambda x: x["date"]):
        if row_data["date"] not in existing:
            append_row(ws, row_data)
            added += 1

    ws.freeze_panes = "A2"
    wb.save(EXCEL_FILE)
    return added


def send_email(added: int):
    sender       = os.environ.get("EMAIL_FROM")
    app_password = os.environ.get("EMAIL_APP_PASSWORD")
    recipient    = os.environ.get("EMAIL_TO")

    if not all([sender, app_password, recipient]):
        print("⚠️  Email credentials not set — skipping email.")
        print("   Set EMAIL_FROM, EMAIL_APP_PASSWORD, and EMAIL_TO in ~/.bashrc")
        return

    today = datetime.date.today().strftime("%d %b %Y")

    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = recipient
    msg["Subject"] = f"📊 Garmin Data — {today}"

    body = f"""Hi,

Your daily Garmin data export is attached ({today}).

{added} new row(s) were added today.

Tracked: Steps, Distance, Calories, Heart Rate, Sleep, Stress, Body Battery & Activities.

— Your Garmin Exporter
"""
    msg.attach(MIMEText(body, "plain"))

    with open(EXCEL_FILE, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="garmin_data_{today}.xlsx"')
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, app_password)
        server.sendmail(sender, recipient, msg.as_string())

    print(f"📧 Email sent to {recipient}")


def main():
    print("=" * 50)
    print("   Garmin Connect → Excel Exporter")
    print("=" * 50)
    print(f"Fetching last {DAYS_TO_FETCH} day(s) of data...\n")

    client = login()

    today = datetime.date.today()
    dates = [today - datetime.timedelta(days=i) for i in range(DAYS_TO_FETCH - 1, -1, -1)]

    all_data = []
    for date in dates:
        print(f"  📅 Fetching {date.isoformat()}...", end=" ", flush=True)
        row = fetch_day(client, date)
        all_data.append(row)
        print("done")

    print(f"\n💾 Saving to {EXCEL_FILE}...")
    added = save_to_excel(all_data)
    print(f"✅ {added} new row(s) added → {os.path.abspath(EXCEL_FILE)}")

    print("\n📧 Sending email...")
    send_email(added)


if __name__ == "__main__":
    main()
