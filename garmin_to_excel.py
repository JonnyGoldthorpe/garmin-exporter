"""
Health Data → Excel Daily Exporter
------------------------------------
Pulls Garmin activity, wellness, heart rate, sleep and activity data.
Renpho body composition is merged from a CSV file dropped into the repo.

Requirements:
    pip install garminconnect openpyxl

Environment variables (set in GitHub Secrets):
    GARMIN_EMAIL        Garmin login email
    GARMIN_PASSWORD     Garmin login password
    EMAIL_FROM          Gmail address to send from
    EMAIL_APP_PASSWORD  Gmail App Password
    EMAIL_TO            Address to receive the file
"""

import csv
import datetime
import glob
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from getpass import getpass

try:
    import garminconnect
except ImportError:
    raise SystemExit("Please run: pip install garminconnect openpyxl")

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("Please run: pip install garminconnect openpyxl")


# ── CONFIG ────────────────────────────────────────────────────────────────────
EXCEL_FILE    = "health_data.xlsx"
DAYS_TO_FETCH = 7
RENPHO_CSV_GLOB = "renpho*.csv"  # matches any file starting with "renpho"
# ─────────────────────────────────────────────────────────────────────────────


COLUMNS = [
    # Date
    ("date",               "Date"),

    # Garmin: Activity
    ("steps",              "Steps"),
    ("distance_km",        "Distance (km)"),
    ("active_calories",    "Active Calories"),
    ("total_calories",     "Total Calories"),
    ("floors_climbed",     "Floors Climbed"),
    ("active_minutes",     "Active Minutes"),

    # Garmin: Heart Rate
    ("resting_heart_rate", "Resting HR"),
    ("min_heart_rate",     "Min HR"),
    ("max_heart_rate",     "Max HR"),

    # Garmin: Wellness
    ("avg_stress",         "Avg Stress"),
    ("body_battery_high",  "Body Battery High"),
    ("body_battery_low",   "Body Battery Low"),

    # Garmin: Sleep
    ("sleep_hours",        "Sleep (hrs)"),
    ("sleep_score",        "Sleep Score"),
    ("deep_sleep_min",     "Deep Sleep (min)"),
    ("light_sleep_min",    "Light Sleep (min)"),
    ("rem_sleep_min",      "REM Sleep (min)"),
    ("awake_min",          "Awake (min)"),

    # Garmin: Activity detail
    ("activity_type",      "Activity Type"),
    ("activity_name",      "Activity Name"),
    ("activity_duration",  "Activity Duration (min)"),
    ("activity_distance",  "Activity Distance (km)"),
    ("activity_avg_hr",    "Activity Avg HR"),
    ("activity_max_hr",    "Activity Max HR"),
    ("activity_calories",  "Activity Calories"),

    # Renpho: Body composition (from CSV)
    ("weight_kg",          "Weight (kg)"),
    ("bmi",                "BMI"),
    ("body_fat_pct",       "Body Fat %"),
    ("muscle_mass_kg",     "Muscle Mass (kg)"),
    ("bone_mass_kg",       "Bone Mass (kg)"),
    ("body_water_pct",     "Body Water %"),
    ("visceral_fat",       "Visceral Fat"),
    ("bmr_kcal",           "BMR (kcal)"),
    ("metabolic_age",      "Metabolic Age"),
]

RENPHO_START_KEY = "weight_kg"
RENPHO_START_COL = next(i + 1 for i, (k, _) in enumerate(COLUMNS) if k == RENPHO_START_KEY)


# ══════════════════════════════════════════════════════════════════════════════
# GARMIN
# ══════════════════════════════════════════════════════════════════════════════

def garmin_login():
    email    = os.environ.get("GARMIN_EMAIL")    or input("Garmin email: ")
    password = os.environ.get("GARMIN_PASSWORD") or getpass("Garmin password: ")
    client = garminconnect.Garmin(email, password)
    client.login()
    print("✅ Logged in to Garmin")
    return client


def fetch_garmin_day(client, date: datetime.date) -> dict:
    d = date.isoformat()
    data = {"date": d}

    try:
        steps_data = client.get_steps_data(d)
        data["steps"] = sum(s.get("steps", 0) for s in steps_data) if steps_data else 0
    except Exception:
        data["steps"] = ""

    try:
        daily = client.get_stats(d)
        data["distance_km"]        = round(daily.get("totalDistanceMeters", 0) / 1000, 2)
        data["active_calories"]    = daily.get("activeKilocalories", "")
        data["total_calories"]     = daily.get("totalKilocalories", "")
        data["floors_climbed"]     = daily.get("floorsAscended", "")
        data["active_minutes"]     = daily.get("highlyActiveSeconds", 0) // 60
        data["resting_heart_rate"] = daily.get("restingHeartRate", "")
        data["avg_stress"]         = daily.get("averageStressLevel", "")
        data["body_battery_high"]  = daily.get("bodyBatteryHighestValue", "")
        data["body_battery_low"]   = daily.get("bodyBatteryLowestValue", "")
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


# ══════════════════════════════════════════════════════════════════════════════
# RENPHO CSV
# ══════════════════════════════════════════════════════════════════════════════

# Map from possible Renpho CSV column names → our internal keys
RENPHO_COL_MAP = {
    "weight(kg)":       "weight_kg",
    "weight":           "weight_kg",
    "bmi":              "bmi",
    "body fat(%)":      "body_fat_pct",
    "body fat":         "body_fat_pct",
    "muscle mass(kg)":  "muscle_mass_kg",
    "muscle mass":      "muscle_mass_kg",
    "bone mass(kg)":    "bone_mass_kg",
    "bone mass":        "bone_mass_kg",
    "body water(%)":    "body_water_pct",
    "body water":       "body_water_pct",
    "visceral fat":     "visceral_fat",
    "bmr(kcal)":        "bmr_kcal",
    "bmr":              "bmr_kcal",
    "metabolic age":    "metabolic_age",
}

def load_renpho_csv() -> dict:
    """Load all Renpho CSV files in the repo, return dict keyed by date."""
    files = glob.glob(RENPHO_CSV_GLOB)
    if not files:
        print("  No Renpho CSV file found — body composition columns will be empty.")
        return {}

    by_date = {}
    for filepath in files:
        try:
            with open(filepath, newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # Normalise header names to lowercase stripped
                    norm = {k.lower().strip(): v for k, v in row.items()}

                    # Find the date column
                    date_str = norm.get("time of measurement") or norm.get("date") or norm.get("measurement time")
                    if not date_str:
                        continue

                    # Parse date — handle common formats
                    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
                        try:
                            date = datetime.datetime.strptime(date_str.strip(), fmt).strftime("%Y-%m-%d")
                            break
                        except ValueError:
                            continue
                    else:
                        continue

                    entry = {}
                    for csv_col, our_key in RENPHO_COL_MAP.items():
                        val = norm.get(csv_col, "")
                        if val:
                            try:
                                entry[our_key] = round(float(val), 2)
                            except ValueError:
                                entry[our_key] = val

                    if entry:
                        by_date[date] = entry  # most recent entry wins if duplicates

        except Exception as e:
            print(f"  ⚠️  Could not read {filepath}: {e}")

    print(f"  Loaded {len(by_date)} Renpho measurement(s) from {len(files)} file(s)")
    return by_date


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def style_header(ws):
    for col_idx, (_, label) in enumerate(COLUMNS, start=1):
        colour = "1F4E79" if col_idx < RENPHO_START_COL else "1E5631"
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.fill = PatternFill("solid", fgColor=colour)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(label) + 4, 14)


def get_existing_dates(ws) -> dict:
    return {ws.cell(row=r, column=1).value: r for r in range(2, ws.max_row + 1)}


def append_row(ws, data: dict):
    row = [data.get(key, "") for key, _ in COLUMNS]
    ws.append(row)
    last_row = ws.max_row
    if last_row % 2 == 0:
        for col_idx in range(1, len(COLUMNS) + 1):
            colour = "D6E4F0" if col_idx < RENPHO_START_COL else "D8F0DC"
            ws.cell(row=last_row, column=col_idx).fill = PatternFill("solid", fgColor=colour)


def update_row(ws, row_num: int, data: dict):
    for col_idx, (key, _) in enumerate(COLUMNS, start=1):
        val = data.get(key, "")
        if val != "":
            ws.cell(row=row_num, column=col_idx).value = val


def save_to_excel(garmin_data: list, renpho_by_date: dict) -> tuple:
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Health Data"
        style_header(ws)

    date_to_row = get_existing_dates(ws)

    garmin_added = 0
    renpho_merged = 0

    for row_data in sorted(garmin_data, key=lambda x: x["date"]):
        date = row_data["date"]

        # Merge Renpho data if available
        if date in renpho_by_date:
            row_data.update(renpho_by_date[date])
            renpho_merged += 1

        if date not in date_to_row:
            append_row(ws, row_data)
            date_to_row[date] = ws.max_row
            garmin_added += 1
        else:
            update_row(ws, date_to_row[date], row_data)

    ws.freeze_panes = "A2"
    wb.save(EXCEL_FILE)
    return garmin_added, renpho_merged


# ══════════════════════════════════════════════════════════════════════════════
# EMAIL
# ══════════════════════════════════════════════════════════════════════════════

def send_email(garmin_added: int, renpho_merged: int):
    sender       = os.environ.get("EMAIL_FROM")
    app_password = os.environ.get("EMAIL_APP_PASSWORD")
    recipient    = os.environ.get("EMAIL_TO")

    if not all([sender, app_password, recipient]):
        print("⚠️  Email credentials not set — skipping email.")
        return

    today = datetime.date.today().strftime("%d %b %Y")
    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = recipient
    msg["Subject"] = f"💪 Health Data — {today}"

    body = f"""Hi,

Your daily health data export is attached ({today}).

  • Garmin:  {garmin_added} new row(s)
  • Renpho:  {renpho_merged} day(s) with body composition data

Blue columns = Garmin | Green columns = Renpho (from CSV)

— Your Health Exporter
"""
    msg.attach(MIMEText(body, "plain"))

    with open(EXCEL_FILE, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="health_data_{today}.xlsx"')
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, app_password)
        server.sendmail(sender, recipient, msg.as_string())

    print(f"📧 Email sent to {recipient}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 50)
    print("   Health Data → Excel Exporter")
    print("=" * 50)

    # Garmin
    print(f"\n📊 Fetching Garmin data (last {DAYS_TO_FETCH} days)...")
    client = garmin_login()
    today = datetime.date.today()
    dates = [today - datetime.timedelta(days=i) for i in range(DAYS_TO_FETCH - 1, -1, -1)]
    garmin_data = []
    for date in dates:
        print(f"  📅 {date.isoformat()}...", end=" ", flush=True)
        garmin_data.append(fetch_garmin_day(client, date))
        print("done")

    # Renpho CSV
    print(f"\n⚖️  Loading Renpho CSV...")
    renpho_by_date = load_renpho_csv()

    # Save
    print(f"\n💾 Saving to {EXCEL_FILE}...")
    garmin_added, renpho_merged = save_to_excel(garmin_data, renpho_by_date)
    print(f"✅ Garmin: {garmin_added} new row(s) | Renpho: {renpho_merged} day(s) merged")

    # Email
    print("\n📧 Sending email...")
    send_email(garmin_added, renpho_merged)


if __name__ == "__main__":
    main()
