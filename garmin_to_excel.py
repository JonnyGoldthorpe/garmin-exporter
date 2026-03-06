"""
Health Data → Excel Daily Exporter
------------------------------------
Single sheet combining Garmin + Renpho data by date.

Garmin columns: steps, distance, sleep, HR, stress, body battery, activities
Renpho columns: weight, BMI, body fat, muscle mass, and more

Requirements:
    pip install garminconnect openpyxl requests

Environment variables (set in GitHub Secrets):
    GARMIN_EMAIL        Garmin login email
    GARMIN_PASSWORD     Garmin login password
    RENPHO_EMAIL        Renpho login email
    RENPHO_PASSWORD     Renpho login password
    EMAIL_FROM          Gmail address to send from
    EMAIL_APP_PASSWORD  Gmail App Password
    EMAIL_TO            Address to receive the file

NOTE: Running this script will log you out of the Renpho mobile app.
"""

import datetime
import hashlib
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from getpass import getpass

import requests

try:
    import garminconnect
except ImportError:
    raise SystemExit("Please run: pip install garminconnect openpyxl requests")

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("Please run: pip install garminconnect openpyxl requests")


# ── CONFIG ────────────────────────────────────────────────────────────────────
EXCEL_FILE    = "health_data.xlsx"
DAYS_TO_FETCH = 7
# ─────────────────────────────────────────────────────────────────────────────


# ══════════════════════════════════════════════════════════════════════════════
# COLUMNS — order defines the sheet layout
# ══════════════════════════════════════════════════════════════════════════════

COLUMNS = [
    # ── Date ──
    ("date",               "Date"),

    # ── Garmin: Activity ──
    ("steps",              "Steps"),
    ("distance_km",        "Distance (km)"),
    ("active_calories",    "Active Calories"),
    ("total_calories",     "Total Calories"),
    ("floors_climbed",     "Floors Climbed"),
    ("active_minutes",     "Active Minutes"),

    # ── Garmin: Heart Rate ──
    ("resting_heart_rate", "Resting HR"),
    ("min_heart_rate",     "Min HR"),
    ("max_heart_rate",     "Max HR"),

    # ── Garmin: Wellness ──
    ("avg_stress",         "Avg Stress"),
    ("body_battery_high",  "Body Battery High"),
    ("body_battery_low",   "Body Battery Low"),

    # ── Garmin: Sleep ──
    ("sleep_hours",        "Sleep (hrs)"),
    ("sleep_score",        "Sleep Score"),
    ("deep_sleep_min",     "Deep Sleep (min)"),
    ("light_sleep_min",    "Light Sleep (min)"),
    ("rem_sleep_min",      "REM Sleep (min)"),
    ("awake_min",          "Awake (min)"),

    # ── Garmin: Activity detail ──
    ("activity_type",      "Activity Type"),
    ("activity_name",      "Activity Name"),
    ("activity_duration",  "Activity Duration (min)"),
    ("activity_distance",  "Activity Distance (km)"),
    ("activity_avg_hr",    "Activity Avg HR"),
    ("activity_max_hr",    "Activity Max HR"),
    ("activity_calories",  "Activity Calories"),

    # ── Renpho: Body ──
    ("weight_kg",          "Weight (kg)"),
    ("bmi",                "BMI"),
    ("body_fat_pct",       "Body Fat %"),
    ("muscle_mass_kg",     "Muscle Mass (kg)"),
    ("bone_mass_kg",       "Bone Mass (kg)"),
    ("water_pct",          "Water %"),
    ("visceral_fat",       "Visceral Fat"),
    ("bmr_kcal",           "BMR (kcal)"),
    ("metabolic_age",      "Metabolic Age"),
    ("protein_pct",        "Protein %"),
]

# Column index where Renpho data starts (0-based within COLUMNS list)
RENPHO_START_KEY = "weight_kg"


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
# RENPHO
# ══════════════════════════════════════════════════════════════════════════════

RENPHO_BASE = "https://renpho.qnclouds.com/api"

def renpho_login(email: str, password: str) -> dict:
    hashed = hashlib.md5(password.encode()).hexdigest()
    resp = requests.post(
        f"{RENPHO_BASE}/v3/users/sign_in.json?app_id=Renpho",
        data={"secure_flag": 1, "email": email, "password": hashed}
    )
    resp.raise_for_status()
    data = resp.json()
    return {
        "session_key": data["terminal_user_session_key"],
        "user_id": str(data["user_info"]["id"])
    }

def fetch_renpho_data() -> dict:
    """Returns a dict keyed by date string → renpho measurement dict."""
    email    = os.environ.get("RENPHO_EMAIL")    or input("Renpho email: ")
    password = os.environ.get("RENPHO_PASSWORD") or getpass("Renpho password: ")

    try:
        session = renpho_login(email, password)
        print("✅ Logged in to Renpho")
        last_at = int((datetime.datetime.now() - datetime.timedelta(days=DAYS_TO_FETCH)).timestamp())
        resp = requests.get(
            f"{RENPHO_BASE}/v2/measurements/list.json",
            params={
                "user_id": session["user_id"],
                "last_at": last_at,
                "locale": "en",
                "app_id": "Renpho",
                "terminal_user_session_key": session["session_key"]
            }
        )
        resp.raise_for_status()
        measurements = resp.json().get("last_ary", [])
    except Exception as e:
        print(f"⚠️  Renpho fetch failed: {e}")
        return {}

    # If multiple weigh-ins on same day, keep the most recent
    by_date = {}
    for m in measurements:
        try:
            ts   = datetime.datetime.fromtimestamp(int(m.get("time_stamp", 0)))
            date = ts.strftime("%Y-%m-%d")
            by_date[date] = {
                "weight_kg":      m.get("weight", ""),
                "bmi":            m.get("bmi", ""),
                "body_fat_pct":   m.get("bodyfat", ""),
                "muscle_mass_kg": m.get("muscle", ""),
                "bone_mass_kg":   m.get("bone", ""),
                "water_pct":      m.get("water", ""),
                "visceral_fat":   m.get("visceral_fat", ""),
                "bmr_kcal":       m.get("bmr", ""),
                "metabolic_age":  m.get("body_age", ""),
                "protein_pct":    m.get("protein", ""),
            }
        except Exception:
            continue

    print(f"  Found {len(by_date)} Renpho measurement(s)")
    return by_date


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL
# ══════════════════════════════════════════════════════════════════════════════

# Find where Renpho columns start (1-based column index)
RENPHO_START_COL = next(
    (i + 1 for i, (key, _) in enumerate(COLUMNS) if key == RENPHO_START_KEY), None
)

def style_header(ws):
    for col_idx, (_, label) in enumerate(COLUMNS, start=1):
        # Blue for Garmin columns, green for Renpho columns
        if col_idx < RENPHO_START_COL:
            colour = "1F4E79"  # dark blue
        else:
            colour = "1E5631"  # dark green

        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.fill = PatternFill("solid", fgColor=colour)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(label) + 4, 14)

def get_existing_dates(ws) -> set:
    return {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)}

def append_row(ws, data: dict, row_num: int):
    row = [data.get(key, "") for key, _ in COLUMNS]
    ws.append(row)
    if row_num % 2 == 0:
        for col_idx in range(1, len(COLUMNS) + 1):
            colour = "D6E4F0" if col_idx < RENPHO_START_COL else "D8F0DC"
            ws.cell(row=ws.max_row, column=col_idx).fill = PatternFill("solid", fgColor=colour)

def update_renpho_in_row(ws, row_num: int, renpho: dict):
    """Fill in Renpho columns for an existing row."""
    for col_idx, (key, _) in enumerate(COLUMNS, start=1):
        if key in renpho:
            ws.cell(row=row_num, column=col_idx).value = renpho[key]

def save_to_excel(garmin_data: list, renpho_by_date: dict) -> tuple:
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Health Data"
        style_header(ws)

    # Build map of existing date → excel row number
    date_to_row = {
        ws.cell(row=r, column=1).value: r
        for r in range(2, ws.max_row + 1)
    }

    garmin_added = 0
    renpho_added = 0

    for row_data in sorted(garmin_data, key=lambda x: x["date"]):
        date = row_data["date"]

        # Merge Renpho data if available for this date
        if date in renpho_by_date:
            row_data.update(renpho_by_date[date])

        if date not in date_to_row:
            # New row
            row_num = ws.max_row + 1 - 1  # will be appended as next row
            append_row(ws, row_data, ws.max_row + 1)
            date_to_row[date] = ws.max_row
            garmin_added += 1
            if date in renpho_by_date:
                renpho_added += 1
        else:
            # Row exists — update Renpho columns only if we have new data
            if date in renpho_by_date:
                update_renpho_in_row(ws, date_to_row[date], renpho_by_date[date])
                renpho_added += 1

    ws.freeze_panes = "A2"
    wb.save(EXCEL_FILE)
    return garmin_added, renpho_added


# ══════════════════════════════════════════════════════════════════════════════
# EMAIL
# ══════════════════════════════════════════════════════════════════════════════

def send_email(garmin_added: int, renpho_added: int):
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
  • Renpho:  {renpho_added} measurement(s) added/updated

All data is on one sheet — blue columns are Garmin, green columns are Renpho.
Renpho columns will be blank on days you didn't weigh in.

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

    # ── Garmin ──
    print(f"\n📊 Fetching Garmin data (last {DAYS_TO_FETCH} days)...")
    garmin_client = garmin_login()
    today = datetime.date.today()
    dates = [today - datetime.timedelta(days=i) for i in range(DAYS_TO_FETCH - 1, -1, -1)]
    garmin_data = []
    for date in dates:
        print(f"  📅 {date.isoformat()}...", end=" ", flush=True)
        garmin_data.append(fetch_garmin_day(garmin_client, date))
        print("done")

    # ── Renpho ──
    print(f"\n⚖️  Fetching Renpho data (last {DAYS_TO_FETCH} days)...")
    renpho_by_date = fetch_renpho_data()

    # ── Save ──
    print(f"\n💾 Saving to {EXCEL_FILE}...")
    garmin_added, renpho_added = save_to_excel(garmin_data, renpho_by_date)
    print(f"✅ Garmin: {garmin_added} new row(s) | Renpho: {renpho_added} measurement(s)")

    # ── Email ──
    print("\n📧 Sending email...")
    send_email(garmin_added, renpho_added)


if __name__ == "__main__":
    main()
