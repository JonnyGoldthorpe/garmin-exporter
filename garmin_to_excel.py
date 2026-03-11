"""
Garmin Health Data → Excel Weekly Exporter
-------------------------------------------
Pulls comprehensive data from Garmin Connect and emails:
  1. health_data.xlsx — all stats in one sheet
  2. fit_files.zip    — raw FIT files for each activity (for detailed analysis)

Requirements:
    pip install garminconnect openpyxl

Environment variables (set in GitHub Secrets):
    GARMIN_EMAIL        Garmin login email
    GARMIN_PASSWORD     Garmin login password
    EMAIL_FROM          Gmail address to send from
    EMAIL_APP_PASSWORD  Gmail App Password
    EMAIL_TO            Address to receive the file
"""

import datetime
import io
import os
import smtplib
import zipfile
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
EXCEL_FILE      = "health_data.xlsx"
DAYS_TO_FETCH   = 7
MAX_ACTIVITIES  = 3
# ─────────────────────────────────────────────────────────────────────────────


def activity_columns(n: int) -> list:
    prefix = f"act{n}_"
    label  = f"Activity {n}: "
    return [
        (prefix + "type",                   label + "Type"),
        (prefix + "name",                   label + "Name"),
        (prefix + "start_time",             label + "Start Time"),
        (prefix + "duration",               label + "Duration (min)"),
        (prefix + "distance",               label + "Distance (km)"),
        (prefix + "avg_hr",                 label + "Avg HR"),
        (prefix + "max_hr",                 label + "Max HR"),
        (prefix + "calories",               label + "Calories"),
        (prefix + "training_load",          label + "Training Load"),
        (prefix + "avg_pace",               label + "Avg Pace (min/km)"),
        (prefix + "best_pace",              label + "Best Pace (min/km)"),
        (prefix + "avg_cadence",            label + "Avg Cadence (spm)"),
        (prefix + "max_cadence",            label + "Max Cadence (spm)"),
        (prefix + "avg_power",              label + "Avg Power (W)"),
        (prefix + "max_power",              label + "Max Power (W)"),
        (prefix + "elevation_gain",         label + "Elevation Gain (m)"),
        (prefix + "elevation_loss",         label + "Elevation Loss (m)"),
        (prefix + "aerobic_effect",         label + "Aerobic Effect"),
        (prefix + "anaerobic_effect",       label + "Anaerobic Effect"),
        (prefix + "exercise_load",          label + "Exercise Load"),
        (prefix + "primary_benefit",        label + "Primary Benefit"),
        (prefix + "stamina_start",          label + "Stamina Start (%)"),
        (prefix + "stamina_end",            label + "Stamina End (%)"),
        (prefix + "stamina_min",            label + "Stamina Min (%)"),
        (prefix + "avg_vert_osc",           label + "Avg Vertical Osc (cm)"),
        (prefix + "avg_vert_ratio",         label + "Avg Vertical Ratio (%)"),
        (prefix + "avg_ground_contact",     label + "Avg Ground Contact (ms)"),
        (prefix + "avg_ground_balance",     label + "Avg Ground Balance (%)"),
        (prefix + "avg_stride_length",      label + "Avg Stride Length (m)"),
    ]


COLUMNS = [
    ("date",                    "Date"),
    ("steps",                   "Steps"),
    ("distance_km",             "Distance (km)"),
    ("active_calories",         "Active Calories"),
    ("total_calories",          "Total Calories"),
    ("floors_climbed",          "Floors Climbed"),
    ("active_minutes",          "Active Minutes"),
    ("resting_heart_rate",      "Resting HR"),
    ("min_heart_rate",          "Min HR"),
    ("max_heart_rate",          "Max HR"),
    ("hrv",                     "HRV"),
    ("avg_stress",              "Avg Stress"),
    ("body_battery_high",       "Body Battery High"),
    ("body_battery_low",        "Body Battery Low"),
    ("spo2",                    "SpO2 (%)"),
    ("respiration_avg",         "Avg Respiration"),
    ("hydration_goal_ml",       "Hydration Goal (ml)"),
    ("hydration_intake_ml",     "Hydration Intake (ml)"),
    ("sleep_hours",             "Sleep (hrs)"),
    ("sleep_score",             "Sleep Score"),
    ("deep_sleep_min",          "Deep Sleep (min)"),
    ("light_sleep_min",         "Light Sleep (min)"),
    ("rem_sleep_min",           "REM Sleep (min)"),
    ("awake_min",               "Awake (min)"),
    ("vo2_max",                 "VO2 Max"),
    ("race_5k",                 "Race Pred 5K"),
    ("race_10k",                "Race Pred 10K"),
    ("race_half",               "Race Pred Half Marathon"),
    ("race_marathon",           "Race Pred Marathon"),
]
for i in range(1, MAX_ACTIVITIES + 1):
    COLUMNS.extend(activity_columns(i))


COLUMN_GROUPS = [
    ("date",                "2C3E50"),
    ("steps",               "1F4E79"),
    ("resting_heart_rate",  "922B21"),
    ("avg_stress",          "1A5276"),
    ("sleep_hours",         "4A235A"),
    ("vo2_max",             "145A32"),
    ("act1_type",           "784212"),
    ("act2_type",           "6E2F00"),
    ("act3_type",           "5D2506"),
]

def get_col_colour(col_idx):
    colour = "2C3E50"
    for key, hex_colour in COLUMN_GROUPS:
        group_idx = next((i + 1 for i, (k, _) in enumerate(COLUMNS) if k == key), None)
        if group_idx and col_idx >= group_idx:
            colour = hex_colour
    return colour

ROW_TINTS = {
    "2C3E50": "F2F3F4", "1F4E79": "D6E4F0", "922B21": "FADBD8",
    "1A5276": "D4E6F1", "4A235A": "E8DAEF", "145A32": "D5F5E3",
    "784212": "FAE5D3", "6E2F00": "F5CBA7", "5D2506": "F0B27A",
}

def get_row_tint(col_idx):
    return ROW_TINTS.get(get_col_colour(col_idx), "F2F3F4")


def seconds_to_pace(spk):
    if not spk:
        return ""
    try:
        t = int(spk)
        return f"{t // 60}:{t % 60:02d}"
    except Exception:
        return ""

def seconds_to_time(s):
    if not s:
        return ""
    try:
        h = int(s) // 3600
        m = (int(s) % 3600) // 60
        sc = int(s) % 60
        return f"{h}:{m:02d}:{sc:02d}" if h else f"{m}:{sc:02d}"
    except Exception:
        return ""


def garmin_login():
    email    = os.environ.get("GARMIN_EMAIL")    or input("Garmin email: ")
    password = os.environ.get("GARMIN_PASSWORD") or getpass("Garmin password: ")
    client = garminconnect.Garmin(email, password)
    client.login()
    print("✅ Logged in to Garmin")
    return client


def parse_activity_to_dict(act, details, prefix):
    row = {}
    row[prefix + "type"]           = act.get("activityType", {}).get("typeKey", "")
    row[prefix + "name"]           = act.get("activityName", "")
    row[prefix + "start_time"]     = act.get("startTimeLocal", "")[11:16]
    row[prefix + "duration"]       = round((act.get("duration", 0) or 0) / 60, 1)
    row[prefix + "distance"]       = round((act.get("distance", 0) or 0) / 1000, 2)
    row[prefix + "avg_hr"]         = act.get("averageHR", "")
    row[prefix + "max_hr"]         = act.get("maxHR", "")
    row[prefix + "calories"]       = act.get("calories", "")
    row[prefix + "training_load"]  = act.get("activityTrainingLoad", "")
    row[prefix + "avg_cadence"]    = act.get("averageRunningCadenceInStepsPerMinute", "") or \
                                     act.get("averageBikingCadenceInRevPerMinute", "")
    row[prefix + "max_cadence"]    = act.get("maxRunningCadenceInStepsPerMinute", "") or \
                                     act.get("maxBikingCadenceInRevPerMinute", "")
    row[prefix + "avg_power"]      = act.get("avgPower", "")
    row[prefix + "max_power"]      = act.get("maxPower", "")
    row[prefix + "elevation_gain"] = act.get("elevationGain", "")
    row[prefix + "elevation_loss"] = act.get("elevationLoss", "")

    avg_speed = act.get("averageSpeed", 0)
    max_speed = act.get("maxSpeed", 0)
    row[prefix + "avg_pace"]  = seconds_to_pace(1000 / avg_speed) if avg_speed else ""
    row[prefix + "best_pace"] = seconds_to_pace(1000 / max_speed) if max_speed else ""

    if details:
        summary = details.get("summaryDTO", {})
        row[prefix + "aerobic_effect"]     = summary.get("aerobicTrainingEffect", "")
        row[prefix + "anaerobic_effect"]   = summary.get("anaerobicTrainingEffect", "")
        row[prefix + "exercise_load"]      = summary.get("activityTrainingLoad", "")
        row[prefix + "primary_benefit"]    = summary.get("aerobicTrainingEffectMessage", "")
        row[prefix + "stamina_start"]      = summary.get("startStamina", "")
        row[prefix + "stamina_end"]        = summary.get("endStamina", "")
        row[prefix + "stamina_min"]        = summary.get("minStamina", "")
        row[prefix + "avg_vert_osc"]       = summary.get("avgVerticalOscillation", "")
        row[prefix + "avg_vert_ratio"]     = summary.get("avgVerticalRatio", "")
        row[prefix + "avg_ground_contact"] = summary.get("avgGroundContactTime", "")
        row[prefix + "avg_ground_balance"] = summary.get("avgGroundContactBalance", "")
        row[prefix + "avg_stride_length"]  = round(summary.get("avgStrideLength", 0) / 100, 2) \
                                             if summary.get("avgStrideLength") else ""
    return row


def fetch_garmin_day(client, date):
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
        data["hydration_goal_ml"]   = daily.get("dailyHydrationGoal", "")
        data["hydration_intake_ml"] = daily.get("totalHydrationIntakeInOz", "")
    except Exception:
        for k in ["distance_km","active_calories","total_calories","floors_climbed","active_minutes",
                  "resting_heart_rate","avg_stress","body_battery_high","body_battery_low",
                  "hydration_goal_ml","hydration_intake_ml"]:
            data.setdefault(k, "")

    try:
        hr = client.get_heart_rates(d)
        data["max_heart_rate"] = hr.get("maxHeartRate", "")
        data["min_heart_rate"] = hr.get("minHeartRate", "")
    except Exception:
        data["max_heart_rate"] = ""
        data["min_heart_rate"] = ""

    try:
        hrv = client.get_hrv_data(d)
        data["hrv"] = hrv.get("hrvSummary", {}).get("weeklyAvg", "") if hrv else ""
    except Exception:
        data["hrv"] = ""

    try:
        spo2 = client.get_spo2_data(d)
        data["spo2"] = spo2.get("averageSpO2", "") if spo2 else ""
    except Exception:
        data["spo2"] = ""

    try:
        resp = client.get_respiration_data(d)
        data["respiration_avg"] = resp.get("avgWakingRespirationValue", "") if resp else ""
    except Exception:
        data["respiration_avg"] = ""

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
        vo2 = client.get_max_metrics(d)
        if vo2 and len(vo2) > 0:
            v = vo2[0]
            data["vo2_max"]       = v.get("generic", {}).get("vo2MaxPreciseValue", "")
            run_race = v.get("running", {})
            data["race_5k"]       = seconds_to_time(run_race.get("vo2MaxRacePredictions", {}).get("5K"))
            data["race_10k"]      = seconds_to_time(run_race.get("vo2MaxRacePredictions", {}).get("10K"))
            data["race_half"]     = seconds_to_time(run_race.get("vo2MaxRacePredictions", {}).get("halfMarathon"))
            data["race_marathon"] = seconds_to_time(run_race.get("vo2MaxRacePredictions", {}).get("marathon"))
    except Exception:
        for k in ["vo2_max","race_5k","race_10k","race_half","race_marathon"]:
            data.setdefault(k, "")

    try:
        activities = client.get_activities_by_date(d, d)
        for i, act in enumerate(activities[:MAX_ACTIVITIES], start=1):
            prefix = f"act{i}_"
            act_id = act.get("activityId")
            details = None
            try:
                details = client.get_activity(act_id)
            except Exception:
                pass
            data.update(parse_activity_to_dict(act, details, prefix))
    except Exception:
        pass

    return data


def download_fit_files(client, dates):
    """Download FIT files for all activities in the date range. Returns dict of filename→bytes."""
    fit_files = {}
    print("\n📦 Downloading FIT files...")

    for date in dates:
        d = date.isoformat()
        try:
            activities = client.get_activities_by_date(d, d)
            for act in activities[:MAX_ACTIVITIES]:
                act_id   = act.get("activityId")
                act_type = act.get("activityType", {}).get("typeKey", "activity")
                act_name = act.get("activityName", "").replace(" ", "-").lower()
                filename = f"{d}_{act_type}_{act_name}_{act_id}.fit"

                try:
                    fit_data = client.download_activity(
                        act_id,
                        dl_fmt=client.ActivityDownloadFormat.ORIGINAL
                    )
                    fit_files[filename] = fit_data
                    print(f"  ✅ {filename}")
                except Exception as e:
                    print(f"  ⚠️  Could not download {filename}: {e}")
        except Exception as e:
            print(f"  ⚠️  Could not fetch activities for {d}: {e}")

    return fit_files


def create_zip(fit_files: dict) -> bytes:
    """Zip all FIT files into a single bytes object."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, data in fit_files.items():
            zf.writestr(filename, data)
    return buf.getvalue()


def style_header(ws):
    for col_idx, (_, label) in enumerate(COLUMNS, start=1):
        colour = get_col_colour(col_idx)
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.fill = PatternFill("solid", fgColor=colour)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(label) + 4, 14)


def get_existing_dates(ws):
    return {ws.cell(row=r, column=1).value: r for r in range(2, ws.max_row + 1)}


def append_row(ws, data):
    row = [data.get(key, "") for key, _ in COLUMNS]
    ws.append(row)
    if ws.max_row % 2 == 0:
        for col_idx in range(1, len(COLUMNS) + 1):
            ws.cell(row=ws.max_row, column=col_idx).fill = PatternFill("solid", fgColor=get_row_tint(col_idx))


def update_row(ws, row_num, data):
    for col_idx, (key, _) in enumerate(COLUMNS, start=1):
        val = data.get(key, "")
        if val != "":
            ws.cell(row=row_num, column=col_idx).value = val


def save_to_excel(all_data):
    if os.path.exists(EXCEL_FILE):
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Health Data"
        style_header(ws)

    date_to_row = get_existing_dates(ws)
    added = 0

    for row_data in sorted(all_data, key=lambda x: x["date"]):
        date = row_data["date"]
        if date not in date_to_row:
            append_row(ws, row_data)
            date_to_row[date] = ws.max_row
            added += 1
        else:
            update_row(ws, date_to_row[date], row_data)

    ws.freeze_panes = "A2"
    wb.save(EXCEL_FILE)
    return added


def send_email(added, fit_files):
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
    msg["Subject"] = f"📊 Garmin Data — {today}"

    body = f"""Hi,

Your weekly Garmin data export is attached ({today}).

  • health_data.xlsx  — {added} new row(s) added this week
  • fit_files.zip     — {len(fit_files)} FIT file(s) for detailed activity analysis

Upload any FIT file to Claude for full HR, pace, split and GPS analysis.

— Your Health Exporter
"""
    msg.attach(MIMEText(body, "plain"))

    # Attach Excel
    with open(EXCEL_FILE, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="health_data_{today}.xlsx"')
        msg.attach(part)

    # Attach zipped FIT files
    if fit_files:
        zip_data = create_zip(fit_files)
        part2 = MIMEBase("application", "zip")
        part2.set_payload(zip_data)
        encoders.encode_base64(part2)
        part2.add_header("Content-Disposition", f'attachment; filename="fit_files_{today}.zip"')
        msg.attach(part2)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, app_password)
        server.sendmail(sender, recipient, msg.as_string())

    print(f"📧 Email sent to {recipient}")


def main():
    print("=" * 50)
    print("   Garmin Health Data → Excel Exporter")
    print("=" * 50)

    print(f"\n📊 Fetching Garmin data (last {DAYS_TO_FETCH} days)...")
    client = garmin_login()

    today = datetime.date.today()
    dates = [today - datetime.timedelta(days=i) for i in range(DAYS_TO_FETCH - 1, -1, -1)]

    all_data = []
    for date in dates:
        print(f"  📅 {date.isoformat()}...", end=" ", flush=True)
        all_data.append(fetch_garmin_day(client, date))
        print("done")

    print(f"\n💾 Saving to {EXCEL_FILE}...")
    added = save_to_excel(all_data)
    print(f"✅ {added} new row(s) added")

    fit_files = download_fit_files(client, dates)
    print(f"✅ {len(fit_files)} FIT file(s) downloaded")

    print("\n📧 Sending email...")
    send_email(added, fit_files)


if __name__ == "__main__":
    main()
