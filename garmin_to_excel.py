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
DAYS_TO_FETCH   = 8
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
        sc
