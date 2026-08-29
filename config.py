"""
config.py

Application Configuration
RPGF Log Card Generator v2.0
"""

from pathlib import Path

# ==========================================================
# Application Information
# ==========================================================

APP_NAME = "RPGF Log Card Generator"

VERSION = "2.0"

COMPANY = "Renata PLC"

AUTHOR = "Engineering Department"

# ==========================================================
# Base Directory
# ==========================================================

BASE_DIR = Path(__file__).resolve().parent

# ==========================================================
# Project Folders
# ==========================================================

ASSETS_DIR = BASE_DIR / "assets"

DATA_DIR = BASE_DIR / "data"

TEMPLATES_DIR = BASE_DIR / "templates"

OUTPUT_DIR = BASE_DIR / "output"

LOG_DIR = BASE_DIR / "logs"

# ==========================================================
# Create Required Folders
# ==========================================================

OUTPUT_DIR.mkdir(exist_ok=True)

LOG_DIR.mkdir(exist_ok=True)

# ==========================================================
# Application Files
# ==========================================================

SETTINGS_FILE = BASE_DIR / "settings.json"

EQUIPMENT_FILE = DATA_DIR / "Equipment.xlsx"

WORD_TEMPLATE = TEMPLATES_DIR / "Template.docx"

LOGO_FILE = ASSETS_DIR / "renata.png"

# Future Database
RESPONSIBILITY_DATABASE = DATA_DIR / "responsibility.db"

# ==========================================================
# Default Output Files
# ==========================================================

DEFAULT_LOGCARD_NAME = "Preventive_Maintenance_Log_Cards.docx"

DEFAULT_TREND_NAME = "Trend_Report.xlsx"

# ==========================================================
# Supported Weekdays
# ==========================================================

WEEKDAYS = [

    "Monday",

    "Tuesday",

    "Wednesday",

    "Thursday",

    "Friday",

    "Saturday",

    "Sunday"

]

# ==========================================================
# Supported Quarters
# ==========================================================

QUARTERS = [

    "Q1",

    "Q2",

    "Q3",

    "Q4"

]

# ==========================================================
# Workbook Sheets
# ==========================================================

EQUIPMENT_SHEETS = [

    "EN&WH",

    "PR Part 1",

    "PR Part 2"

]

# ==========================================================
# GUI
# ==========================================================

WINDOW_WIDTH = 900

WINDOW_HEIGHT = 650

WINDOW_RESIZABLE = False

# ==========================================================
# Trend Report
# ==========================================================

TREND_MONTHS = {

    "H1": [

        "Jan",

        "Feb",

        "Mar",

        "Apr",

        "May",

        "Jun"

    ],

    "H2": [

        "Jul",

        "Aug",

        "Sep",

        "Oct",

        "Nov",

        "Dec"

    ]

}