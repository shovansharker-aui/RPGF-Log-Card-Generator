"""
workspace.py

One folder holds the equipment workbook and every template. The user picks
the folder once; the files are found by name.
"""

import json
from pathlib import Path

SETTINGS_FILE = "settings.json"

# Sheet names inside the single equipment workbook.
DUE_TAG_SHEET = "Due Tag"
LOG_CARD_SHEETS = ["EN&WH", "PR Part 1", "PR Part 2"]

LABELS = {
    "workbook": "Equipment workbook",
    "log_template": "Log card template",
    "trend_template": "Trend report template",
    "due_template": "Due tag template",
}

# Exact file names expected inside the workspace folder (case-insensitive).
FILE_NAMES = {
    "workbook": "Equipment_Workbook.xlsx",
    "log_template": "Log_Card_Template.docx",
    "trend_template": "Trend_Analysis_Template.xlsx",
    "due_template": "Due_Tag_Template.docx",
}


def load_folder():

    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as file:
            return json.load(file).get("workspace", "")
    except Exception:
        return ""


def save_folder(folder):

    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
    except Exception:
        data = {}

    data["workspace"] = str(folder)

    with open(SETTINGS_FILE, "w", encoding="utf-8") as file:
        json.dump(data, file, indent=4)


def find_files(folder):
    """Returns {key: Path or None} for the four files the app needs."""

    result = {key: None for key in FILE_NAMES}

    folder = Path(folder) if folder else None

    if not folder or not folder.is_dir():
        return result

    present = {
        path.name.lower(): path
        for path in folder.iterdir()
        if path.is_file() and not path.name.startswith("~$")
    }

    for key, name in FILE_NAMES.items():
        result[key] = present.get(name.lower())

    return result


def missing(files, *keys):
    """Human readable names of the required files that were not found."""

    return [f"{LABELS[key]}  ({FILE_NAMES[key]})"
            for key in keys if not files.get(key)]
