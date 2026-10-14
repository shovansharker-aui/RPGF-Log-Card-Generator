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


def inspect_workbook(path):
    """Quick health check of the equipment workbook.

    Returns {"counts": {sheet: machines}, "missing_sheets": [...],
    "error": str or None}. Reads only the first columns, so it is fast."""

    required = LOG_CARD_SHEETS + [DUE_TAG_SHEET]

    result = {"counts": {}, "missing_sheets": [], "error": None}

    try:
        from openpyxl import load_workbook

        book = load_workbook(path, read_only=True, data_only=True)

        try:
            for name in required:

                if name not in book.sheetnames:
                    result["missing_sheets"].append(name)
                    continue

                sheet = book[name]

                column = None
                count = 0

                for row in sheet.iter_rows(values_only=True):

                    if column is None:
                        headers = [str(v).strip() if v is not None else ""
                                   for v in row]
                        if "Equipment No." in headers:
                            column = headers.index("Equipment No.")
                        continue

                    if column < len(row) and row[column] not in (None, ""):
                        count += 1

                result["counts"][name] = count
        finally:
            book.close()

    except Exception as error:
        result["error"] = str(error)

    return result


def missing(files, *keys):
    """Human readable names of the required files that were not found."""

    return [f"{LABELS[key]}  ({FILE_NAMES[key]})"
            for key in keys if not files.get(key)]
