"""
mail_helper.py

Creates an Outlook draft (opened on screen, never sent automatically) with
the schedule PDF attached.
"""

import html
import json
from pathlib import Path

import mail_config

SETTINGS_FILE = "settings.json"

MAIL_FIELDS = ("to", "cc", "subject", "body")


def default_mail_settings():
    """Factory values from mail_config.py (body: one paragraph per line)."""

    return {
        "to": mail_config.TO,
        "cc": mail_config.CC,
        "subject": mail_config.SUBJECT,
        "body": "\n".join(mail_config.BODY),
    }


def get_mail_settings():
    """Saved values from settings.json, falling back to the defaults."""

    settings = default_mail_settings()

    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as file:
            saved = json.load(file).get("mail", {})
    except Exception:
        saved = {}

    for field in MAIL_FIELDS:
        if isinstance(saved.get(field), str):
            settings[field] = saved[field]

    return settings


def save_mail_settings(settings):

    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
    except Exception:
        data = {}

    data["mail"] = {field: settings.get(field, "") for field in MAIL_FIELDS}

    with open(SETTINGS_FILE, "w", encoding="utf-8") as file:
        json.dump(data, file, indent=4)


def _ordinal(day):

    if 10 <= day % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")

    return suffix


def long_date(when):
    """06th October, 2026 (with a superscript suffix, as HTML)."""

    return (
        f"{when.day:02d}<sup>{_ordinal(when.day)}</sup> "
        f"{when.strftime('%B')}, {when.year}"
    )


def _paragraphs(lines, **values):
    """One <p> per line. Placeholders like {date_long} are swapped in after
    escaping, so any other braces the user types are left alone."""

    paragraphs = []

    for line in lines:

        text = html.escape(line)

        for key, value in values.items():
            text = text.replace("{" + key + "}", value)

        paragraphs.append("<p>" + text + "</p>")

    return "".join(paragraphs)


def create_outlook_draft(pdf_file, mail_date):
    """Open a new Outlook message addressed from mail_config with the PDF
    attached. The user reviews and sends it manually."""

    pdf_file = Path(pdf_file)

    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF not found:\n{pdf_file}")

    try:
        import win32com.client
    except ImportError as error:
        raise RuntimeError(
            "Outlook automation (pywin32) is not installed.") from error

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as error:
        raise RuntimeError(
            "Could not start Outlook. Make sure the classic Outlook desktop "
            "app is installed.") from error

    message = outlook.CreateItem(0)  # olMailItem

    settings = get_mail_settings()

    message.To = settings["to"]
    message.CC = settings["cc"]
    message.Subject = settings["subject"].replace(
        "{date}", mail_date.strftime("%d.%m.%Y"))

    try:
        message.Recipients.ResolveAll()
    except Exception:
        pass

    message.Attachments.Add(str(pdf_file.resolve()))

    # Display first so Outlook inserts the user's default signature, then
    # put the body above it.
    message.Display()

    existing = message.HTMLBody or ""

    lines = [line.strip() for line in settings["body"].splitlines()
             if line.strip()]

    body = _paragraphs(lines, date_long=long_date(mail_date),
                       date=mail_date.strftime("%d.%m.%Y"))

    lowered = existing.lower()

    if "<body" in lowered:
        marker = lowered.index(">", lowered.index("<body")) + 1
        message.HTMLBody = existing[:marker] + body + existing[marker:]
    else:
        message.HTMLBody = body + existing

    return message
