"""
schedule_controller.py

Weekly Maintenance Schedule: lists the machines that need maintenance on
the upcoming Friday, based on the maintenance matrix workbook.
"""

import calendar
import re
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

FRIDAY = 4

# A date counts as "this Friday" when it is within this many days of it.
WINDOW_DAYS = 3

# Last allowable date = scheduled date + this many days.
GRACE_DAYS = 3

# Order frequencies are written in the output.
FREQUENCY_ORDER = ["W", "M", "3M", "6M", "Y", "2Y", "6Y"]

# Years that must pass since the last 2Y/6Y maintenance before it is due.
LONG_INTERVALS = {"2Y": 2, "6Y": 6}

HEADERS = ["Sl No", "Equipment Name", "Equipment ID", "Frequency",
           "Schedule Date", "Last Date"]


def upcoming_friday(today=None):
    """Today if it is Friday, otherwise the next Friday."""

    today = today or date.today()

    return today + timedelta(days=(FRIDAY - today.weekday()) % 7)


class ScheduleController:

    def __init__(self, excel_file, output_file, friday=None, progress=None):

        self.excel_file = Path(excel_file)
        self.output_file = Path(output_file)
        self.friday = friday or upcoming_friday()
        self.progress = progress

    def _report(self, current, total, message):

        if self.progress:
            self.progress(current, total, message)

    # ==================================================
    # Parsing helpers
    # ==================================================

    @staticmethod
    def _tokens(value):

        return {
            token.strip().upper()
            for token in str(value).split(",")
            if token.strip() and token.strip() != "-"
        }

    @staticmethod
    def _days(value):
        """'25, 30' -> (25, 30); '25' -> (25, 25); '-' or '' -> (None, None)."""

        parts = []

        for part in str(value).split(","):

            match = re.search(r"\d+", part)

            if match:
                parts.append(int(match.group()))

        if not parts:
            return None, None

        return parts[0], parts[1] if len(parts) > 1 else parts[0]

    @staticmethod
    def _year(value):

        text = str(value).strip()

        if text in ("", "-", "nan"):
            return None

        try:
            return int(float(text))
        except ValueError:
            parsed = pd.to_datetime(text, errors="coerce")
            return None if pd.isna(parsed) else parsed.year

    def _long_interval_due(self, token, last_done):

        if last_done is None:
            return True

        return self.friday.year - last_done >= LONG_INTERVALS[token]

    # ==================================================
    # Rules
    # ==================================================

    def _candidate_dates(self, day):
        """The matrix day-of-month in the previous, current and next month
        of the Friday, as real dates (clamped for short months)."""

        candidates = []

        for offset in (-1, 0, 1):

            month_index = self.friday.year * 12 + self.friday.month - 1 + offset

            year, month = divmod(month_index, 12)

            month += 1

            last_day = calendar.monthrange(year, month)[1]

            candidates.append(date(year, month, min(day, last_day)))

        return candidates

    def evaluate(self, row):
        """Returns (frequencies, scheduled_dates) due this Friday."""

        frequencies = []

        dates = []

        friday_month = MONTHS[self.friday.month - 1]

        if "W" in self._tokens(row.get(friday_month, "")):

            frequencies.append("W")

            dates.append(self.friday)

        monthly_day, other_day = self._days(row.get("Maintenance Date", ""))

        last_done = self._year(
            row.get("Last 2Y/6Y Done", row.get("Last 2Y Done", ""))
        )

        wanted = {
            "M": monthly_day,
            "3M": other_day, "6M": other_day, "Y": other_day,
            "2Y": other_day, "6Y": other_day,
        }

        for frequency, day in wanted.items():

            if day is None:
                continue

            if frequency in LONG_INTERVALS and not self._long_interval_due(
                    frequency, last_done):
                continue

            for candidate in self._candidate_dates(day):

                if abs((candidate - self.friday).days) > WINDOW_DAYS:
                    continue

                month_tokens = self._tokens(
                    row.get(MONTHS[candidate.month - 1], ""))

                if frequency in month_tokens:

                    frequencies.append(frequency)

                    dates.append(candidate)

                    break

        frequencies.sort(key=FREQUENCY_ORDER.index)

        return frequencies, dates

    # ==================================================
    # Load / build
    # ==================================================

    @staticmethod
    def _sort_key(equipment_id):

        return [int(part) if part.isdigit() else part.lower()
                for part in re.split(r"(\d+)", str(equipment_id).strip())]

    def build_rows(self):

        if not self.excel_file.exists():
            raise FileNotFoundError(
                f"Equipment workbook not found:\n{self.excel_file}")

        self._report(0, 100, "Reading maintenance matrix...")

        workbook = pd.ExcelFile(self.excel_file)

        sheets = [name for name in workbook.sheet_names if name != "Settings"]

        results = []

        for index, sheet in enumerate(sheets, start=1):

            self._report(index / len(sheets) * 90, 100, f"Checking {sheet}...")

            data = pd.read_excel(workbook, sheet_name=sheet).fillna("")

            data.columns = data.columns.astype(str).str.strip()

            missing = {"Equipment No.", "Maintenance Date"} - set(data.columns)

            if missing:
                raise ValueError(
                    f"Sheet '{sheet}' is missing: {', '.join(sorted(missing))}")

            for _, row in data.iterrows():

                equipment_id = str(row["Equipment No."]).strip()

                if not equipment_id:
                    continue

                frequencies, dates = self.evaluate(row)

                if not frequencies:
                    continue

                scheduled = min(dates)

                results.append({
                    "name": str(row.get("Equipment Name", "")).strip(),
                    "id": equipment_id,
                    "frequency": ", ".join(frequencies),
                    "date": scheduled,
                    "last": scheduled + timedelta(days=GRACE_DAYS),
                })

        results.sort(key=lambda item: self._sort_key(item["id"]))

        return results

    # ==================================================
    # Output
    # ==================================================

    def generate(self):

        rows = self.build_rows()

        self._report(95, 100, "Writing schedule...")

        workbook = Workbook()

        sheet = workbook.active

        sheet.title = "Maintenance Schedule"

        thin = Side(style="thin", color="000000")

        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        sheet.append(HEADERS)

        for number, item in enumerate(rows, start=1):

            sheet.append([number, item["name"], item["id"],
                          item["frequency"], item["date"], item["last"]])

        widths = [8, 46, 26, 20, 16, 16]

        for column, width in enumerate(widths, start=1):

            sheet.column_dimensions[
                sheet.cell(1, column).column_letter].width = width

        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):

            for cell in row:

                cell.border = border

                cell.alignment = Alignment(
                    horizontal="left" if cell.column in (2, 3) else "center",
                    vertical="center", wrap_text=True)

                if cell.row > 1 and cell.column in (5, 6):

                    cell.number_format = "DD-MMM-YYYY"

        for cell in sheet[1]:

            cell.font = Font(bold=True, color="FFFFFF")

            cell.fill = PatternFill("solid", fgColor="1F4E78")

        sheet.freeze_panes = "A2"

        self.output_file.parent.mkdir(parents=True, exist_ok=True)

        workbook.save(self.output_file)

        self._report(100, 100, "Done")

        return rows
