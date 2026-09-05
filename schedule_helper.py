"""
schedule_helper.py

Creates all placeholder values for the Preventive Maintenance Log Card.
"""

from date_helper import get_months, get_weekday_dates


class ScheduleHelper:

    MONTH_NAME_TO_NUMBER = {

        "Jan": 1,
        "Feb": 2,
        "Mar": 3,
        "Apr": 4,
        "May": 5,
        "Jun": 6,
        "Jul": 7,
        "Aug": 8,
        "Sep": 9,
        "Oct": 10,
        "Nov": 11,
        "Dec": 12,

    }

    QUARTERS = {

        "Q1": ["Jan", "Feb", "Mar"],
        "Q2": ["Apr", "May", "Jun"],
        "Q3": ["Jul", "Aug", "Sep"],
        "Q4": ["Oct", "Nov", "Dec"],

    }

    def __init__(

        self,
        equipment_row,
        year,
        quarter,
        weekday

    ):

        self.row = equipment_row

        self.year = year

        self.quarter = quarter

        self.weekday = weekday

        self.placeholders = {}

    # ----------------------------------------------------------
    # Build Everything
    # ----------------------------------------------------------

    def build(self):

        self.placeholders = {}

        self.build_equipment()

        self.build_month_titles()

        self.build_weekly()

        self.build_frequencies()

        return self.placeholders

    # ----------------------------------------------------------
    # Check whether any maintenance exists in this quarter
    # ----------------------------------------------------------

    def has_maintenance(self):

        months = self.QUARTERS[self.quarter]

        for month in months:

            value = str(self.row[month]).strip()

            if value != "-":

                return True

        return False

    # ----------------------------------------------------------
    # Equipment
    # ----------------------------------------------------------

    def build_equipment(self):

        self.placeholders["(Enter Equipment Name)"] = str(
            self.row["Equipment Name"]
        )

        self.placeholders["(Enter Equipment ID)"] = str(
            self.row["Equipment No."]
        )

        self.placeholders["(Enter Equipment Location)"] = str(
            self.row["Location"]
        )

    # ----------------------------------------------------------
    # Month Titles
    # ----------------------------------------------------------

    def build_month_titles(self):

        months = get_months(

            self.year,
            self.quarter

        )

        self.placeholders["(Enter Month1’Year)"] = months[0]

        self.placeholders["(Enter Month2’Year)"] = months[1]

        self.placeholders["(Enter Month3’Year)"] = months[2]

    # ----------------------------------------------------------
    # Weekly Dates
    # ----------------------------------------------------------

    def build_weekly(self):

        weekday_map = {

            "Monday": 0,
            "Tuesday": 1,
            "Wednesday": 2,
            "Thursday": 3,
            "Friday": 4,
            "Saturday": 5,
            "Sunday": 6

        }

        weekday_index = weekday_map[self.weekday]

        months = self.QUARTERS[self.quarter]

        for month_index, month_name in enumerate(months, start=1):

            schedule = str(
                self.row[month_name]
            ).strip()

            # ----------------------------------------
            # No maintenance this month
            # ----------------------------------------

            if schedule == "-":

                for week in range(1, 6):

                    self.placeholders[
                        f"(M{month_index}W{week})"
                    ] = "N/A"

                    self.placeholders[
                        f"(dM{month_index}W{week})"
                    ] = "N/A"

                continue

            month_number = self.MONTH_NAME_TO_NUMBER[
                month_name
            ]

            dates = get_weekday_dates(

                self.year,
                month_number,
                weekday_index

            )

            tokens = [

                item.strip()

                for item in schedule.split(",")

                if item.strip()

            ]

            weekly_enabled = "W" in tokens

            for week in range(5):

                date_placeholder = f"(M{month_index}W{week+1})"

                sign_placeholder = f"(dM{month_index}W{week+1})"

                date = dates[week]

                self.placeholders[
                    date_placeholder
                ] = date

                if date == "N/A":

                    self.placeholders[
                        sign_placeholder
                    ] = "N/A"

                elif weekly_enabled:

                    self.placeholders[
                        sign_placeholder
                    ] = ""

                else:

                    self.placeholders[
                        sign_placeholder
                    ] = "N/A"   
    # ----------------------------------------------------------
    # Safe Value Parsing
    #
    # Equipment.xlsx uses "-" (and sometimes blank cells) to mean
    # "not set" / "never done". int("-") used to crash the whole
    # generator. These helpers parse safely instead.
    # ----------------------------------------------------------

    @staticmethod
    def _safe_maintenance_date(value):

        text = str(value).strip()

        try:

            return f"{int(float(text)):02d}"

        except (TypeError, ValueError):

            # No usable maintenance date on file (e.g. "-").
            # Fall back to day "01" instead of crashing.
            return "01"

    @staticmethod
    def _safe_year(value):

        text = str(value).strip()

        if text in ("", "-"):

            return None

        try:

            return int(float(text))

        except (TypeError, ValueError):

            return None

    # ----------------------------------------------------------
    # Monthly / Quarterly / Half Yearly / Yearly / Two Yearly
    # ----------------------------------------------------------

    def build_frequencies(self):

        maintenance_date = self._safe_maintenance_date(
            self.row["Maintenance Date"]
        )

        last_2y = self._safe_year(
            self.row["Last 2Y Done"]
        )

        months = self.QUARTERS[self.quarter]

        frequency_map = {

            "M": ("M", "dM"),
            "3M": ("Q", "dQ"),
            "6M": ("H", "dH"),
            "Y": ("Y", "dY"),
            "2Y": ("T", "dT")

        }

        for month_index, month_name in enumerate(months, start=1):

            schedule = str(
                self.row[month_name]
            ).strip()

            # ----------------------------------------
            # No maintenance this month
            # ----------------------------------------

            if schedule == "-":

                for value_prefix, sign_prefix in frequency_map.values():

                    self.placeholders[
                        f"({value_prefix}{month_index})"
                    ] = "N/A"

                    self.placeholders[
                        f"({sign_prefix}{month_index})"
                    ] = "N/A"

                continue

            # ----------------------------------------
            # Split schedule into tokens
            # Example:
            # M, 3M, Y
            # ----------------------------------------

            tokens = [

                token.strip()

                for token in schedule.split(",")

                if token.strip()

            ]

            # ----------------------------------------
            # Monthly
            # ----------------------------------------

            if "M" in tokens:

                self.placeholders[
                    f"(M{month_index})"
                ] = maintenance_date

                self.placeholders[
                    f"(dM{month_index})"
                ] = ""

            else:

                self.placeholders[
                    f"(M{month_index})"
                ] = "N/A"

                self.placeholders[
                    f"(dM{month_index})"
                ] = "N/A"

            # ----------------------------------------
            # Quarterly
            # ----------------------------------------

            if "3M" in tokens:

                self.placeholders[
                    f"(Q{month_index})"
                ] = maintenance_date

                self.placeholders[
                    f"(dQ{month_index})"
                ] = ""

            else:

                self.placeholders[
                    f"(Q{month_index})"
                ] = "N/A"

                self.placeholders[
                    f"(dQ{month_index})"
                ] = "N/A"

            # ----------------------------------------
            # Half Yearly
            # ----------------------------------------

            if "6M" in tokens:

                self.placeholders[
                    f"(H{month_index})"
                ] = maintenance_date

                self.placeholders[
                    f"(dH{month_index})"
                ] = ""

            else:

                self.placeholders[
                    f"(H{month_index})"
                ] = "N/A"

                self.placeholders[
                    f"(dH{month_index})"
                ] = "N/A"

            # ----------------------------------------
            # Yearly
            # ----------------------------------------

            if "Y" in tokens:

                self.placeholders[
                    f"(Y{month_index})"
                ] = maintenance_date

                self.placeholders[
                    f"(dY{month_index})"
                ] = ""

            else:

                self.placeholders[
                    f"(Y{month_index})"
                ] = "N/A"

                self.placeholders[
                    f"(dY{month_index})"
                ] = "N/A"

            # ----------------------------------------
            # Two Yearly
            # ----------------------------------------

            due = last_2y is None or (self.year - last_2y) >= 2

            if "2Y" in tokens and due:

                self.placeholders[
                    f"(T{month_index})"
                ] = maintenance_date

                self.placeholders[
                    f"(dT{month_index})"
                ] = ""

            else:

                self.placeholders[
                    f"(T{month_index})"
                ] = "N/A"

                self.placeholders[
                    f"(dT{month_index})"
                ] = "N/A"