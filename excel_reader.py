"""
excel_reader.py

Reads and validates the Equipment workbook.
"""

from pathlib import Path

import pandas as pd


class ExcelReader:

    REQUIRED_COLUMNS = [

        "Equipment Name",

        "Equipment No.",

        "Location",

        "Maintenance Date",

        "Last 2Y Done"

    ]

    def __init__(self, excel_file):

        self.excel_file = Path(excel_file)

        self.df = None

    # ==================================================
    # Load Workbook
    # ==================================================

    def load(self, sheet_name="Equipment"):

        if not self.excel_file.exists():

            raise FileNotFoundError(

                f"Excel file not found:\n{self.excel_file}"

            )

        self.df = pd.read_excel(

            self.excel_file,

            sheet_name=sheet_name

        )

    # ==================================================
    # Validate Workbook
    # ==================================================

    def validate(self):

        if self.df is None:

            raise RuntimeError(

                "Workbook has not been loaded."

            )

        missing = [

            column

            for column in self.REQUIRED_COLUMNS

            if column not in self.df.columns

        ]

        if missing:

            raise ValueError(

                "Missing required columns:\n"

                + "\n".join(missing)

            )

        if self.df.empty:

            raise ValueError(

                "Equipment sheet is empty."

            )

    # ==================================================
    # Clean Data
    # ==================================================

    def clean(self):

        self.df = self.df.fillna("")

        self.df = self.df.astype(str)

        self.df.columns = self.df.columns.str.strip()

        self.df = self.df.apply(

            lambda column: column.str.strip()

        )

    # ==================================================
    # Get Equipment
    # ==================================================

    def get_equipment(self):

        return self.df.copy()

    # ==================================================
    # Statistics
    # ==================================================

    def count(self):

        return len(self.df)

    # ==================================================
    # Workbook Information
    # ==================================================

    def columns(self):

        return self.df.columns.tolist()

    def sheets(self):

        return pd.ExcelFile(

            self.excel_file

        ).sheet_names