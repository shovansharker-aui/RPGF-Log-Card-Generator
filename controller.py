"""
controller.py

Main application controller
"""

from docxcompose.composer import Composer

from excel_reader import ExcelReader
from schedule_helper import ScheduleHelper
from word_builder import WordBuilder
from logger import Logger


class Controller:

    def __init__(
        self,
        template_path,
        excel_path,
        output_file,
        year,
        quarter,
        weekday,
        sheet_name="EN&WH",
        progress=None
    ):

        self.template = template_path
        self.excel = excel_path
        self.output = output_file

        self.year = year
        self.quarter = quarter
        self.weekday = weekday
        self.sheet_name = sheet_name

        self.progress = progress

        self.logger = Logger()

        # "Equipment Name - Equipment No." entries with 6Y maintenance due
        self.six_year_due = []

    def _report(self, current, total, message):

        if self.progress:

            self.progress(current, total, message)

    # ==========================================================
    # Main
    # ==========================================================

    def run(self):

        try:

            # --------------------------------------------------
            # Read Excel
            # --------------------------------------------------

            reader = ExcelReader(self.excel)

            reader.load(self.sheet_name)

            reader.validate()

            reader.clean()

            equipment = reader.get_equipment()

            self.logger.info(

                f"Loaded {reader.count()} equipment from '{self.sheet_name}'."

            )

            self.six_year_due = []

            composer = None

            first_document = True

            generated = 0

            skipped = 0

            # --------------------------------------------------
            # Generate Log Cards
            # --------------------------------------------------

            total = len(equipment)

            self._report(0, total, "Reading equipment...")

            for position, (_, row) in enumerate(equipment.iterrows(), start=1):

                self._report(
                    position - 1,
                    total,
                    f"Processing {row['Equipment No.']} ({position}/{total})"
                )

                schedule = ScheduleHelper(

                    equipment_row=row,

                    year=self.year,

                    quarter=self.quarter,

                    weekday=self.weekday

                )

                if schedule.has_six_year_due():

                    self.six_year_due.append(
                        f"{row['Equipment Name']} - {row['Equipment No.']}"
                    )

                if not schedule.has_maintenance():

                    skipped += 1

                    self.logger.info(

                        f"Skipped : {row['Equipment No.']}"

                    )

                    continue

                placeholders = schedule.build()

                builder = WordBuilder(

                    self.template

                )

                builder.load()

                builder.replace_all(

                    placeholders

                )

                if first_document:

                    composer = Composer(

                        builder.doc

                    )

                    first_document = False

                else:

                    composer.append(

                        builder.doc

                    )

                generated += 1

                self.logger.info(

                    f"Generated : {row['Equipment No.']}"

                )

            # --------------------------------------------------
            # Save Output
            # --------------------------------------------------

            if composer is None:

                self.logger.info(

                    "No log cards generated."

                )

                return False

            self._report(total, total, "Saving document...")

            composer.save(

                self.output

            )

            self._report(total, total, "Done")

            self.logger.info(

                f"Saved : {self.output}"

            )

            self.logger.info(

                f"Generated={generated}, Skipped={skipped}"

            )

            return True

        except Exception as e:

            self.logger.error(

                str(e)

            )

            raise