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
        sheet_name="EN&WH"
    ):

        self.template = template_path
        self.excel = excel_path
        self.output = output_file

        self.year = year
        self.quarter = quarter
        self.weekday = weekday
        self.sheet_name = sheet_name

        self.logger = Logger()

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

            composer = None

            first_document = True

            generated = 0

            skipped = 0

            # --------------------------------------------------
            # Generate Log Cards
            # --------------------------------------------------

            for _, row in equipment.iterrows():

                schedule = ScheduleHelper(

                    equipment_row=row,

                    year=self.year,

                    quarter=self.quarter,

                    weekday=self.weekday

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

            composer.save(

                self.output

            )

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