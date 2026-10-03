"""
due_controller.py

Maintenance Due label tags: reads the machine list and fills the
three-boxes-per-page Word template.
"""

import re
from pathlib import Path

import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docxcompose.composer import Composer

BOXES_PER_PAGE = 3

# Accepted column headers (first match wins) for each field.
COLUMN_ALIASES = {
    "name": ["Equipment Name", "Name"],
    "id": ["Equipment No.", "Equipment ID", "Equipment No", "ID"],
    "frequency": ["Maintenance Frequency", "Frequency"],
    "date": ["Maintenance Date", "Scheduled Date", "Date"],
    "last": ["Last Allowable Date", "Last Date", "Last Maintenance Date"],
}

PLACEHOLDER = re.compile(r"\((name|id|fr|s_date|l_d)\)")


class DueController:

    # ==================================================
    # Machine list
    # ==================================================

    @staticmethod
    def sheet_names(excel_file):

        return [
            name for name in pd.ExcelFile(excel_file).sheet_names
            if name != "Settings"
        ]

    @staticmethod
    def _format(value):

        if value is None or (not isinstance(value, str) and pd.isna(value)):

            return ""

        if isinstance(value, pd.Timestamp):

            return value.strftime("%d-%b-%Y")

        text = str(value).strip()

        return "" if text == "-" else text

    @classmethod
    def load_machines(cls, excel_file, sheet_names=None):

        machines = []

        for sheet in sheet_names or cls.sheet_names(excel_file):

            data = pd.read_excel(excel_file, sheet_name=sheet)

            data.columns = data.columns.astype(str).str.strip()

            columns = {}

            for field, aliases in COLUMN_ALIASES.items():

                columns[field] = next(
                    (alias for alias in aliases if alias in data.columns),
                    None
                )

            if columns["name"] is None or columns["id"] is None:

                raise ValueError(
                    f"Sheet '{sheet}' needs 'Equipment Name' and "
                    "'Equipment No.' columns."
                )

            for _, row in data.iterrows():

                machine = {
                    field: cls._format(row[column]) if column else ""
                    for field, column in columns.items()
                }

                if machine["id"] or machine["name"]:

                    machines.append(machine)

        return machines

    # ==================================================
    # Template filling
    # ==================================================

    @staticmethod
    def _boxes(document):

        """One box per row of the template's outer table (3 per page)."""

        outer = document.element.body.find(qn("w:tbl"))

        if outer is None:

            raise ValueError("Template has no label table.")

        return [
            cell
            for row in outer.findall(qn("w:tr"))
            for cell in row.findall(qn("w:tc"))
        ]

    @staticmethod
    def _fill_paragraph(paragraph, values):

        runs = [
            text for run in paragraph.iter(qn("w:r"))
            for text in run.findall(qn("w:t"))
        ]

        full = "".join(text.text or "" for text in runs)

        matches = list(PLACEHOLDER.finditer(full))

        if not matches:

            return

        # Work back to front so earlier offsets stay valid.
        for match in reversed(matches):

            value = values.get(match.group(0), "")

            position = 0

            placed = False

            for text in runs:

                original = text.text or ""

                start, end = position, position + len(original)

                position = end

                if end <= match.start() or start >= match.end():

                    continue

                cut_from = max(match.start() - start, 0)

                cut_to = min(match.end() - start, len(original))

                insert = "" if placed else value

                text.text = original[:cut_from] + insert + original[cut_to:]

                text.set(
                    "{http://www.w3.org/XML/1998/namespace}space",
                    "preserve"
                )

                placed = True

    @classmethod
    def _fill_page(cls, template, page_machines):

        document = Document(template)

        boxes = cls._boxes(document)

        for index, box in enumerate(boxes):

            machine = (
                page_machines[index] if index < len(page_machines) else None
            )

            values = {
                "(name)": machine["name"] if machine else "",
                "(id)": machine["id"] if machine else "",
                "(fr)": machine["frequency"] if machine else "",
                "(s_date)": machine["date"] if machine else "",
                "(l_d)": machine["last"] if machine else "",
            }

            for paragraph in box.iter(qn("w:p")):

                cls._fill_paragraph(paragraph, values)

        return document

    # ==================================================
    # Generate
    # ==================================================

    def generate(self, template, machines, output, progress=None):

        if not Path(template).exists():

            raise FileNotFoundError(f"Template not found:\n{template}")

        if not machines:

            raise ValueError("No machines selected.")

        pages = [
            machines[i:i + BOXES_PER_PAGE]
            for i in range(0, len(machines), BOXES_PER_PAGE)
        ]

        composer = None

        for number, page_machines in enumerate(pages, start=1):

            if progress:

                progress(number - 1, len(pages) + 1,
                         f"Filling page {number} of {len(pages)}...")

            document = self._fill_page(template, page_machines)

            if composer is None:

                composer = Composer(document)

            else:

                composer.append(document)

        if progress:

            progress(len(pages), len(pages) + 1, "Saving document...")

        Path(output).parent.mkdir(parents=True, exist_ok=True)

        composer.save(output)

        if progress:

            progress(len(pages) + 1, len(pages) + 1, "Done")

        return output
