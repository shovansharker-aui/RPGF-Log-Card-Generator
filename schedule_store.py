"""
schedule_store.py

Remembers the last generated weekly schedule (and which machines were
ticked) so it is still there when the app or the module is opened again.

Stored in the user's profile, not next to the exe, so rebuilding or
replacing the app folder never loses it.
"""

import json
import os
from datetime import date, datetime
from pathlib import Path


def store_file():

    base = os.environ.get("APPDATA") or str(Path.home())

    folder = Path(base) / "RPGF Engineering Suite"

    folder.mkdir(parents=True, exist_ok=True)

    return folder / "last_schedule.json"


def _iso(value):

    return value.isoformat()


def save(rows, friday, pdf_file, excel_file, checked, generated=None):

    data = {
        "version": 1,
        "friday": _iso(friday),
        "generated": _iso(generated or date.today()),
        "pdf": str(pdf_file) if pdf_file else "",
        "excel": str(excel_file) if excel_file else "",
        "checked": sorted(checked),
        "rows": [
            {
                "name": row["name"],
                "id": row["id"],
                "frequency": row["frequency"],
                "date": _iso(row["date"]),
                "items": [
                    {
                        "frequency": item["frequency"],
                        "date": _iso(item["date"]),
                        "last": _iso(item["last"]),
                    }
                    for item in row["items"]
                ],
            }
            for row in rows
        ],
    }

    target = store_file()

    temp = target.with_suffix(".tmp")

    temp.write_text(json.dumps(data, indent=2), encoding="utf-8")

    temp.replace(target)


def load():
    """Returns the saved schedule or None when there is none / unreadable."""

    try:
        data = json.loads(store_file().read_text(encoding="utf-8"))

        rows = [
            {
                "name": row["name"],
                "id": row["id"],
                "frequency": row["frequency"],
                "date": date.fromisoformat(row["date"]),
                "items": [
                    {
                        "frequency": item["frequency"],
                        "date": date.fromisoformat(item["date"]),
                        "last": date.fromisoformat(item["last"]),
                    }
                    for item in row["items"]
                ],
            }
            for row in data["rows"]
        ]

        return {
            "friday": date.fromisoformat(data["friday"]),
            "generated": date.fromisoformat(data["generated"]),
            "pdf": data.get("pdf", ""),
            "excel": data.get("excel", ""),
            "checked": {
                index for index in data.get("checked", [])
                if isinstance(index, int) and 0 <= index < len(rows)
            },
            "rows": rows,
        }

    except Exception:
        return None
