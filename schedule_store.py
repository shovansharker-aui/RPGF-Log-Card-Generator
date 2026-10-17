"""
schedule_store.py

Remembers which machines were ticked in the schedule list, so the ticks are
still there when the app or the module is opened again.

The schedule itself is not stored here: it is read from the latest
"Maintenance Schedule_*.xlsx" in the output folder.

Stored in the user's profile, not next to the exe, so rebuilding or
replacing the app folder never loses it.
"""

import json
import os
from pathlib import Path


def store_file():

    base = os.environ.get("APPDATA") or str(Path.home())

    folder = Path(base) / "RPGF Engineering Suite"

    folder.mkdir(parents=True, exist_ok=True)

    return folder / "ticked_machines.json"


def save_ticks(excel_file, equipment_ids):
    """Remember the ticked Equipment IDs for this schedule file."""

    data = {
        "excel": str(excel_file) if excel_file else "",
        "ids": sorted(equipment_ids),
    }

    target = store_file()

    temp = target.with_suffix(".tmp")

    temp.write_text(json.dumps(data, indent=2), encoding="utf-8")

    temp.replace(target)


def load_ticks(excel_file):
    """Ticked Equipment IDs saved for this exact schedule file, else empty."""

    try:
        data = json.loads(store_file().read_text(encoding="utf-8"))

        if str(excel_file) == data.get("excel"):
            return set(data.get("ids", []))

    except Exception:
        pass

    return set()
