"""
ui_common.py

Shared UI helpers: high-DPI support and thread-safe progress tracking.
"""

import ctypes
import threading

_scale = 1.0


def enable_high_dpi():
    """Make the process DPI aware so Tk and the native file dialogs render
    crisp instead of being bitmap-stretched by Windows. Call before Tk()."""

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def init_scaling(root):
    """Sync Tk's point scaling with the real monitor DPI. Call after Tk()."""

    global _scale

    try:
        dpi = root.winfo_fpixels("1i")
    except Exception:
        dpi = 96.0

    root.tk.call("tk", "scaling", dpi / 72.0)

    _scale = max(dpi / 96.0, 1.0)


def px(value):
    """Scale a 96-DPI pixel value to the current display."""

    return int(round(value * _scale))


class ProgressTracker:
    """Worker threads call update(); the GUI thread polls snapshot().

    Avoids calling Tk from the worker thread and lets the progress bar
    keep moving while the worker is busy."""

    def __init__(self):

        self._lock = threading.Lock()

        self._value = 0.0

        self._message = ""

    def update(self, current, total, message=""):

        with self._lock:

            self._value = (current / total * 100.0) if total else 0.0

            if message:

                self._message = message

    def snapshot(self):

        with self._lock:

            return self._value, self._message


class ProgressPanel:
    """Status strip with a determinate progress bar, percentage and message.

    The bar eases toward the worker's real progress so it animates smoothly
    even when the worker thread hogs the interpreter between updates."""

    def __init__(self, parent, root, status_var):

        import tkinter as tk
        from tkinter import ttk

        self.root = root

        self.status = status_var

        self.tracker = None

        self.target = 0.0

        self.shown = 0.0

        self.running = False

        self.frame = tk.Frame(
            parent, bg="#EAF3F8",
            highlightbackground="#D2E4EE", highlightthickness=1
        )

        style = ttk.Style(root)

        style.configure(
            "Suite.Horizontal.TProgressbar",
            troughcolor="#D7E8F1", background="#167D9A",
            bordercolor="#D7E8F1", lightcolor="#167D9A",
            darkcolor="#167D9A", thickness=px(12)
        )

        self.percent = tk.StringVar(value="")

        self.bar = ttk.Progressbar(
            self.frame, mode="determinate", maximum=100,
            length=px(240), style="Suite.Horizontal.TProgressbar"
        )

        self.bar.pack(side="right", padx=(6, 16), pady=13)

        tk.Label(
            self.frame, textvariable=self.percent, bg="#EAF3F8",
            fg="#34536B", width=5, anchor="e",
            font=("Segoe UI", 10, "bold")
        ).pack(side="right")

        tk.Label(
            self.frame, textvariable=self.status, bg="#EAF3F8",
            fg="#34536B", font=("Segoe UI", 10)
        ).pack(side="left", padx=16, pady=13)

    def pack(self, **options):

        self.frame.pack(**options)

    def start(self, tracker):

        self.tracker = tracker

        self.target = 0.0

        self.shown = 0.0

        self.bar["value"] = 0

        self.percent.set("0%")

        self.running = True

        self._poll()

    def _poll(self):

        if not self.running:

            return

        value, message = self.tracker.snapshot()

        self.target = value

        self.shown += (self.target - self.shown) * 0.3

        if self.target - self.shown < 0.3:

            self.shown = self.target

        self.bar["value"] = self.shown

        self.percent.set(f"{int(self.shown)}%")

        if message:

            self.status.set(message)

        self.root.after(40, self._poll)

    def stop(self, complete=True):

        self.running = False

        if complete:

            self.bar["value"] = 100

            self.percent.set("100%")

        else:

            self.bar["value"] = 0

            self.percent.set("")
