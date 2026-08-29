"""
trend_gui.py

Trend Report Generator
RPGF Engineering Suite v2.0
"""

import json
import threading
import sys
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk

from config import DEFAULT_TREND_NAME, OUTPUT_DIR
from trend_controller import TrendController


class TrendWindow:

    def __init__(self, dashboard=None):

        self.dashboard = dashboard

        if dashboard is None:
            self.root = tk.Tk()
        else:
            self.root = tk.Toplevel(dashboard)

        self.root.title("RPGF Trend Report Generator v2.0")
        self.root.geometry("920x650")
        self.root.resizable(False, False)

        self.settings_file = "settings.json"

        self.logo = self.load_logo()

        self.create_variables()
        self.configure_styles()
        self.build_ui()
        self.load_settings()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

        if dashboard is None:
            self.root.mainloop()

    def resource_path(self, relative_path):

        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = Path(__file__).parent

        return str(Path(base_path) / relative_path)

    def load_logo(self):

        try:
            image = Image.open(
                self.resource_path("assets/renata.png")
            )

            image = image.resize(
                (70, 70),
                Image.LANCZOS
            )

            return ImageTk.PhotoImage(image)

        except Exception:
            return None

    def create_variables(self):

        self.excel_path = tk.StringVar()

        self.output_path = tk.StringVar()

        self.year = tk.IntVar(value=2026)

        self.half = tk.StringVar(value="Jul-Dec")

        self.weekday = tk.StringVar(value="Friday")

        self.status = tk.StringVar(value="Ready")

    def save_settings(self):

        data = {}

        if Path(self.settings_file).exists():
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}

        data["trend_excel"] = self.excel_path.get()
        data["trend_output"] = self.output_path.get()
        data["trend_year"] = self.year.get()
        data["trend_half"] = self.half.get()
        data["trend_weekday"] = self.weekday.get()

        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

    def load_settings(self):

        if not Path(self.settings_file).exists():
            return

        try:
            with open(self.settings_file, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.excel_path.set(data.get("trend_excel", ""))
            self.output_path.set(data.get("trend_output", ""))
            self.year.set(data.get("trend_year", 2026))
            self.half.set(data.get("trend_half", "Jul-Dec"))
            self.weekday.set(data.get("trend_weekday", "Friday"))

        except Exception:
            pass

        if not self.output_path.get():
            self.output_path.set(str(OUTPUT_DIR / DEFAULT_TREND_NAME))

    def build_ui(self):
        self.root.configure(bg="#F4F7FB")
        main = tk.Frame(self.root, bg="#F4F7FB")
        main.pack(fill="both", expand=True)

        header = tk.Frame(main, bg="#123B5D", height=132)
        header.pack(fill="x")
        header.pack_propagate(False)
        if self.logo:
            tk.Label(header, image=self.logo, bg="#123B5D").place(x=40, y=28)
        tk.Label(header, text="Trend Report Generator", bg="#123B5D", fg="white",
                 font=("Segoe UI", 22, "bold")).place(x=135, y=33)

        content = tk.Frame(main, bg="#F4F7FB")
        content.pack(fill="both", expand=True, padx=42, pady=24)

        card = tk.Frame(content, bg="white", highlightbackground="#DCE4EC", highlightthickness=1)
        card.pack(fill="x")
        tk.Label(card, text="Report setup", bg="white", fg="#123B5D",
                 font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=3,
                                                        sticky="w", padx=24, pady=(20, 3))
        tk.Label(card, text="Select the source workbook, reporting period, and output location.",
                 bg="white", fg="#6B7785", font=("Segoe UI", 9)).grid(
                     row=1, column=0, columnspan=3, sticky="w", padx=24, pady=(0, 18))

        form = tk.Frame(card, bg="white")
        form.grid(row=2, column=0, columnspan=3, sticky="ew", padx=24, pady=(0, 22))
        form.columnconfigure(1, weight=1)
        self.create_path_field(form, 0, "Equipment workbook", self.excel_path, self.browse_excel)
        self.create_path_field(form, 1, "Output workbook", self.output_path, self.browse_output)

        settings = tk.Frame(card, bg="white")
        settings.grid(row=3, column=0, columnspan=3, sticky="ew", padx=24, pady=(0, 24))
        for column in range(3):
            settings.columnconfigure(column, weight=1)
        self.create_combo_field(settings, 0, "Report year", self.year, [2025, 2026, 2027, 2028, 2029, 2030])
        self.create_combo_field(settings, 1, "Reporting period", self.half, ["Jan-Jun", "Jul-Dec"])
        self.create_combo_field(settings, 2, "Weekly maintenance day", self.weekday,
                                ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])

        action = tk.Frame(content, bg="#F4F7FB")
        action.pack(fill="x", pady=(18, 0))
        tk.Button(action, text="Generate Trend Report", command=self.start_generation,
                  bg="#167D9A", fg="white", activebackground="#10667E", activeforeground="white",
                  relief="flat", cursor="hand2", font=("Segoe UI", 11, "bold"), padx=25, pady=10).pack(side="left")
        tk.Button(action, text="Back to Dashboard", command=self.go_back, bg="#F4F7FB", fg="#34536B",
                  activebackground="#E3EAF0", relief="flat", cursor="hand2", font=("Segoe UI", 10),
                  padx=12, pady=10).pack(side="right")

        status_card = tk.Frame(content, bg="#EAF3F8", highlightbackground="#D2E4EE", highlightthickness=1)
        status_card.pack(fill="x", pady=(18, 0))
        self.progress = ttk.Progressbar(status_card, mode="indeterminate", length=180, style="Trend.Horizontal.TProgressbar")
        self.progress.pack(side="right", padx=16, pady=13)
        tk.Label(status_card, textvariable=self.status, bg="#EAF3F8", fg="#34536B",
                 font=("Segoe UI", 10)).pack(side="left", padx=16, pady=13)

    def configure_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("Trend.TEntry", padding=8, fieldbackground="white", bordercolor="#B8C6D1")
        style.configure("Trend.TCombobox", padding=7, fieldbackground="white", bordercolor="#B8C6D1")
        style.configure("Trend.Horizontal.TProgressbar", troughcolor="#D7E8F1", background="#167D9A",
                        bordercolor="#D7E8F1", lightcolor="#167D9A", darkcolor="#167D9A")

    def create_path_field(self, parent, row, label, variable, command):
        tk.Label(parent, text=label, bg="white", fg="#263847", font=("Segoe UI", 10, "bold")).grid(
            row=row, column=0, sticky="w", pady=7)
        ttk.Entry(parent, textvariable=variable, style="Trend.TEntry").grid(
            row=row, column=1, sticky="ew", padx=(18, 10), pady=7, ipady=2)
        tk.Button(parent, text="Browse", command=command, bg="#E7F0F5", fg="#1D516B",
                  activebackground="#D4E5EF", relief="flat", cursor="hand2", font=("Segoe UI", 9, "bold"),
                  padx=15, pady=6).grid(row=row, column=2, pady=7)

    def create_combo_field(self, parent, column, label, variable, values):
        field = tk.Frame(parent, bg="white")
        field.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 12, 0))
        tk.Label(field, text=label, bg="white", fg="#263847", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Combobox(field, textvariable=variable, values=values, state="readonly",
                     style="Trend.TCombobox").pack(fill="x", pady=(6, 0), ipady=2)

    def browse_excel(self):

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if file_path:
            self.excel_path.set(file_path)

    def browse_output(self):

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if file_path:
            self.output_path.set(file_path)

    def start_generation(self):
        if not self.excel_path.get():
            messagebox.showerror("Trend Report", "Please select the Equipment workbook.")
            return
        if not self.output_path.get():
            messagebox.showerror("Trend Report", "Please select an output workbook.")
            return

        self.save_settings()
        self.status.set("Generating trend report...")
        self.progress.start()
        threading.Thread(target=self.generate, daemon=True).start()

    def generate(self):
        try:
            output = TrendController(
                self.excel_path.get(), self.output_path.get(), self.year.get(),
                self.half.get(), self.weekday.get()
            ).generate()
        except Exception as error:
            self.root.after(0, self.generation_failed, str(error))
            return
        self.root.after(0, self.generation_finished, str(output))

    def generation_finished(self, output):
        self.progress.stop()
        self.status.set("Trend report created successfully.")
        messagebox.showinfo("Trend Report", f"Trend report saved to:\n{output}")

    def generation_failed(self, error):
        self.progress.stop()
        self.status.set("Trend report generation failed.")
        messagebox.showerror("Trend Report", error)

    def go_back(self):

        self.save_settings()

        self.root.destroy()

        if self.dashboard:
            self.dashboard.deiconify()

    def on_close(self):

        self.save_settings()

        self.root.destroy()

        if self.dashboard:
            self.dashboard.deiconify()


if __name__ == "__main__":
    TrendWindow()
