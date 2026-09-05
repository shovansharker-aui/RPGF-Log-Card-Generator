"""
logcard_gui.py

Log Card Generator
RPGF Engineering Suite v2.0
"""

import json
import threading
import sys
from pathlib import Path

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

from PIL import Image
from PIL import ImageTk

from controller import Controller


class MainWindow:

    # ==========================================================
    # Constructor
    # ==========================================================

    def __init__(self, dashboard=None):

        self.dashboard = dashboard

        if dashboard is None:

            self.root = tk.Tk()

        else:

            self.root = tk.Toplevel(dashboard)

        self.root.title(
            "RPGF Log Card Generator v2.0"
        )

        self.root.geometry("920x650")

        self.root.resizable(False, False)

        self.settings_file = "settings.json"

        self.logo = self.load_logo()

        self.create_variables()
        self.configure_styles()
        self.build_modern_ui()

        self.load_settings()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

        if dashboard is None:

            self.root.mainloop()

    # ==========================================================
    # Resource Path
    # ==========================================================

    def resource_path(self, relative_path):

        try:

            base_path = sys._MEIPASS

        except Exception:

            base_path = Path(__file__).parent

        return str(

            Path(base_path) / relative_path

        )

    # ==========================================================
    # Load Logo
    # ==========================================================

    def load_logo(self):

        try:

            image = Image.open(

                self.resource_path(

                    "assets/renata.png"

                )

            )

            image = image.resize(

                (70, 70),

                Image.LANCZOS

            )

            return ImageTk.PhotoImage(image)

        except Exception:

            return None

    # ==========================================================
    # Variables
    # ==========================================================

    def create_variables(self):

        self.excel_path = tk.StringVar()

        self.template_path = tk.StringVar()

        self.output_path = tk.StringVar()

        self.sheet_name = tk.StringVar(
            value="EN&WH"
        )

        self.year = tk.IntVar(
            value=2026
        )

        self.quarter = tk.StringVar(
            value="Q3"
        )

        self.weekday = tk.StringVar(
            value="Friday"
        )

        self.status = tk.StringVar(
            value="Ready"
        )

    # ==========================================================
    # Save Settings
    # ==========================================================

    def save_settings(self):

        data = {

            "excel": self.excel_path.get(),

            "template": self.template_path.get(),

            "output": self.output_path.get(),

            "sheet": self.sheet_name.get(),

            "year": self.year.get(),

            "quarter": self.quarter.get(),

            "weekday": self.weekday.get()

        }

        with open(

            self.settings_file,

            "w",

            encoding="utf-8"

        ) as file:

            json.dump(

                data,

                file,

                indent=4

            )

    # ==========================================================
    # Load Settings
    # ==========================================================

    def load_settings(self):

        settings = Path(

            self.settings_file

        )

        if not settings.exists():

            return

        try:

            with open(

                settings,

                "r",

                encoding="utf-8"

            ) as file:

                data = json.load(file)

            self.excel_path.set(

                data.get("excel", "")

            )

            self.template_path.set(

                data.get("template", "")

            )

            self.output_path.set(

                data.get("output", "")

            )

            self.sheet_name.set(

                data.get(

                    "sheet",

                    "EN&WH"

                )

            )

            self.year.set(

                data.get(

                    "year",

                    2026

                )

            )

            self.quarter.set(

                data.get(

                    "quarter",

                    "Q3"

                )

            )

            self.weekday.set(

                data.get(

                    "weekday",

                    "Friday"

                )

            )

        except Exception:

            pass

    # ==========================================================
    # Build User Interface
    # ==========================================================

    def configure_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("Suite.TEntry", padding=8, fieldbackground="white", bordercolor="#B8C6D1")
        style.configure("Suite.TCombobox", padding=7, fieldbackground="white", bordercolor="#B8C6D1")
        style.configure("Suite.Horizontal.TProgressbar", troughcolor="#D7E8F1", background="#167D9A",
                        bordercolor="#D7E8F1", lightcolor="#167D9A", darkcolor="#167D9A")

    def build_modern_ui(self):
        self.root.configure(bg="#F4F7FB")
        main = tk.Frame(self.root, bg="#F4F7FB")
        main.pack(fill="both", expand=True)
        header = tk.Frame(main, bg="#123B5D", height=132)
        header.pack(fill="x")
        header.pack_propagate(False)
        if self.logo:
            tk.Label(header, image=self.logo, bg="#123B5D").place(x=40, y=28)
        tk.Label(header, text="Log Card Generator", bg="#123B5D", fg="white",
                 font=("Segoe UI", 22, "bold")).place(x=135, y=45)

        content = tk.Frame(main, bg="#F4F7FB")
        content.pack(fill="both", expand=True, padx=42, pady=24)
        card = tk.Frame(content, bg="white", highlightbackground="#DCE4EC", highlightthickness=1)
        card.pack(fill="x")
        tk.Label(card, text="Generation setup", bg="white", fg="#123B5D", font=("Segoe UI", 14, "bold")).pack(
            anchor="w", padx=24, pady=(20, 3))
        tk.Label(card, text="Choose the equipment data, log-card template, and reporting period.", bg="white",
                 fg="#6B7785", font=("Segoe UI", 9)).pack(anchor="w", padx=24, pady=(0, 14))

        form = tk.Frame(card, bg="white")
        form.pack(fill="x", padx=24)
        form.columnconfigure(1, weight=1)
        self.create_path_field(form, 0, "Equipment workbook", self.excel_path, self.browse_excel)
        self.create_path_field(form, 1, "Word template", self.template_path, self.browse_template)
        self.create_path_field(form, 2, "Output document", self.output_path, self.browse_output)

        choices = tk.Frame(card, bg="white")
        choices.pack(fill="x", padx=24, pady=(12, 24))
        for column in range(4):
            choices.columnconfigure(column, weight=1)
        self.create_combo_field(choices, 0, "Equipment sheet", self.sheet_name, ["EN&WH", "PR Part 1", "PR Part 2"])
        self.create_combo_field(choices, 1, "Report year", self.year, list(range(2024, 2041)))
        self.create_combo_field(choices, 2, "Quarter", self.quarter, ["Q1", "Q2", "Q3", "Q4"])
        self.create_combo_field(choices, 3, "Weekly maintenance day", self.weekday,
                                ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])

        actions = tk.Frame(content, bg="#F4F7FB")
        actions.pack(fill="x", pady=(18, 0))
        self.generate_button = tk.Button(actions, text="Generate Log Cards", command=self.generate,
                                         bg="#167D9A", fg="white", activebackground="#10667E",
                                         activeforeground="white", relief="flat", cursor="hand2",
                                         font=("Segoe UI", 11, "bold"), padx=25, pady=10)
        self.generate_button.pack(side="left")
        tk.Button(actions, text="Back to Dashboard", command=self.go_back, bg="#F4F7FB", fg="#34536B",
                  activebackground="#E3EAF0", relief="flat", cursor="hand2", font=("Segoe UI", 10),
                  padx=12, pady=10).pack(side="right")
        status = tk.Frame(content, bg="#EAF3F8", highlightbackground="#D2E4EE", highlightthickness=1)
        status.pack(fill="x", pady=(18, 0))
        self.progress = ttk.Progressbar(status, mode="indeterminate", length=180, style="Suite.Horizontal.TProgressbar")
        self.progress.pack(side="right", padx=16, pady=13)
        tk.Label(status, textvariable=self.status, bg="#EAF3F8", fg="#34536B", font=("Segoe UI", 10)).pack(
            side="left", padx=16, pady=13)

    def create_path_field(self, parent, row, label, variable, command):
        tk.Label(parent, text=label, bg="white", fg="#263847", font=("Segoe UI", 10, "bold")).grid(
            row=row, column=0, sticky="w", pady=7)
        ttk.Entry(parent, textvariable=variable, style="Suite.TEntry").grid(
            row=row, column=1, sticky="ew", padx=(18, 10), pady=7, ipady=2)
        tk.Button(parent, text="Browse", command=command, bg="#E7F0F5", fg="#1D516B",
                  activebackground="#D4E5EF", relief="flat", cursor="hand2", font=("Segoe UI", 9, "bold"),
                  padx=15, pady=6).grid(row=row, column=2, pady=7)

    def create_combo_field(self, parent, column, label, variable, values):
        field = tk.Frame(parent, bg="white")
        field.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 12, 0))
        tk.Label(field, text=label, bg="white", fg="#263847", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Combobox(field, textvariable=variable, values=values, state="readonly", style="Suite.TCombobox").pack(
            fill="x", pady=(6, 0), ipady=2)

    def browse_excel(self):

        filename = filedialog.askopenfilename(

            title="Select Equipment Excel",

            filetypes=[

                ("Excel Workbook", "*.xlsx"),

                ("All Files", "*.*")

            ]

        )

        if filename:

            self.excel_path.set(filename)

    # ==========================================================
    # Browse Word Template
    # ==========================================================

    def browse_template(self):

        filename = filedialog.askopenfilename(

            title="Select Word Template",

            filetypes=[

                ("Word Document", "*.docx"),

                ("All Files", "*.*")

            ]

        )

        if filename:

            self.template_path.set(filename)

    # ==========================================================
    # Browse Output File
    # ==========================================================

    def browse_output(self):

        filename = filedialog.asksaveasfilename(

            title="Save Output",

            defaultextension=".docx",

            filetypes=[

                ("Word Document", "*.docx")

            ]

        )

        if filename:

            self.output_path.set(filename)

    # ==========================================================
    # Validate Inputs
    # ==========================================================

    def validate(self):

        if not self.excel_path.get():

            messagebox.showerror(

                "Missing File",

                "Please select Equipment Excel."

            )

            return False

        if not self.template_path.get():

            messagebox.showerror(

                "Missing File",

                "Please select Word Template."

            )

            return False

        if not self.output_path.get():

            messagebox.showerror(

                "Missing File",

                "Please select Output File."

            )

            return False

        return True

    # ==========================================================
    # Generate Button
    # ==========================================================

    def generate(self):

        if not self.validate():

            return

        self.save_settings()

        self.generate_button.config(

            state="disabled"

        )

        self.progress.start(10)

        self.status.set(

            "Generating Log Cards..."

        )

        threading.Thread(

            target=self.generate_thread,

            daemon=True

        ).start()

    # ==========================================================
    # Background Thread
    # ==========================================================

    def generate_thread(self):

        try:

            controller = Controller(

                template_path=self.template_path.get(),

                excel_path=self.excel_path.get(),

                output_file=self.output_path.get(),

                year=self.year.get(),

                quarter=self.quarter.get(),

                weekday=self.weekday.get(),

                sheet_name=self.sheet_name.get()

            )

            success = controller.run()

            self.root.after(

                0,

                lambda: self.finish_generation(success)

            )

        except Exception as e:

            self.root.after(

                0,

                lambda: self.show_error(str(e))

            )

    # ==========================================================
    # Finish Generation
    # ==========================================================

    def finish_generation(self, success):

        self.progress.stop()

        self.generate_button.config(

            state="normal"

        )

        if success:

            self.status.set(

                "Generation Complete"

            )

            messagebox.showinfo(

                "Success",

                "Log Cards generated successfully."

            )

        else:

            self.status.set(

                "Nothing Generated"

            )

            messagebox.showwarning(

                "Finished",

                "No equipment requires maintenance."

            )

    # ==========================================================
    # Show Error
    # ==========================================================

    def show_error(self, message):

        self.progress.stop()

        self.generate_button.config(

            state="normal"

        )

        self.status.set(

            "Error"

        )

        messagebox.showerror(

            "Error",

            message

        )
            # ==========================================================
    # Back to Dashboard
    # ==========================================================

    def go_back(self):

        self.save_settings()

        self.root.destroy()

        if self.dashboard is not None:

            self.dashboard.deiconify()

    # ==========================================================
    # Window Close
    # ==========================================================

    def on_close(self):

        self.save_settings()

        self.root.destroy()

        if self.dashboard is not None:

            try:

                self.dashboard.destroy()

            except Exception:

                pass

    # ==========================================================
    # Update Status
    # ==========================================================

    def set_status(self, message):

        self.status.set(message)

        self.root.update_idletasks()

    # ==========================================================
    # Enable / Disable Generate Button
    # ==========================================================

    def enable_generate(self):

        self.generate_button.config(

            state="normal"

        )

    def disable_generate(self):

        self.generate_button.config(

            state="disabled"

        )

    # ==========================================================
    # Start Progress Bar
    # ==========================================================

    def start_progress(self):

        self.progress.start(10)

    # ==========================================================
    # Stop Progress Bar
    # ==========================================================

    def stop_progress(self):

        self.progress.stop()
