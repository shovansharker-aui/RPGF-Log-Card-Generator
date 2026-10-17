"""
schedule_gui.py

Weekly Maintenance Schedule generator
RPGF Engineering Suite v2.0
"""

import json
import sys
import threading
from datetime import date
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk

import schedule_store
import workspace
from config import OUTPUT_DIR
from due_controller import DueController
from mail_helper import (create_outlook_draft, default_mail_settings,
                         get_mail_settings, save_mail_settings)
from schedule_controller import (ScheduleController, find_latest_schedule,
                                 fridays_in_month, read_schedule_file,
                                 unique_dir, unique_path,
                                 upcoming_friday)
from ui_common import ProgressPanel, ProgressTracker, px


CHECKED = "\u2611"
UNCHECKED = "\u2610"

MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]

ORDINALS = {1: "1st", 2: "2nd", 3: "3rd", 4: "4th", 5: "5th"}


class ScheduleWindow:

    def __init__(self, root, on_back, get_files):

        self.root = root
        self.get_files = get_files
        self.on_back = on_back
        self.page = tk.Frame(root, bg="#F4F7FB")
        self.page.pack(fill="both", expand=True)

        self.settings_file = "settings.json"

        # The schedule currently shown; the Friday to generate is picked
        # with the Year / Month / Friday boxes.
        self.friday = upcoming_friday()

        self.logo = self.load_logo()

        self.rows = []
        self.checked = set()
        self.pdf_file = None
        self.excel_file = None
        self.header_text = tk.StringVar(
            value="No schedule generated yet")

        target = upcoming_friday()
        self.year_var = tk.StringVar(value=str(target.year))
        self.month_var = tk.StringVar(value=MONTH_NAMES[target.month - 1])
        self.friday_var = tk.StringVar()
        self.friday_choices = {}

        self.output_dir = tk.StringVar(value=str(OUTPUT_DIR))
        self.status = tk.StringVar(value="Ready")
        self.count_text = tk.StringVar(value="")
        self.file_text = tk.StringVar(value="")

        self.build_ui()
        self.load_settings()
        self.refresh_fridays(select=target)
        self.load_latest()

    # ==================================================
    # Resources / settings
    # ==================================================

    @staticmethod
    def resource_path(relative_path):

        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = Path(__file__).parent

        return str(Path(base_path) / relative_path)

    def load_logo(self):

        try:
            image = Image.open(self.resource_path("assets/renata.png"))
            return ImageTk.PhotoImage(
                image.resize((px(70), px(70)), Image.LANCZOS))
        except Exception:
            return None

    def _read_settings(self):

        try:
            with open(self.settings_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def save_settings(self):

        data = self._read_settings()

        data["sched_output_dir"] = self.output_dir.get()

        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

    def load_settings(self):

        data = self._read_settings()

        if data.get("sched_output_dir"):
            self.output_dir.set(data["sched_output_dir"])

    # ==================================================
    # UI
    # ==================================================

    def build_ui(self):

        style = ttk.Style(self.root)
        style.configure("Suite.TEntry", padding=8,
                        fieldbackground="white", bordercolor="#B8C6D1")
        style.configure("Sched.Treeview", rowheight=px(28),
                        font=("Segoe UI", 10), background="white",
                        fieldbackground="white")
        style.configure("Sched.Treeview.Heading",
                        font=("Segoe UI", 10, "bold"),
                        background="#E7F0F5", foreground="#123B5D")

        header = tk.Frame(self.page, bg="#123B5D", height=px(110))
        header.pack(fill="x")
        header.pack_propagate(False)

        if self.logo:
            tk.Label(header, image=self.logo, bg="#123B5D").place(
                x=px(40), y=px(20))

        tk.Label(header, text="Maintenance Schedule", bg="#123B5D",
                 fg="white", font=("Segoe UI", 22, "bold")).place(
                     x=px(135), y=px(22))

        tk.Label(header,
                 textvariable=self.header_text,
                 bg="#123B5D", fg="#B9D8E7",
                 font=("Segoe UI", 11)).place(x=px(137), y=px(66))

        tk.Button(header, text="\u2699  Settings",
                  command=self.open_mail_settings, bg="#1F5A85", fg="white",
                  activebackground="#2A6FA3", activeforeground="white",
                  relief="flat", cursor="hand2",
                  font=("Segoe UI", 10, "bold"), padx=14, pady=6).place(
                      relx=1.0, x=-px(24), y=px(20), anchor="ne")

        content = tk.Frame(self.page, bg="#F4F7FB")
        content.pack(fill="both", expand=True, padx=px(42), pady=px(18))

        card = tk.Frame(content, bg="white",
                        highlightbackground="#DCE4EC", highlightthickness=1)
        card.pack(fill="x")

        form = tk.Frame(card, bg="white")
        form.pack(fill="x", padx=24, pady=16)
        form.columnconfigure(1, weight=1)

        files = self.get_files()

        for row, key in enumerate(("workbook", "due_template")):
            found = files.get(key)
            tk.Label(form, bg="white", anchor="w", font=("Segoe UI", 9),
                     fg="#2E7D32" if found else "#B3261E",
                     text=(f"\u2713  {workspace.LABELS[key]}: {found.name}"
                           if found else
                           f"\u2717  {workspace.LABELS[key]}: not found in "
                           "the workspace folder")).grid(
                               row=row, column=0, columnspan=3, sticky="w")

        self.path_field(form, 2, "Output folder", self.output_dir,
                        self.browse_output)

        picker = tk.Frame(card, bg="white")
        picker.pack(fill="x", padx=24, pady=(0, 16))

        tk.Label(picker, text="Generate schedule for", bg="white",
                 fg="#263847", font=("Segoe UI", 10, "bold")).pack(
                     side="left", padx=(0, 14))

        year_box = ttk.Combobox(
            picker, textvariable=self.year_var, width=6, state="readonly",
            values=[str(y) for y in range(date.today().year - 1,
                                          date.today().year + 3)])
        year_box.pack(side="left", padx=(0, 8))

        month_box = ttk.Combobox(
            picker, textvariable=self.month_var, width=12, state="readonly",
            values=MONTH_NAMES)
        month_box.pack(side="left", padx=(0, 8))

        self.friday_box = ttk.Combobox(
            picker, textvariable=self.friday_var, width=28, state="readonly")
        self.friday_box.pack(side="left")

        year_box.bind("<<ComboboxSelected>>", self.on_period_changed)
        month_box.bind("<<ComboboxSelected>>", self.on_period_changed)

        actions = tk.Frame(content, bg="#F4F7FB")
        actions.pack(fill="x", pady=(14, 8))

        self.generate_button = tk.Button(
            actions, text="Generate Schedule", command=self.generate,
            bg="#167D9A", fg="white", activebackground="#10667E",
            activeforeground="white", relief="flat", cursor="hand2",
            font=("Segoe UI", 11, "bold"), padx=25, pady=10)
        self.generate_button.pack(side="left")

        self.tag_button = tk.Button(
            actions, text="Generate Due Tags", command=self.generate_tags,
            bg="#B4690E", fg="white", activebackground="#8F550B",
            activeforeground="white", relief="flat", cursor="hand2",
            font=("Segoe UI", 11, "bold"), padx=25, pady=10,
            state="disabled")
        self.tag_button.pack(side="left", padx=(10, 0))

        self.mail_button = tk.Button(
            actions, text="Create Email Draft", command=self.create_email,
            bg="#2D6A8E", fg="white", activebackground="#245672",
            activeforeground="white", relief="flat", cursor="hand2",
            font=("Segoe UI", 11, "bold"), padx=25, pady=10,
            state="disabled")
        self.mail_button.pack(side="left", padx=(10, 0))

        tk.Button(actions, text="Back to Dashboard", command=self.go_back,
                  bg="#F4F7FB", fg="#34536B", activebackground="#E3EAF0",
                  relief="flat", cursor="hand2", font=("Segoe UI", 10),
                  padx=12, pady=10).pack(side="right")

        toolbar = tk.Frame(content, bg="#F4F7FB")
        toolbar.pack(fill="x", pady=(0, 6))

        tk.Label(toolbar, text="Select to generate Due Tag",
                 bg="#F4F7FB", fg="#34536B",
                 font=("Segoe UI", 10)).pack(side="left")

        tk.Label(toolbar, textvariable=self.file_text, bg="#F4F7FB",
                 fg="#6B7785", font=("Segoe UI", 9)).pack(
                     side="left", padx=(18, 0))

        tk.Label(toolbar, textvariable=self.count_text, bg="#F4F7FB",
                 fg="#123B5D", font=("Segoe UI", 10, "bold")).pack(side="right")

        table = tk.Frame(content, bg="white",
                         highlightbackground="#DCE4EC", highlightthickness=1)
        table.pack(fill="both", expand=True)

        columns = ("check", "sl", "name", "id", "frequency", "date")

        self.tree = ttk.Treeview(table, columns=columns, show="headings",
                                 style="Sched.Treeview", selectmode="browse")

        # Alternate row colours so a row is easy to follow across the list.
        self.tree.tag_configure("odd", background="#FFFFFF")
        self.tree.tag_configure("even", background="#E8EFF6")

        headings = (("check", "", 44, "center"),
                    ("sl", "Sl No", 60, "center"),
                    ("name", "Equipment Name", 300, "w"),
                    ("id", "Equipment ID", 170, "w"),
                    ("frequency", "Frequency", 150, "center"),
                    ("date", "Schedule Date", 120, "center"))

        for key, title, width, anchor in headings:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=px(width), anchor=anchor,
                             stretch=(key == "name"))

        scrollbar = ttk.Scrollbar(table, orient="vertical",
                                  command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<space>", self.on_space)

        self.progress = ProgressPanel(content, self.root, self.status)
        self.progress.pack(fill="x", pady=(14, 0))

    @staticmethod
    def path_field(parent, row, label, variable, browse):

        tk.Label(parent, text=label, bg="white", fg="#263847",
                 font=("Segoe UI", 10, "bold")).grid(
                     row=row, column=0, sticky="w", pady=5)

        ttk.Entry(parent, textvariable=variable, style="Suite.TEntry").grid(
            row=row, column=1, sticky="ew", padx=(18, 10), pady=5, ipady=2)

        tk.Button(parent, text="Browse", command=browse, bg="#E7F0F5",
                  fg="#1D516B", activebackground="#D4E5EF", relief="flat",
                  cursor="hand2", font=("Segoe UI", 9, "bold"),
                  padx=15, pady=5).grid(row=row, column=2, pady=5)

    def browse_output(self):

        path = filedialog.askdirectory(title="Select Output Folder")

        if path:
            self.output_dir.set(path)
            self.save_settings()
            self.load_latest()

    # ==================================================
    # Month / Friday pickers
    # ==================================================

    def on_period_changed(self, _event=None):

        self.refresh_fridays()

    def refresh_fridays(self, select=None):
        """Fill the Friday box with every Friday of the chosen month."""

        year = int(self.year_var.get())
        month = MONTH_NAMES.index(self.month_var.get()) + 1

        self.friday_choices = {}

        labels = []

        for number, friday in enumerate(fridays_in_month(year, month), 1):
            label = (f"{ORDINALS[number]} Friday  \u2013  "
                     f"{friday.strftime('%d.%m.%Y')}")
            self.friday_choices[label] = friday
            labels.append(label)

        self.friday_box.config(values=labels)

        wanted = None

        if select:
            wanted = next((l for l, d in self.friday_choices.items()
                           if d == select), None)

        if wanted is None:
            current = self.friday_var.get()
            wanted = current if current in self.friday_choices else (
                labels[0] if labels else "")

        self.friday_var.set(wanted)

    # ==================================================
    # Generate
    # ==================================================

    def generate(self):

        files = self.get_files()

        if not files.get("workbook"):
            messagebox.showerror(
                "Maintenance Schedule",
                "The equipment workbook was not found in the workspace "
                "folder.\n\nChoose the folder on the dashboard.")
            return

        self.save_settings()

        target = self.friday_choices.get(self.friday_var.get())

        if target is None:
            messagebox.showerror("Maintenance Schedule",
                                 "Please choose the Friday to generate for.")
            return

        # Every press creates its own dated folder (e.g. 17OCT2026) holding
        # the Excel, the PDF and the Word file; earlier ones are kept.
        today = date.today()

        folder = unique_dir(
            Path(self.output_dir.get()) / today.strftime("%d%b%Y").upper())

        output = str(
            folder / f"Maintenance Schedule_{today.strftime('%d_%m_%y')}.xlsx")

        self.generate_button.config(state="disabled")

        self.tracker = ProgressTracker()
        self.progress.start(self.tracker)
        self.status.set("Building schedule...")

        threading.Thread(
            target=self.worker,
            args=(str(files["workbook"]), output, target),
            daemon=True).start()

    def worker(self, excel, output, target):

        try:
            rows = ScheduleController(
                excel, output, friday=target,
                progress=self.tracker.update).generate()
        except Exception as error:
            message = str(error)
            self.root.after(0, lambda: self.failed(message))
            return

        self.root.after(0, lambda: self.finished(rows, output))

    def finished(self, rows, output):

        self.progress.stop()
        self.generate_button.config(state="normal")

        # The file just written is now the latest one: show it from there.
        self.load_latest()

        self.status.set("Schedule created successfully.")

        messagebox.showinfo(
            "Maintenance Schedule",
            f"{len(rows)} machine(s) scheduled for "
            f"{self.friday.strftime('%d.%m.%Y')}.\n\n"
            f"Folder:\n{Path(output).parent}\n\n"
            "Files created:\n"
            f"  \u2022 {Path(output).name}\n"
            f"  \u2022 {Path(output).with_suffix('.pdf').name}\n"
            f"  \u2022 {Path(output).with_suffix('.docx').name}")

    # ==================================================
    # Latest schedule file
    # ==================================================

    def load_latest(self, keep_ticks=False):
        """Show the newest 'Maintenance Schedule_*.xlsx' of the output
        folder. Ticks are kept (by Equipment ID) when asked, or restored from
        the last session when the same file is still the latest."""

        latest = find_latest_schedule(self.output_dir.get())

        if latest is None:
            self.rows = []
            self.checked = set()
            self.excel_file = None
            self.pdf_file = None
            self.tree.delete(*self.tree.get_children())
            self.header_text.set("No schedule generated yet")
            self.file_text.set("")
            self.tag_button.config(state="disabled")
            self.mail_button.config(state="disabled")
            self.update_count()
            return False

        previous_ids = {
            self.rows[index]["id"] for index in self.checked
            if index < len(self.rows)
        } if keep_ticks else None

        try:
            loaded = read_schedule_file(latest)
        except Exception as error:
            messagebox.showerror(
                "Maintenance Schedule",
                f"Could not read the latest schedule:\n{latest.name}\n\n"
                f"{error}")
            return False

        ticked = (previous_ids if previous_ids is not None
                  else schedule_store.load_ticks(latest))

        self.excel_file = latest
        self.pdf_file = latest.with_suffix(".pdf")
        self.friday = loaded["friday"]
        self.rows = loaded["rows"]
        self.checked = {
            index for index, item in enumerate(self.rows)
            if item["id"] in ticked
        }

        self.tree.delete(*self.tree.get_children())

        for index, item in enumerate(self.rows):
            self.tree.insert("", "end", iid=str(index), values=(
                CHECKED if index in self.checked else UNCHECKED,
                index + 1, item["name"], item["id"], item["frequency"],
                item["date"].strftime("%d.%m.%Y")),
                tags=("even" if index % 2 else "odd",))

        self.header_text.set(
            f"Schedule for Friday, {self.friday.strftime('%d.%m.%Y')}"
            f"   \u2022   generated "
            f"{loaded['generated'].strftime('%d.%m.%Y')}")

        self.file_text.set(f"File: {latest.name}")

        self.mail_button.config(
            state="normal" if self.pdf_file.exists() else "disabled")

        self.update_count()

        return True

    def save_state(self):
        """Remember the ticks for the next time the module opens."""

        if not self.excel_file:
            return

        try:
            schedule_store.save_ticks(
                self.excel_file,
                {self.rows[i]["id"] for i in self.checked
                 if i < len(self.rows)})
        except Exception:
            pass

    # ==================================================
    # Selection
    # ==================================================

    def update_count(self):

        total = len(self.rows)

        selected = [item for index, item in enumerate(self.rows)
                    if index in self.checked]

        tags = sum(len(item["items"]) for item in selected)

        self.count_text.set(
            f"{total} scheduled  \u2022  {len(selected)} selected  \u2022  "
            f"{tags} due tags")

        if self.rows:
            self.tag_button.config(
                state="normal" if selected else "disabled")

    def toggle(self, iid):

        index = int(iid)

        if index in self.checked:
            self.checked.discard(index)
        else:
            self.checked.add(index)

        self.tree.set(iid, "check",
                      CHECKED if index in self.checked else UNCHECKED)

        self.update_count()
        self.save_state()

    def on_click(self, event):

        if self.tree.identify_region(event.x, event.y) != "cell":
            return

        iid = self.tree.identify_row(event.y)

        if iid:
            self.toggle(iid)
            return "break"

    def on_space(self, _event):

        for iid in self.tree.selection():
            self.toggle(iid)

        return "break"

    def refresh_checks(self):

        for iid in self.tree.get_children():
            self.tree.set(iid, "check",
                          CHECKED if int(iid) in self.checked else UNCHECKED)

        self.update_count()

    # ==================================================
    # Mail settings
    # ==================================================

    def open_mail_settings(self):

        window = tk.Toplevel(self.root)
        window.title("Mail Settings")
        window.configure(bg="white")
        window.transient(self.root)
        window.resizable(False, False)

        settings = get_mail_settings()

        variables = {
            "to": tk.StringVar(value=settings["to"]),
            "cc": tk.StringVar(value=settings["cc"]),
            "subject": tk.StringVar(value=settings["subject"]),
        }

        tk.Label(window, text="Email template", bg="white", fg="#123B5D",
                 font=("Segoe UI", 14, "bold")).grid(
                     row=0, column=0, columnspan=2, sticky="w",
                     padx=24, pady=(20, 2))

        tk.Label(window, text="Used when you click Create Email Draft. Your "
                 "Outlook signature is added below the body automatically.",
                 bg="white", fg="#6B7785", font=("Segoe UI", 9),
                 wraplength=px(520), justify="left").grid(
                     row=1, column=0, columnspan=2, sticky="w",
                     padx=24, pady=(0, 14))

        for row, (key, label) in enumerate(
                (("to", "To"), ("cc", "CC"), ("subject", "Subject")),
                start=2):

            tk.Label(window, text=label, bg="white", fg="#263847",
                     font=("Segoe UI", 10, "bold")).grid(
                         row=row, column=0, sticky="nw", padx=(24, 12),
                         pady=6)

            ttk.Entry(window, textvariable=variables[key], width=62,
                      style="Suite.TEntry").grid(
                          row=row, column=1, sticky="ew", padx=(0, 24),
                          pady=6, ipady=2)

        tk.Label(window, text="Body", bg="white", fg="#263847",
                 font=("Segoe UI", 10, "bold")).grid(
                     row=5, column=0, sticky="nw", padx=(24, 12), pady=6)

        body = tk.Text(window, width=62, height=6, wrap="word",
                       font=("Segoe UI", 10), relief="solid", bd=1,
                       padx=8, pady=6)
        body.grid(row=5, column=1, sticky="ew", padx=(0, 24), pady=6)
        body.insert("1.0", settings["body"])

        tk.Label(window,
                 text="Separate multiple recipients with ;   Each line of the "
                 "body is one paragraph.   {date_long} = 14th October, 2026"
                 "   {date} = 14-Oct-2026",
                 bg="white", fg="#6B7785", font=("Segoe UI", 9),
                 wraplength=px(520), justify="left").grid(
                     row=6, column=1, sticky="w", padx=(0, 24), pady=(0, 10))

        buttons = tk.Frame(window, bg="white")
        buttons.grid(row=7, column=0, columnspan=2, sticky="ew",
                     padx=24, pady=(6, 20))

        def save():
            save_mail_settings({
                "to": variables["to"].get().strip(),
                "cc": variables["cc"].get().strip(),
                "subject": variables["subject"].get().strip(),
                "body": body.get("1.0", "end").strip(),
            })
            self.status.set("Mail settings saved.")
            window.destroy()

        def reset():
            defaults = default_mail_settings()
            for key in variables:
                variables[key].set(defaults[key])
            body.delete("1.0", "end")
            body.insert("1.0", defaults["body"])

        tk.Button(buttons, text="Save", command=save, bg="#167D9A",
                  fg="white", activebackground="#10667E",
                  activeforeground="white", relief="flat", cursor="hand2",
                  font=("Segoe UI", 10, "bold"), padx=22, pady=7).pack(
                      side="left")

        tk.Button(buttons, text="Cancel", command=window.destroy,
                  bg="#E7F0F5", fg="#1D516B", activebackground="#D4E5EF",
                  relief="flat", cursor="hand2",
                  font=("Segoe UI", 10), padx=18, pady=7).pack(
                      side="left", padx=(10, 0))

        tk.Button(buttons, text="Reset to default", command=reset,
                  bg="white", fg="#B3261E", activebackground="#FDECEA",
                  relief="flat", cursor="hand2",
                  font=("Segoe UI", 9), padx=10, pady=7).pack(side="right")

        window.update_idletasks()

        x = self.root.winfo_rootx() + (self.root.winfo_width()
                                       - window.winfo_width()) // 2

        # Keep the whole dialog (and its Save button) above the taskbar.
        limit = self.root.winfo_screenheight() - px(90) - window.winfo_height()

        y = min(self.root.winfo_rooty() + px(40), limit)

        window.geometry(f"+{max(x, 0)}+{max(y, 0)}")

        window.grab_set()

    # ==================================================
    # Email draft
    # ==================================================

    def create_email(self):

        if not self.pdf_file or not Path(self.pdf_file).exists():
            messagebox.showerror("Maintenance Schedule",
                                 "Generate the schedule first so the PDF "
                                 "exists.")
            return

        try:
            create_outlook_draft(self.pdf_file, date.today())
        except Exception as error:
            messagebox.showerror("Maintenance Schedule", str(error))
            return

        self.status.set("Outlook draft created.")

    # ==================================================
    # Due tags
    # ==================================================

    def generate_tags(self):

        previous = self.excel_file

        # Always work from the latest schedule file in the output folder.
        self.load_latest(keep_ticks=True)

        if self.excel_file != previous:
            self.status.set(f"Using the latest schedule: "
                            f"{self.excel_file.name if self.excel_file else '-'}")

        if not self.rows:
            messagebox.showinfo("Maintenance Schedule",
                                "No schedule found in the output folder yet.")
            return

        # One tag per frequency for every ticked machine.
        selected = [
            {
                "name": item["name"],
                "id": item["id"],
                "frequency": entry["frequency"],
                "date": entry["date"].strftime("%d.%m.%Y"),
                "last": entry["last"].strftime("%d.%m.%Y"),
            }
            for index, item in enumerate(self.rows)
            if index in self.checked
            for entry in item["items"]
        ]

        if not selected:
            messagebox.showinfo("Maintenance Schedule",
                                "Select at least one machine to generate "
                                "due tags.")
            return

        files = self.get_files()

        if not files.get("due_template"):
            messagebox.showerror(
                "Maintenance Schedule",
                "The due tag template was not found in the workspace "
                "folder.\n\nChoose the folder on the dashboard.")
            return

        self.save_settings()

        # The tags go into the same dated folder as the schedule they use.
        output = str(unique_path(
            Path(self.excel_file).parent /
            f"Maintenance_Due_Tags_{self.friday.strftime('%d.%m.%Y')}.docx",
            also=()))

        self.generate_button.config(state="disabled")
        self.tag_button.config(state="disabled")

        self.tracker = ProgressTracker()
        self.progress.start(self.tracker)
        self.status.set("Generating due tags...")

        threading.Thread(
            target=self.tag_worker,
            args=(selected, str(files["due_template"]), output),
            daemon=True).start()

    def tag_worker(self, selected, template, output):

        try:
            DueController().generate(
                template, selected, output, progress=self.tracker.update)
        except Exception as error:
            message = str(error)
            self.root.after(0, lambda: self.tags_failed(message))
            return

        count = len(selected)
        self.root.after(0, lambda: self.tags_finished(count, output))

    def tags_finished(self, count, output):

        self.progress.stop()
        self.generate_button.config(state="normal")
        self.tag_button.config(state="normal")
        self.status.set(f"Created {count} due tag(s).")

        messagebox.showinfo("Maintenance Schedule",
                            f"{count} due tag(s) saved to:\n{output}")

    def tags_failed(self, message):

        self.progress.stop(complete=False)
        self.generate_button.config(state="normal")
        self.tag_button.config(state="normal")
        self.status.set("Due tag generation failed.")

        messagebox.showerror("Maintenance Schedule", message)

    def failed(self, message):

        self.progress.stop(complete=False)
        self.generate_button.config(state="normal")
        self.status.set("Schedule generation failed.")

        messagebox.showerror("Maintenance Schedule", message)

    def go_back(self):

        self.save_settings()

        self.page.destroy()

        self.on_back()
