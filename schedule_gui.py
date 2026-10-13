"""
schedule_gui.py

Weekly Maintenance Schedule generator
RPGF Engineering Suite v2.0
"""

import json
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk

from config import DUE_TEMPLATE, OUTPUT_DIR
from due_controller import DueController
from schedule_controller import ScheduleController, upcoming_friday
from ui_common import ProgressPanel, ProgressTracker, px


CHECKED = "\u2611"
UNCHECKED = "\u2610"


class ScheduleWindow:

    def __init__(self, root, on_back):

        self.root = root
        self.on_back = on_back
        self.page = tk.Frame(root, bg="#F4F7FB")
        self.page.pack(fill="both", expand=True)

        self.settings_file = "settings.json"

        # Worked out when the module opens, from the current date.
        self.friday = upcoming_friday()

        self.logo = self.load_logo()

        self.rows = []
        self.checked = set()

        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar(value=str(DUE_TEMPLATE))
        self.output_dir = tk.StringVar(value=str(OUTPUT_DIR))
        self.status = tk.StringVar(value="Ready")
        self.count_text = tk.StringVar(value="")

        self.build_ui()
        self.load_settings()

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

        data["sched_excel"] = self.excel_path.get()
        data["sched_output_dir"] = self.output_dir.get()
        data["sched_tag_template"] = self.template_path.get()

        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

    def load_settings(self):

        data = self._read_settings()

        self.excel_path.set(
            data.get("sched_excel")
            or data.get("trend_excel")
            or data.get("excel", ""))

        if data.get("sched_output_dir"):
            self.output_dir.set(data["sched_output_dir"])

        if data.get("sched_tag_template"):
            self.template_path.set(data["sched_tag_template"])

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
                 text=f"Upcoming Friday: {self.friday.strftime('%A, %d %B %Y')}",
                 bg="#123B5D", fg="#B9D8E7",
                 font=("Segoe UI", 11)).place(x=px(137), y=px(66))

        content = tk.Frame(self.page, bg="#F4F7FB")
        content.pack(fill="both", expand=True, padx=px(42), pady=px(18))

        card = tk.Frame(content, bg="white",
                        highlightbackground="#DCE4EC", highlightthickness=1)
        card.pack(fill="x")

        form = tk.Frame(card, bg="white")
        form.pack(fill="x", padx=24, pady=16)
        form.columnconfigure(1, weight=1)

        self.path_field(form, 0, "Maintenance matrix", self.excel_path,
                        self.browse_excel)
        self.path_field(form, 1, "Output folder", self.output_dir,
                        self.browse_output)
        self.path_field(form, 2, "Due tag template", self.template_path,
                        self.browse_template)

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

        tk.Button(actions, text="Back to Dashboard", command=self.go_back,
                  bg="#F4F7FB", fg="#34536B", activebackground="#E3EAF0",
                  relief="flat", cursor="hand2", font=("Segoe UI", 10),
                  padx=12, pady=10).pack(side="right")

        toolbar = tk.Frame(content, bg="#F4F7FB")
        toolbar.pack(fill="x", pady=(0, 6))

        tk.Label(toolbar, text="Tick the machines already done. Unticked "
                 "machines get a due tag.", bg="#F4F7FB", fg="#34536B",
                 font=("Segoe UI", 10)).pack(side="left")

        tk.Label(toolbar, textvariable=self.count_text, bg="#F4F7FB",
                 fg="#123B5D", font=("Segoe UI", 10, "bold")).pack(side="right")

        for text, command in (("Clear All", self.clear_all),
                              ("Select All", self.select_all)):
            tk.Button(toolbar, text=text, command=command, bg="#E7F0F5",
                      fg="#1D516B", activebackground="#D4E5EF", relief="flat",
                      cursor="hand2", font=("Segoe UI", 9, "bold"),
                      padx=12, pady=4).pack(side="right", padx=(0, 8))

        table = tk.Frame(content, bg="white",
                         highlightbackground="#DCE4EC", highlightthickness=1)
        table.pack(fill="both", expand=True)

        columns = ("check", "sl", "name", "id", "frequency", "date", "last")

        self.tree = ttk.Treeview(table, columns=columns, show="headings",
                                 style="Sched.Treeview", selectmode="browse")

        headings = (("check", "", 44, "center"),
                    ("sl", "Sl No", 60, "center"),
                    ("name", "Equipment Name", 300, "w"),
                    ("id", "Equipment ID", 170, "w"),
                    ("frequency", "Frequency", 150, "center"),
                    ("date", "Schedule Date", 120, "center"),
                    ("last", "Last Date", 120, "center"))

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

    def browse_excel(self):

        path = filedialog.askopenfilename(
            title="Select Maintenance Matrix",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")])

        if path:
            self.excel_path.set(path)

    def browse_template(self):

        path = filedialog.askopenfilename(
            title="Select Due Tag Template",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")])

        if path:
            self.template_path.set(path)

    def browse_output(self):

        path = filedialog.askdirectory(title="Select Output Folder")

        if path:
            self.output_dir.set(path)

    # ==================================================
    # Generate
    # ==================================================

    def generate(self):

        if not self.excel_path.get():
            messagebox.showerror("Maintenance Schedule",
                                 "Please select the maintenance matrix.")
            return

        self.save_settings()

        output = str(
            Path(self.output_dir.get()) /
            f"Maintenance_Schedule_{self.friday.strftime('%d-%b-%Y')}.xlsx")

        self.generate_button.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.count_text.set("")

        self.tracker = ProgressTracker()
        self.progress.start(self.tracker)
        self.status.set("Building schedule...")

        threading.Thread(
            target=self.worker,
            args=(self.excel_path.get(), output),
            daemon=True).start()

    def worker(self, excel, output):

        try:
            rows = ScheduleController(
                excel, output, friday=self.friday,
                progress=self.tracker.update).generate()
        except Exception as error:
            message = str(error)
            self.root.after(0, lambda: self.failed(message))
            return

        self.root.after(0, lambda: self.finished(rows, output))

    def finished(self, rows, output):

        self.progress.stop()
        self.generate_button.config(state="normal")

        self.rows = rows
        self.checked = set()
        self.tree.delete(*self.tree.get_children())

        for index, item in enumerate(rows):
            self.tree.insert("", "end", iid=str(index), values=(
                UNCHECKED, index + 1, item["name"], item["id"],
                item["frequency"], item["date"].strftime("%d-%b-%Y"),
                item["last"].strftime("%d-%b-%Y")))

        self.tag_button.config(state="normal" if rows else "disabled")
        self.update_count()
        self.status.set("Schedule created successfully.")

        messagebox.showinfo(
            "Maintenance Schedule",
            f"{len(rows)} machine(s) scheduled for "
            f"{self.friday.strftime('%d-%b-%Y')}.\n\nSaved to:\n{output}")

    # ==================================================
    # Selection
    # ==================================================

    def update_count(self):

        total = len(self.rows)

        self.count_text.set(
            f"{total} scheduled  \u2022  {total - len(self.checked)} "
            "will get a due tag")

    def toggle(self, iid):

        index = int(iid)

        if index in self.checked:
            self.checked.discard(index)
        else:
            self.checked.add(index)

        self.tree.set(iid, "check",
                      CHECKED if index in self.checked else UNCHECKED)

        self.update_count()

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

    def select_all(self):

        self.checked = set(range(len(self.rows)))
        self.refresh_checks()

    def clear_all(self):

        self.checked = set()
        self.refresh_checks()

    def refresh_checks(self):

        for iid in self.tree.get_children():
            self.tree.set(iid, "check",
                          CHECKED if int(iid) in self.checked else UNCHECKED)

        self.update_count()

    # ==================================================
    # Due tags
    # ==================================================

    def generate_tags(self):

        selected = [
            {
                "name": item["name"],
                "id": item["id"],
                "frequency": item["frequency"],
                "date": item["date"].strftime("%d-%b-%Y"),
                "last": item["last"].strftime("%d-%b-%Y"),
            }
            for index, item in enumerate(self.rows)
            if index not in self.checked
        ]

        if not selected:
            messagebox.showinfo("Maintenance Schedule",
                                "Every machine is ticked, so there is "
                                "nothing to tag.")
            return

        self.save_settings()

        output = str(
            Path(self.output_dir.get()) /
            f"Maintenance_Due_Tags_{self.friday.strftime('%d-%b-%Y')}.docx")

        self.generate_button.config(state="disabled")
        self.tag_button.config(state="disabled")

        self.tracker = ProgressTracker()
        self.progress.start(self.tracker)
        self.status.set("Generating due tags...")

        threading.Thread(
            target=self.tag_worker,
            args=(selected, self.template_path.get(), output),
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
