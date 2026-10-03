"""
due_gui.py

Maintenance Due label tag generator
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
from ui_common import ProgressPanel, ProgressTracker, px

CHECKED = "☑"
UNCHECKED = "☐"


class DueWindow:

    def __init__(self, root, on_back):

        self.root = root
        self.on_back = on_back
        self.page = tk.Frame(root, bg="#F4F7FB")
        self.page.pack(fill="both", expand=True)

        self.settings_file = "settings.json"
        self.machines = []
        self.checked = set()

        self.logo = self.load_logo()

        self.excel_path = tk.StringVar()
        self.template_path = tk.StringVar(value=str(DUE_TEMPLATE))
        self.output_path = tk.StringVar(
            value=str(OUTPUT_DIR / "Maintenance_Due_Labels.docx"))
        self.status = tk.StringVar(value="Load an equipment workbook to begin.")
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

        data["due_excel"] = self.excel_path.get()
        data["due_template"] = self.template_path.get()
        data["due_output"] = self.output_path.get()

        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

    def load_settings(self):

        data = self._read_settings()

        self.excel_path.set(data.get("due_excel", ""))

        if data.get("due_template"):
            self.template_path.set(data["due_template"])

        if data.get("due_output"):
            self.output_path.set(data["due_output"])

    # ==================================================
    # UI
    # ==================================================

    def build_ui(self):

        style = ttk.Style(self.root)
        style.configure("Suite.TEntry", padding=8,
                        fieldbackground="white", bordercolor="#B8C6D1")
        style.configure("Due.Treeview", rowheight=px(30), font=("Segoe UI", 10),
                        background="white", fieldbackground="white")
        style.configure("Due.Treeview.Heading", font=("Segoe UI", 10, "bold"),
                        background="#E7F0F5", foreground="#123B5D")
        style.map("Due.Treeview", background=[("selected", "#D7E8F1")],
                  foreground=[("selected", "#123B5D")])

        header = tk.Frame(self.page, bg="#123B5D", height=px(110))
        header.pack(fill="x")
        header.pack_propagate(False)

        if self.logo:
            tk.Label(header, image=self.logo, bg="#123B5D").place(
                x=px(40), y=px(20))

        tk.Label(header, text="Maintenance Due Labels", bg="#123B5D",
                 fg="white", font=("Segoe UI", 22, "bold")).place(
                     x=px(135), y=px(30))

        content = tk.Frame(self.page, bg="#F4F7FB")
        content.pack(fill="both", expand=True, padx=px(42), pady=px(18))

        card = tk.Frame(content, bg="white",
                        highlightbackground="#DCE4EC", highlightthickness=1)
        card.pack(fill="x")

        form = tk.Frame(card, bg="white")
        form.pack(fill="x", padx=24, pady=16)
        form.columnconfigure(1, weight=1)

        self.path_field(form, 0, "Equipment workbook", self.excel_path,
                        self.browse_excel, "Load List", self.load_list)
        self.path_field(form, 1, "Label template", self.template_path,
                        self.browse_template)
        self.path_field(form, 2, "Output document", self.output_path,
                        self.browse_output)

        # ---- machine list ----
        toolbar = tk.Frame(content, bg="#F4F7FB")
        toolbar.pack(fill="x", pady=(14, 6))

        tk.Label(toolbar, text="Tick the machines to leave out. Unticked "
                 "machines get a label.", bg="#F4F7FB", fg="#34536B",
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
                                 style="Due.Treeview", selectmode="browse")

        headings = (("check", "", 44, "center"), ("sl", "Sl", 50, "center"),
                    ("name", "Equipment Name", 280, "w"),
                    ("id", "Equipment ID", 130, "w"),
                    ("frequency", "Frequency", 90, "center"),
                    ("date", "Maintenance Date", 130, "center"),
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

        # ---- actions ----
        actions = tk.Frame(content, bg="#F4F7FB")
        actions.pack(fill="x", pady=(14, 0))

        self.confirm_button = tk.Button(
            actions, text="Confirm & Generate", command=self.confirm,
            bg="#167D9A", fg="white", activebackground="#10667E",
            activeforeground="white", relief="flat", cursor="hand2",
            font=("Segoe UI", 11, "bold"), padx=25, pady=10)
        self.confirm_button.pack(side="left")

        tk.Button(actions, text="Back to Dashboard", command=self.go_back,
                  bg="#F4F7FB", fg="#34536B", activebackground="#E3EAF0",
                  relief="flat", cursor="hand2", font=("Segoe UI", 10),
                  padx=12, pady=10).pack(side="right")

        self.progress = ProgressPanel(content, self.root, self.status)
        self.progress.pack(fill="x", pady=(14, 0))

    @staticmethod
    def path_field(parent, row, label, variable, browse, extra_text=None,
                   extra_command=None):

        tk.Label(parent, text=label, bg="white", fg="#263847",
                 font=("Segoe UI", 10, "bold")).grid(
                     row=row, column=0, sticky="w", pady=5)

        ttk.Entry(parent, textvariable=variable, style="Suite.TEntry").grid(
            row=row, column=1, sticky="ew", padx=(18, 10), pady=5, ipady=2)

        tk.Button(parent, text="Browse", command=browse, bg="#E7F0F5",
                  fg="#1D516B", activebackground="#D4E5EF", relief="flat",
                  cursor="hand2", font=("Segoe UI", 9, "bold"),
                  padx=15, pady=5).grid(row=row, column=2, pady=5)

        if extra_text:
            tk.Button(parent, text=extra_text, command=extra_command,
                      bg="#167D9A", fg="white", activebackground="#10667E",
                      activeforeground="white", relief="flat", cursor="hand2",
                      font=("Segoe UI", 9, "bold"), padx=15, pady=5).grid(
                          row=row, column=3, padx=(8, 0), pady=5)

    # ==================================================
    # Browse
    # ==================================================

    def browse_excel(self):

        path = filedialog.askopenfilename(
            title="Select Equipment Workbook",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")])

        if path:
            self.excel_path.set(path)
            self.load_list()

    def browse_template(self):

        path = filedialog.askopenfilename(
            title="Select Label Template",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")])

        if path:
            self.template_path.set(path)

    def browse_output(self):

        path = filedialog.asksaveasfilename(
            title="Save Output", defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")])

        if path:
            self.output_path.set(path)

    # ==================================================
    # List handling
    # ==================================================

    def load_list(self):

        if not self.excel_path.get():
            messagebox.showerror("Maintenance Due",
                                 "Please select the equipment workbook.")
            return

        try:
            self.machines = DueController.load_machines(self.excel_path.get())
        except Exception as error:
            messagebox.showerror("Maintenance Due", str(error))
            return

        self.checked = set()
        self.refresh_tree()
        self.status.set(f"Loaded {len(self.machines)} machines.")

    def refresh_tree(self):

        self.tree.delete(*self.tree.get_children())

        for index, machine in enumerate(self.machines):
            self.tree.insert("", "end", iid=str(index), values=(
                UNCHECKED, index + 1, machine["name"], machine["id"],
                machine["frequency"], machine["date"], machine["last"]))

        self.update_count()

    def update_count(self):

        total = len(self.machines)
        remaining = total - len(self.checked)

        self.count_text.set(f"{remaining} of {total} will get a label")

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

        self.checked = set(range(len(self.machines)))
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
    # Generate
    # ==================================================

    def confirm(self):

        if not self.machines:
            messagebox.showerror("Maintenance Due",
                                 "Load the equipment list first.")
            return

        selected = [machine for index, machine in enumerate(self.machines)
                    if index not in self.checked]

        if not selected:
            messagebox.showinfo("Maintenance Due",
                                "Every machine is ticked, so there is "
                                "nothing to print.")
            return

        if not self.output_path.get():
            messagebox.showerror("Maintenance Due",
                                 "Please select an output document.")
            return

        self.save_settings()
        self.confirm_button.config(state="disabled")

        self.tracker = ProgressTracker()
        self.progress.start(self.tracker)
        self.status.set("Generating labels...")

        threading.Thread(
            target=self.generate,
            args=(selected, self.template_path.get(), self.output_path.get()),
            daemon=True).start()

    def generate(self, selected, template, output):

        try:
            DueController().generate(
                template, selected, output, progress=self.tracker.update)
        except Exception as error:
            message = str(error)
            self.root.after(0, lambda: self.failed(message))
            return

        count = len(selected)
        self.root.after(0, lambda: self.finished(count))

    def finished(self, count):

        self.progress.stop()
        self.confirm_button.config(state="normal")
        self.status.set(f"Created {count} label(s).")

        messagebox.showinfo("Maintenance Due",
                            f"{count} label(s) saved to:\n"
                            f"{self.output_path.get()}")

    def failed(self, message):

        self.progress.stop(complete=False)
        self.confirm_button.config(state="normal")
        self.status.set("Label generation failed.")

        messagebox.showerror("Maintenance Due", message)

    def go_back(self):

        self.save_settings()

        self.page.destroy()

        self.on_back()
