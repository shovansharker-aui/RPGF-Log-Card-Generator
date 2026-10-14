"""Main dashboard for RPGF Engineering Suite v2.0."""

import sys
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, ttk

from PIL import Image, ImageTk

import workspace
from ui_common import enable_high_dpi, init_scaling, px, set_app_icon


class Dashboard:

    def __init__(self):
        enable_high_dpi()
        self.root = tk.Tk()
        init_scaling(self.root)
        set_app_icon(self.root, self.resource_path("assets/renata.ico"))
        ttk.Style(self.root).theme_use("clam")
        self.root.title("RPGF Engineering Suite v2.0")
        self.root.configure(bg="#F4F7FB")
        self.center_window(980, 780)
        self.root.minsize(px(820), px(540))
        self.logo = self.load_logo()
        self.folder = tk.StringVar(value=workspace.load_folder())
        self.home = tk.Frame(self.root, bg="#F4F7FB")
        self.build_ui()
        self.show_home()
        self.root.mainloop()

    def center_window(self, width, height):
        width = min(px(width), self.root.winfo_screenwidth() - px(40))
        height = min(px(height), self.root.winfo_screenheight() - px(130))
        x = (self.root.winfo_screenwidth() - width) // 2
        y = max((self.root.winfo_screenheight() - height) // 2 - px(30), 0)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def show_home(self):
        self.root.title("RPGF Engineering Suite v2.0")
        self.home.pack(fill="both", expand=True)

    def open_module(self, module_class, title):
        self.home.pack_forget()
        self.root.title(title)
        module_class(self.root, self.show_home, self.workspace_files)

    # ---- workspace folder ----

    def workspace_files(self):
        """Files found in the chosen folder (rescanned on every call)."""
        return workspace.find_files(self.folder.get())

    def browse_folder(self):
        path = filedialog.askdirectory(title="Select the workspace folder",
                                       initialdir=self.folder.get() or None)
        if path:
            self.folder.set(path)
            workspace.save_folder(path)
            self.refresh_workspace()

    def refresh_workspace(self):
        files = self.workspace_files()
        for key, label in self.file_labels.items():
            found = files.get(key)
            if found:
                label.config(text=f"\u2713  {workspace.LABELS[key]}: {found.name}", fg="#2E7D32")
            else:
                label.config(text=f"\u2717  {workspace.LABELS[key]}: {workspace.FILE_NAMES[key]} not found",
                             fg="#B3261E")

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = Path(__file__).parent
        return str(Path(base_path) / relative_path)

    def load_logo(self):
        try:
            image = Image.open(self.resource_path("assets/renata.png"))
            return ImageTk.PhotoImage(image.resize((px(86), px(86)), Image.LANCZOS))
        except Exception:
            return None

    def build_ui(self):
        header = tk.Frame(self.home, bg="#123B5D", height=px(180))
        header.pack(fill="x")
        header.pack_propagate(False)
        if self.logo:
            tk.Label(header, image=self.logo, bg="#123B5D").place(x=px(52), y=px(42))
        tk.Label(header, text="RENATA PLC", bg="#123B5D", fg="#B9D8E7",
                 font=("Segoe UI", 10, "bold")).place(x=px(158), y=px(48))
        tk.Label(header, text="RPGF Engineering Suite", bg="#123B5D", fg="white",
                 font=("Segoe UI", 25, "bold")).place(x=px(157), y=px(70))
        tk.Label(header, text="Preventive maintenance tools for the Engineering Department", bg="#123B5D",
                 fg="#D8E7F2", font=("Segoe UI", 10)).place(x=px(159), y=px(113))

        content = tk.Frame(self.home, bg="#F4F7FB")
        content.pack(fill="both", expand=True, padx=px(52), pady=px(26))
        box = tk.Frame(content, bg="white", highlightbackground="#DCE4EC", highlightthickness=1)
        box.pack(fill="x", pady=(0, px(18)))
        top = tk.Frame(box, bg="white")
        top.pack(fill="x", padx=18, pady=(14, 6))
        top.columnconfigure(1, weight=1)
        tk.Label(top, text="Workspace folder", bg="white", fg="#263847",
                 font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.folder, state="readonly").grid(
            row=0, column=1, sticky="ew", padx=(16, 10), ipady=2)
        tk.Button(top, text="Browse", command=self.browse_folder, bg="#E7F0F5", fg="#1D516B",
                  activebackground="#D4E5EF", relief="flat", cursor="hand2",
                  font=("Segoe UI", 9, "bold"), padx=15, pady=5).grid(row=0, column=2)
        found = tk.Frame(box, bg="white")
        found.pack(fill="x", padx=18, pady=(0, 12))
        found.columnconfigure(0, weight=1, uniform="f")
        found.columnconfigure(1, weight=1, uniform="f")
        self.file_labels = {}
        for index, key in enumerate(workspace.LABELS):
            label = tk.Label(found, text="", bg="white", anchor="w", font=("Segoe UI", 9))
            label.grid(row=index // 2, column=index % 2, sticky="w", pady=1)
            self.file_labels[key] = label
        self.refresh_workspace()

        tk.Label(content, text="Choose a module", bg="#F4F7FB", fg="#243A4A",
                 font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))

        cards = tk.Frame(content, bg="#F4F7FB")
        cards.pack(fill="both", expand=True)
        for column in range(3):
            cards.columnconfigure(column, weight=1, uniform="card")
        self.create_module_card(cards, 0, 0, "Log Card Generator", "Create preventive maintenance log cards.",
                                "Open Log Cards", self.open_logcard, "#167D9A")
        self.create_module_card(cards, 0, 1, "Trend Report Generator", "Build monthly maintenance trend reports.",
                                "Open Trend Reports", self.open_trend, "#2D6A8E")
        self.create_module_card(cards, 0, 2, "Maintenance Schedule",
                                "Create the weekly schedule for the upcoming Friday and print due tags.",
                                "Open Schedule", self.open_schedule, "#3F8F5B")

        tk.Label(self.home, text="© 2026 RENATA PLC  •  RPGF Engineering Department", bg="#F4F7FB",
                 fg="#74818D", font=("Segoe UI", 9)).pack(pady=(0, 20))

    @staticmethod
    def create_module_card(parent, row, column, title, description, button_text, command, color):
        card = tk.Frame(parent, bg="white", highlightbackground="#DCE4EC", highlightthickness=1)
        card.grid(row=row, column=column, sticky="nsew",
                  padx=(0 if column == 0 else px(10), 0), pady=(0 if row == 0 else px(10), 0))
        tk.Frame(card, bg=color, height=px(7)).pack(fill="x")
        tk.Label(card, text=title, bg="white", fg="#163B56", justify="left", anchor="w",
                 font=("Segoe UI", 14, "bold")).pack(fill="x", padx=20, pady=(18, 6))
        tk.Button(card, text=button_text, command=command, bg="#E7F0F5", fg="#1D516B",
                  activebackground="#D4E5EF", relief="flat", cursor="hand2", font=("Segoe UI", 9, "bold"),
                  padx=10, pady=8).pack(side="bottom", anchor="w", padx=20, pady=(8, 18))
        tk.Label(card, text=description, bg="white", fg="#647482", justify="left", anchor="nw",
                 wraplength=px(230), font=("Segoe UI", 10)).pack(fill="x", padx=20)

    def open_logcard(self):
        from logcard_gui import MainWindow
        self.open_module(MainWindow, "RPGF Log Card Generator v2.0")

    def open_trend(self):
        from trend_gui import TrendWindow
        self.open_module(TrendWindow, "RPGF Trend Report Generator v2.0")

    def open_schedule(self):
        from schedule_gui import ScheduleWindow
        self.open_module(ScheduleWindow, "RPGF Maintenance Schedule v2.0")
