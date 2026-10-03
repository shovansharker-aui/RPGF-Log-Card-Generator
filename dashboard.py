"""Main dashboard for RPGF Engineering Suite v2.0."""

import sys
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk

from PIL import Image, ImageTk

from due_gui import DueWindow
from logcard_gui import MainWindow
from trend_gui import TrendWindow
from ui_common import enable_high_dpi, init_scaling, px


class Dashboard:

    def __init__(self):
        enable_high_dpi()
        self.root = tk.Tk()
        init_scaling(self.root)
        ttk.Style(self.root).theme_use("clam")
        self.root.title("RPGF Engineering Suite v2.0")
        self.root.configure(bg="#F4F7FB")
        self.center_window(980, 780)
        self.root.minsize(px(820), px(540))
        self.logo = self.load_logo()
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
        module_class(self.root, self.show_home)

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
        tk.Label(content, text="Choose a workspace", bg="#F4F7FB", fg="#243A4A",
                 font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(content, text="Create maintenance documents and performance reports from your equipment matrix.",
                 bg="#F4F7FB", fg="#6B7785", font=("Segoe UI", 10)).pack(anchor="w", pady=(3, 18))

        cards = tk.Frame(content, bg="#F4F7FB")
        cards.pack(fill="both", expand=True)
        for column in range(2):
            cards.columnconfigure(column, weight=1, uniform="card")
        for row in range(2):
            cards.rowconfigure(row, weight=1, uniform="card")
        self.create_module_card(cards, 0, 0, "Log Card Generator", "Create preventive maintenance log cards.",
                                "Open Log Cards", self.open_logcard, "#167D9A")
        self.create_module_card(cards, 0, 1, "Trend Report Generator", "Build monthly maintenance trend reports.",
                                "Open Trend Reports", self.open_trend, "#2D6A8E")
        self.create_module_card(cards, 1, 0, "Maintenance Due Labels", "Print maintenance due tags for machines.",
                                "Open Due Labels", self.open_due, "#B4690E")
        self.create_module_card(cards, 1, 1, "Log Card Responsibility", "Manage equipment ownership assignments.",
                                "Coming Soon", self.open_responsibility, "#6D7A86")

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
        tk.Label(card, text=description, bg="white", fg="#647482", justify="left", anchor="w",
                 wraplength=px(330), font=("Segoe UI", 10)).pack(fill="x", padx=20)
        tk.Button(card, text=button_text, command=command, bg="#E7F0F5", fg="#1D516B",
                  activebackground="#D4E5EF", relief="flat", cursor="hand2", font=("Segoe UI", 9, "bold"),
                  padx=10, pady=8).pack(anchor="w", padx=20, pady=(14, 18))

    def open_logcard(self):
        self.open_module(MainWindow, "RPGF Log Card Generator v2.0")

    def open_trend(self):
        self.open_module(TrendWindow, "RPGF Trend Report Generator v2.0")

    def open_due(self):
        self.open_module(DueWindow, "RPGF Maintenance Due Labels v2.0")

    @staticmethod
    def open_responsibility():
        messagebox.showinfo("Coming Soon", "Log Card Responsibility will be implemented in Version 2.2.")
