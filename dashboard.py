"""Main dashboard for RPGF Engineering Suite v2.0."""

import sys
from pathlib import Path

import tkinter as tk
from tkinter import messagebox

from PIL import Image, ImageTk

from logcard_gui import MainWindow
from trend_gui import TrendWindow


class Dashboard:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("RPGF Engineering Suite v2.0")
        self.root.geometry("860x620")
        self.root.resizable(False, False)
        self.logo = self.load_logo()
        self.build_ui()
        self.root.mainloop()

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = Path(__file__).parent
        return str(Path(base_path) / relative_path)

    def load_logo(self):
        try:
            image = Image.open(self.resource_path("assets/renata.png"))
            return ImageTk.PhotoImage(image.resize((86, 86), Image.LANCZOS))
        except Exception:
            return None

    def build_ui(self):
        self.root.configure(bg="#F4F7FB")
        header = tk.Frame(self.root, bg="#123B5D", height=180)
        header.pack(fill="x")
        header.pack_propagate(False)
        if self.logo:
            tk.Label(header, image=self.logo, bg="#123B5D").place(x=52, y=42)
        tk.Label(header, text="RENATA PLC", bg="#123B5D", fg="#B9D8E7",
                 font=("Segoe UI", 10, "bold")).place(x=158, y=48)
        tk.Label(header, text="RPGF Engineering Suite", bg="#123B5D", fg="white",
                 font=("Segoe UI", 25, "bold")).place(x=157, y=70)
        tk.Label(header, text="Preventive maintenance tools for the Engineering Department", bg="#123B5D",
                 fg="#D8E7F2", font=("Segoe UI", 10)).place(x=159, y=113)

        content = tk.Frame(self.root, bg="#F4F7FB")
        content.pack(fill="both", expand=True, padx=52, pady=30)
        tk.Label(content, text="Choose a workspace", bg="#F4F7FB", fg="#243A4A",
                 font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(content, text="Create maintenance documents and performance reports from your equipment matrix.",
                 bg="#F4F7FB", fg="#6B7785", font=("Segoe UI", 10)).pack(anchor="w", pady=(3, 18))

        cards = tk.Frame(content, bg="#F4F7FB")
        cards.pack(fill="x")
        for column in range(3):
            cards.columnconfigure(column, weight=1)
        self.create_module_card(cards, 0, "Log Card\nGenerator", "Create preventive maintenance log cards.",
                                "Open Log Cards", self.open_logcard, "#167D9A")
        self.create_module_card(cards, 1, "Trend Report\nGenerator", "Build monthly maintenance trend reports.",
                                "Open Trend Reports", self.open_trend, "#2D6A8E")
        self.create_module_card(cards, 2, "Log Card\nResponsibility", "Manage equipment ownership assignments.",
                                "Coming Soon", self.open_responsibility, "#6D7A86")

        tk.Label(self.root, text="© 2026 RENATA PLC  •  RPGF Engineering Department", bg="#F4F7FB",
                 fg="#74818D", font=("Segoe UI", 9)).pack(pady=(0, 20))

    @staticmethod
    def create_module_card(parent, column, title, description, button_text, command, color):
        card = tk.Frame(parent, bg="white", highlightbackground="#DCE4EC", highlightthickness=1)
        card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 10, 0))
        tk.Frame(card, bg=color, height=7).pack(fill="x")
        tk.Label(card, text=title, bg="white", fg="#163B56", justify="left", anchor="w",
                 font=("Segoe UI", 14, "bold")).pack(fill="x", padx=20, pady=(20, 8))
        tk.Label(card, text=description, bg="white", fg="#647482", justify="left", anchor="w",
                 wraplength=185, height=3, font=("Segoe UI", 9)).pack(fill="x", padx=20)
        tk.Button(card, text=button_text, command=command, bg="#E7F0F5", fg="#1D516B",
                  activebackground="#D4E5EF", relief="flat", cursor="hand2", font=("Segoe UI", 9, "bold"),
                  padx=10, pady=8).pack(anchor="w", padx=20, pady=(16, 20))

    def open_logcard(self):
        self.root.withdraw()
        MainWindow(dashboard=self.root)

    def open_trend(self):
        self.root.withdraw()
        TrendWindow(dashboard=self.root)

    @staticmethod
    def open_responsibility():
        messagebox.showinfo("Coming Soon", "Log Card Responsibility will be implemented in Version 2.2.")
