"""Shared visual theme for every Bluetech Quotation App window."""
import os
import sys
import tkinter as tk
from tkinter import ttk


COLORS = {
    "background": "#F3F7FC",
    "surface": "#FFFFFF",
    "navy": "#12345B",
    "blue": "#0878D1",
    "blue_hover": "#0565B3",
    "light_blue": "#E7F0FA",
    "border": "#D4E2F0",
    "text": "#17324D",
    "muted": "#667085",
    "green": "#159447",
    "green_hover": "#107638",
    "red": "#C83B3B",
}


def apply_ui_theme(window):
    """Apply shared Tk/ttk colors, typography, controls, and app icon."""
    try:
        window.configure(bg=COLORS["background"])
        window.option_add("*Font", ("Segoe UI", 9))
        window.option_add("*Listbox.background", COLORS["surface"])
        window.option_add("*Listbox.foreground", COLORS["text"])
        window.option_add("*Listbox.selectBackground", COLORS["blue"])
        window.option_add("*Listbox.selectForeground", COLORS["surface"])
        window.option_add("*Text.background", COLORS["surface"])
        window.option_add("*Text.foreground", COLORS["text"])
        window.option_add("*Text.insertBackground", COLORS["blue"])
        window.option_add("*Canvas.background", COLORS["background"])
    except tk.TclError:
        pass

    try:
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base, "BluetechComputers.ico")
        if os.path.exists(icon_path):
            window.iconbitmap(icon_path)
    except (tk.TclError, AttributeError):
        pass

    style = ttk.Style(window)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure("TFrame", background=COLORS["background"])
    style.configure("TLabel", background=COLORS["background"],
                     foreground=COLORS["text"], font=("Segoe UI", 9))
    style.configure("TLabelFrame", background=COLORS["surface"],
                     bordercolor=COLORS["border"], relief="solid", borderwidth=1)
    style.configure("TLabelFrame.Label", background=COLORS["surface"],
                     foreground=COLORS["navy"], font=("Segoe UI", 9, "bold"))
    style.configure("TEntry", font=("Segoe UI", 9), padding=(7, 5),
                     fieldbackground=COLORS["surface"], foreground=COLORS["text"],
                     bordercolor=COLORS["border"])
    style.configure("TCombobox", font=("Segoe UI", 9), padding=(6, 4),
                     fieldbackground=COLORS["surface"], foreground=COLORS["text"])
    style.map("TCombobox", fieldbackground=[("readonly", COLORS["surface"])])
    style.configure("TButton", font=("Segoe UI", 9), foreground=COLORS["text"],
                    background=COLORS["surface"], padding=(11, 7), borderwidth=1)
    style.map("TButton", background=[("active", COLORS["light_blue"])])
    style.configure("Blue.TButton", font=("Segoe UI", 9, "bold"),
                    foreground=COLORS["surface"], background=COLORS["blue"],
                    padding=(12, 7), borderwidth=0)
    style.map("Blue.TButton", background=[("active", COLORS["blue_hover"])])
    style.configure("Green.TButton", font=("Segoe UI", 9, "bold"),
                    foreground=COLORS["surface"], background=COLORS["green"],
                    padding=(12, 7), borderwidth=0)
    style.map("Green.TButton", background=[("active", COLORS["green_hover"])])
    style.configure("Light.TButton", font=("Segoe UI", 9, "bold"),
                    foreground=COLORS["navy"], background=COLORS["light_blue"],
                    padding=(11, 7), borderwidth=0)
    style.map("Light.TButton", background=[("active", "#D7E7F7")])
    style.configure("Treeview", font=("Segoe UI", 9), rowheight=28,
                    background=COLORS["surface"], fieldbackground=COLORS["surface"],
                    foreground=COLORS["text"])
    style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                    background="#E8F2FC", foreground=COLORS["navy"], padding=(7, 7))
    style.map("Treeview", background=[("selected", COLORS["blue"])],
              foreground=[("selected", COLORS["surface"])])
    style.configure("Danger.TButton", font=("Segoe UI", 9, "bold"),
                    foreground=COLORS["surface"], background=COLORS["red"],
                    padding=(11, 7), borderwidth=0)
    style.map("Danger.TButton", background=[("active", "#A92E2E")])
    style.configure("TCheckbutton", background=COLORS["background"],
                    foreground=COLORS["text"], font=("Segoe UI", 9))
    style.configure("TSeparator", background=COLORS["border"])


def add_window_header(window, title):
    """Add the same compact Bluetech brand strip to a secondary window."""
    header = tk.Frame(window, bg=COLORS["navy"], height=48)
    header.pack(side="top", fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="BLUETECH COMPUTERS", bg=COLORS["navy"],
             fg=COLORS["surface"], font=("Segoe UI", 13, "bold")).pack(
                 side="left", padx=16)
    tk.Label(header, text=title, bg=COLORS["navy"], fg="#CFE6FA",
             font=("Segoe UI", 10, "bold")).pack(side="right", padx=16)
    return header

