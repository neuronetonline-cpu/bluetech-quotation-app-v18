import os
import sqlite3
import subprocess
import sys
import webbrowser
from job_chit import open_job_chit, show_job_history, setup_db
from datetime import datetime
from urllib.parse import quote

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import tkinter.font as tkfont

try:
    import win32print
except Exception:
    win32print = None

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


APP_DIR = os.path.join(os.path.expanduser("~"), "BluetechQuotationApp")
os.makedirs(APP_DIR, exist_ok=True)
DB = os.path.join(APP_DIR, "quotations.db")

DEFAULT_PRODUCTS = [
    "MOTHER BOARD", "PROCESSOR", "CPU FAN", "RAMS", "PSU", "CASING", "CASING FANS",
    "SSD", "HDD", "VGA (used -03m)", "MONITOR (used-03m)", "ALL CABLES",
    "MOUSE", "KEYBOARD", "SPEAKER", "WIFI ADAPTER"
]

BLUE = "#075EAA"
DARK_BLUE = "#12345B"
LIGHT_BLUE = "#DCEEFF"
ROW_BLUE = "#DCEEFF"
ROW_WHITE = "#FFFFFF"
LIGHT_GREEN = "#ECF9F0"
GREEN = "#159447"
GREY = "#667085"


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def register_tk_font():
    """Load bundled Deadly Advance font into the Windows Tk process when available."""
    font_path = resource_path("Deadly Advance.ttf")
    if not os.path.exists(font_path) or not sys.platform.startswith("win"):
        return False
    try:
        import ctypes
        FR_PRIVATE = 0x10
        added = ctypes.windll.gdi32.AddFontResourceExW(font_path, FR_PRIVATE, 0)
        return bool(added)
    except Exception:
        return False


DEADLY_TK_AVAILABLE = register_tk_font()


def register_fonts():
    deadly = False
    deadly_path = resource_path("Deadly Advance.ttf")
    if os.path.exists(deadly_path):
        try:
            pdfmetrics.registerFont(TTFont("DeadlyAdvance", deadly_path))
            deadly = True
        except Exception:
            pass
    return deadly


DEADLY_ADVANCE_AVAILABLE = register_fonts()


def db():
    c = sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS quotations(
        id INTEGER PRIMARY KEY AUTOINCREMENT, qno TEXT, customer TEXT, phone TEXT,
        date TEXT, profit REAL DEFAULT 0, warranty90 REAL DEFAULT 0,
        warranty180 REAL DEFAULT 0, weight REAL DEFAULT 0, created_at TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS items(
        id INTEGER PRIMARY KEY AUTOINCREMENT, quotation_id INTEGER,
        product TEXT, description TEXT, qty REAL, cost REAL)""")

    cols = {r[1] for r in c.execute("PRAGMA table_info(quotations)").fetchall()}
    if "warranty90" not in cols:
        c.execute("ALTER TABLE quotations ADD COLUMN warranty90 REAL DEFAULT 0")
    if "warranty180" not in cols:
        c.execute("ALTER TABLE quotations ADD COLUMN warranty180 REAL DEFAULT 0")
    item_cols = {r[1] for r in c.execute("PRAGMA table_info(items)").fetchall()}
    if "cost" not in item_cols:
        c.execute("ALTER TABLE items ADD COLUMN cost REAL DEFAULT 0")
    c.execute("CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, active INTEGER DEFAULT 1)")
    c.execute("""CREATE TABLE IF NOT EXISTS invoices(
        id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_no TEXT UNIQUE, quotation_no TEXT,
        customer TEXT, phone TEXT, date TEXT, total REAL DEFAULT 0, created_at TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS invoice_items(
        id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_id INTEGER,
        product TEXT, description TEXT, qty REAL, unit_price REAL, amount REAL
    )""")
    qcols = {r[1] for r in c.execute("PRAGMA table_info(quotations)").fetchall()}
    if "prepared_by" not in qcols:
        c.execute("ALTER TABLE quotations ADD COLUMN prepared_by TEXT DEFAULT ''")
    defaults = {
        "cod_first_kg": "450",
        "cod_additional_kg": "100",
        "cod_commission": "2.5",
        "cod_min_amount": "20000",
    }
    for key, value in defaults.items():
        c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (key, value))
    c.commit()
    return c


def get_pdf_dir():
    c = db()
    row = c.execute("SELECT value FROM settings WHERE key='pdf_dir'").fetchone()
    c.close()
    folder = row[0] if row and row[0] else os.path.join(APP_DIR, "Quotations")
    os.makedirs(folder, exist_ok=True)
    return folder


def set_pdf_dir(folder):
    folder = os.path.abspath(os.path.expanduser(folder))
    os.makedirs(folder, exist_ok=True)
    c = db()
    c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('pdf_dir',?)", (folder,))
    c.commit()
    c.close()
    return folder


def get_setting(key, default=""):
    c = db()
    row = c.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    c.close()
    return row[0] if row and row[0] is not None else default


def set_setting(key, value):
    c = db()
    c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)", (key, str(value)))
    c.commit()
    c.close()


def get_users():
    c = db()
    rows = c.execute("SELECT id,name FROM users WHERE active=1 ORDER BY name COLLATE NOCASE").fetchall()
    c.close()
    return rows


def next_qno():
    c = db()
    today = datetime.now().strftime("%Y%m%d")
    rows = c.execute("SELECT qno FROM quotations WHERE qno LIKE ?", (f"QT-{today}-%",)).fetchall()
    nums = []
    for (qno,) in rows:
        try:
            nums.append(int(str(qno).rsplit("-", 1)[1]))
        except Exception:
            pass
    n = max(nums, default=0) + 1
    c.close()
    return f"QT-{today}-{n:04d}"


def next_invoice_no():
    c = db()
    today = datetime.now().strftime("%Y%m%d")
    rows = c.execute("SELECT invoice_no FROM invoices WHERE invoice_no LIKE ?", (f"INV-{today}-%",)).fetchall()
    nums = []
    for (ino,) in rows:
        try:
            nums.append(int(str(ino).rsplit("-", 1)[1]))
        except Exception:
            pass
    n = max(nums, default=0) + 1
    c.close()
    return f"INV-{today}-{n:04d}"


def money(v):
    return f"LKR {v:,.2f}"


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Bluetech Computers - Desktop Quotation")
        self.root.geometry("1320x800")
        self.root.minsize(1000, 560)
        self.rows = []
        self.editing_id = None
        setup_db(db)
        self.build()

    def build(self):
        # Polished Bluetech desktop UI. The quotation item list has its own
        # scrollbar so the calculation and action areas always remain visible.
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        self.root.configure(bg="#F3F7FC")
        style.configure("TEntry", font=("Segoe UI", 9), padding=(7, 5),
                        fieldbackground="#FFFFFF", bordercolor="#B8C7D9")
        style.configure("TCombobox", font=("Segoe UI", 9), padding=(6, 4))
        style.configure("Blue.TButton", font=("Segoe UI", 9, "bold"),
                        foreground="#FFFFFF", background="#0878D1", padding=(12, 7), borderwidth=0)
        style.map("Blue.TButton", background=[("active", "#0565B3")])
        style.configure("Green.TButton", font=("Segoe UI", 9, "bold"),
                        foreground="#FFFFFF", background="#159447", padding=(12, 7), borderwidth=0)
        style.map("Green.TButton", background=[("active", "#107638")])
        style.configure("Light.TButton", font=("Segoe UI", 9, "bold"),
                        foreground="#17324D", background="#E7F0FA", padding=(11, 7), borderwidth=0)
        style.map("Light.TButton", background=[("active", "#D7E7F7")])

        # ---------- Header ----------
        header = tk.Frame(self.root, bg="#075EAA", height=82)
        header.pack(fill="x")
        header.pack_propagate(False)

        brand = tk.Frame(header, bg="#075EAA")
        brand.pack(side="left", padx=22, pady=10)
        tk.Label(brand, text="BLUETECH", bg="#075EAA", fg="#62D3FF",
                 font=(("Deadly Advance" if DEADLY_TK_AVAILABLE else "Segoe UI"), 23, "bold")).pack(side="left")
        tk.Label(brand, text=" COMPUTERS", bg="#075EAA", fg="white",
                 font=(("Deadly Advance" if DEADLY_TK_AVAILABLE else "Segoe UI"), 23, "bold")).pack(side="left")
        tk.Label(brand, text="COMPUTER SALES  |  REPAIRS  |  ACCESSORIES   •   YOUR TECH PARTNER",
                 bg="#075EAA", fg="#D9EEFF", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=2)

        hb = tk.Frame(header, bg="#075EAA")
        hb.pack(side="right", padx=18)
        ttk.Button(hb, text="＋  New Quotation", style="Blue.TButton", command=self.new_quote).pack(side="left", padx=4)
        ttk.Button(hb, text="Quotation History", style="Blue.TButton", command=self.history).pack(side="left", padx=4)
        ttk.Button(hb, text="Job Chit History", style="Blue.TButton", command=lambda: show_job_history(self, db, get_pdf_dir)).pack(side="left", padx=4)
        ttk.Button(hb, text="⚙  Settings", style="Blue.TButton", command=self.settings).pack(side="left", padx=4)

        # ---------- Main scrollable content ----------
        # Keep the header fixed. Everything below it scrolls together with the
        # scrollbar on the far right of the application window.
        main_area = tk.Frame(self.root, bg="#F3F7FC")
        main_area.pack(fill="both", expand=True)

        self.page_canvas = tk.Canvas(main_area, bg="#F3F7FC", highlightthickness=0, bd=0)
        self.page_scroll = ttk.Scrollbar(main_area, orient="vertical", command=self.page_canvas.yview)
        self.page_content = tk.Frame(self.page_canvas, bg="#F3F7FC")
        self.page_window = self.page_canvas.create_window((0, 0), window=self.page_content, anchor="nw")
        self.page_canvas.configure(yscrollcommand=self.page_scroll.set)
        self.page_canvas.pack(side="left", fill="both", expand=True)
        self.page_scroll.pack(side="right", fill="y")

        def on_page_content_configure(_event=None):
            self.page_canvas.configure(scrollregion=self.page_canvas.bbox("all"))

        def on_page_canvas_configure(event):
            self.page_canvas.itemconfigure(self.page_window, width=event.width)

        self.page_content.bind("<Configure>", on_page_content_configure)
        self.page_canvas.bind("<Configure>", on_page_canvas_configure)
        self.page_canvas.bind_all("<MouseWheel>", self._page_mousewheel, add="+")

        # All page sections below the fixed header are placed inside this frame.
        page_parent = self.page_content

        # ---------- Section helper ----------
        def section(title, subtitle=""):
            bar = tk.Frame(page_parent, bg="#0878D1", height=34)
            bar.pack(fill="x", padx=12, pady=(8, 0))
            bar.pack_propagate(False)
            tk.Label(bar, text=title, bg="#0878D1", fg="white",
                     font=("Segoe UI", 9, "bold")).pack(side="left", padx=12)
            if subtitle:
                tk.Label(bar, text=subtitle, bg="#0878D1", fg="#DCEFFF",
                         font=("Segoe UI", 8)).pack(side="left", padx=4)
            return bar

        # ---------- Customer details ----------
        section("CUSTOMER / QUOTATION DETAILS")
        info = tk.Frame(page_parent, bg="#FFFFFF", highlightbackground="#B9D7EF",
                        highlightthickness=1, padx=12, pady=10)
        info.pack(fill="x", padx=12)

        self.qno = tk.StringVar(value=next_qno())
        self.customer = tk.StringVar()
        self.phone = tk.StringVar()
        self.qdate = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        users = [name for _, name in get_users()]
        if not users:
            c = db(); c.execute("INSERT OR IGNORE INTO users(name,active) VALUES(?,1)", ("Admin",)); c.commit(); c.close()
            users = ["Admin"]
        self.prepared_by = tk.StringVar(value=users[0])

        fields = [("Quotation No.", self.qno), ("Customer Name", self.customer),
                  ("WhatsApp / Phone", self.phone), ("Date", self.qdate)]
        for i, (lab, var) in enumerate(fields):
            col = i * 2
            info.columnconfigure(col + 1, weight=1)
            tk.Label(info, text=lab.upper(), bg="#FFFFFF", fg="#506176",
                     font=("Segoe UI", 8, "bold")).grid(row=0, column=col, sticky="w", padx=7)
            ttk.Entry(info, textvariable=var).grid(row=1, column=col, columnspan=2,
                                                    sticky="ew", padx=7, pady=(3, 4))

        tk.Label(info, text="PREPARED BY", bg="#FFFFFF", fg="#506176",
                 font=("Segoe UI", 8, "bold")).grid(row=2, column=0, sticky="w", padx=7, pady=(3, 0))
        self.prepared_combo = ttk.Combobox(info, textvariable=self.prepared_by,
                                           values=users, state="readonly")
        self.prepared_combo.grid(row=3, column=0, columnspan=2, sticky="ew", padx=7, pady=(3, 0))

        # ---------- Quotation workspace ----------
        # The quotation table and internal calculation panel share the same
        # horizontal workspace. The table uses the full available width of its
        # left panel; there is no second/inner scrollbar.
        workspace = tk.Frame(page_parent, bg="#F3F7FC")
        workspace.pack(fill="x", padx=12, pady=(8, 0))
        workspace.columnconfigure(0, weight=74)
        workspace.columnconfigure(1, weight=26)
        workspace.rowconfigure(0, weight=1)

        # ---------- Quotation items ----------
        items_panel = tk.Frame(workspace, bg="#FFFFFF", highlightbackground="#B9D7EF", highlightthickness=1)
        items_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        items_bar = tk.Frame(items_panel, bg="#0878D1", height=34)
        items_bar.pack(fill="x")
        items_bar.pack_propagate(False)
        tk.Label(items_bar, text="QUOTATION ITEMS", bg="#0878D1", fg="white",
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=12)
        tk.Label(items_bar, text="•  COST AND PROFIT ARE INTERNAL ONLY", bg="#0878D1", fg="#DCEFFF",
                 font=("Segoe UI", 8)).pack(side="left", padx=4)

        self.table = tk.Frame(items_panel, bg="#FFFFFF")
        self.table.pack(fill="x", padx=5, pady=(5, 0))

        # # small | PRODUCT medium | DESCRIPTION largest | QTY small |
        # COST small/medium | REMOVE small. These proportions stretch to
        # the complete width of the left panel.
        heads = ["#", "PRODUCT", "DESCRIPTION", "QTY", "COST (LKR)", "REMOVE"]
        for j, h in enumerate(heads):
            weight = [0, 20, 52, 8, 14, 0][j]
            minsize = [36, 150, 260, 70, 110, 70][j]
            self.table.columnconfigure(j, weight=weight, minsize=minsize)
            tk.Label(self.table, text=h, bg="#CFE6FA", fg="#12345B",
                     font=("Segoe UI", 8, "bold"), relief="solid", bd=1,
                     padx=5, pady=7).grid(row=0, column=j, sticky="nsew", padx=1, pady=1)

        self.rows = []
        for p in DEFAULT_PRODUCTS:
            self.add_row(p, silent=True)

        addbar = tk.Frame(items_panel, bg="#F3F7FC")
        addbar.pack(fill="x", padx=5, pady=(5, 6))
        ttk.Button(addbar, text="＋  ADD PRODUCT / ROW", style="Blue.TButton",
                   command=lambda: self.add_row("")).pack(side="left")
        tk.Label(addbar, text="Scroll the main quotation page when adding more rows.",
                 bg="#F3F7FC", fg="#667085", font=("Segoe UI", 8)).pack(side="left", padx=12)

        # ---------- Internal calculation ----------
        self.total_cost = tk.StringVar(value="LKR 0.00")
        self.profit = tk.StringVar(value="0")
        self.final90 = tk.StringVar(value="LKR 0.00")
        self.final180 = tk.StringVar(value="LKR 0.00")
        self.weight = tk.StringVar(value="0")
        self.cod_charge = tk.StringVar(value="LKR 0.00")
        self.cod_subtotal = tk.StringVar(value="LKR 0.00")
        self.cod_commission = tk.StringVar(value="LKR 0.00")
        self.cod_final = tk.StringVar(value="LKR 0.00")

        calc_panel = tk.Frame(workspace, bg="#FFFFFF", highlightbackground="#B9D7EF", highlightthickness=1)
        calc_panel.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        calc_bar = tk.Frame(calc_panel, bg="#0878D1", height=34)
        calc_bar.pack(fill="x")
        calc_bar.pack_propagate(False)
        tk.Label(calc_bar, text="INTERNAL CALCULATION", bg="#0878D1", fg="white",
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=12)

        calc_body = tk.Frame(calc_panel, bg="#FFFFFF", padx=8, pady=8)
        calc_body.pack(fill="both", expand=True)
        calc_body.columnconfigure(0, weight=1)
        calc_body.columnconfigure(1, weight=1)

        labels = [
            ("3 Months Final Price", self.final90, True),
            ("6 Months Final Price (+35%)", self.final180, False),
            ("Total Cost", self.total_cost, False),
            ("Requested Profit", self.profit, False),
            ("Weight (KG)", self.weight, False),
            ("COD Charge", self.cod_charge, False),
            ("COD Commission", self.cod_commission, False),
            ("COD Subtotal", self.cod_subtotal, False),
            ("Final COD Price", self.cod_final, False),
        ]

        for i, (lab, var, is_final_3m) in enumerate(labels):
            card_bg = BLUE if is_final_3m else "#F7FAFE"
            card_border = BLUE if is_final_3m else "#D4E2F0"
            card = tk.Frame(calc_body, bg=card_bg, highlightbackground=card_border,
                            highlightthickness=1, padx=8, pady=6)
            card.grid(row=i, column=0, columnspan=2, sticky="ew", pady=2)
            card.columnconfigure(1, weight=1)
            tk.Label(card, text=lab, bg=card_bg,
                     fg=("white" if is_final_3m else "#17324D"),
                     font=("Segoe UI", 8 if is_final_3m else 8, "bold"),
                     anchor="w").grid(row=0, column=0, sticky="w", padx=(2, 8))
            if is_final_3m:
                ent = tk.Entry(card, textvariable=var, justify="right",
                               font=("Segoe UI", 12, "bold"), bg=BLUE, fg="white",
                               insertbackground="white", relief="flat", bd=0,
                               highlightthickness=0)
            else:
                ent = ttk.Entry(card, textvariable=var, justify="right",
                                font=("Segoe UI", 9, "bold"))
            ent.grid(row=0, column=1, sticky="ew")
            if lab in ("Requested Profit", "Weight (KG)"):
                ent.bind("<KeyRelease>", lambda e: self.recalc())

        calc_btn = tk.Button(calc_body, text="CALCULATE", command=self.recalc,
                             bg="#0878D1", fg="white", activebackground="#0565B3",
                             activeforeground="white", font=("Segoe UI", 9, "bold"),
                             relief="flat", padx=15, pady=10, cursor="hand2")
        calc_btn.grid(row=len(labels), column=0, columnspan=2, sticky="ew", pady=(6, 0))

        # ---------- Bottom actions ----------
        actions = tk.Frame(page_parent, bg="#F3F7FC")
        actions.pack(fill="x", padx=12, pady=(6, 10))
        ttk.Button(actions, text="CLEAR", style="Light.TButton", command=self.new_quote).pack(side="left", padx=3)
        ttk.Button(actions, text="SAVE QUOTATION", style="Blue.TButton", command=self.save_quote).pack(side="right", padx=3)
        ttk.Button(actions, text="CREATE JOB CHIT", style="Green.TButton", command=lambda: open_job_chit(self, db, get_pdf_dir)).pack(side="right", padx=3)
        ttk.Button(actions, text="SAVE AS NEW QUOTATION", style="Blue.TButton", command=self.save_as_new_quote).pack(side="right", padx=3)
        ttk.Button(actions, text="PREVIEW / SAVE PDF", style="Blue.TButton", command=self.save_pdf).pack(side="right", padx=3)
        ttk.Button(actions, text="WHATSAPP QUOTATION", style="Green.TButton", command=self.whatsapp_quotation).pack(side="right", padx=3)
        inv_btn = tk.Button(actions, text="CONVERT TO INVOICE", command=self.convert_to_invoice,
                            bg="#E53935", fg="white", activebackground="#C62828", activeforeground="white",
                            font=("Segoe UI", 9, "bold"), relief="flat", padx=16, pady=8, cursor="hand2")
        inv_btn.pack(side="right", padx=3)

        self.recalc()

    def _page_mousewheel(self, event):
        """Scroll the complete quotation page when the pointer is over it."""
        try:
            x, y = self.root.winfo_pointerx(), self.root.winfo_pointery()
            widget = self.root.winfo_containing(x, y)
            if widget is None:
                return
            w = widget
            while w is not None:
                if w == self.page_canvas:
                    self.page_canvas.yview_scroll(int(-event.delta / 120), "units")
                    return "break"
                try:
                    w = w.master
                except Exception:
                    break
        except Exception:
            pass

    def _table_mousewheel(self, event):
        # Scroll only when the pointer is over the quotation-items area.
        try:
            x, y = self.root.winfo_pointerx(), self.root.winfo_pointery()
            widget = self.root.winfo_containing(x, y)
            if widget is not None:
                w = widget
                inside = False
                while w is not None:
                    if w == self.table_canvas:
                        inside = True
                        break
                    try:
                        w = w.master
                    except Exception:
                        break
                if inside:
                    self.table_canvas.yview_scroll(int(-event.delta / 120), "units")
        except Exception:
            pass

    def add_row(self, product="", silent=False):
        r = len(self.rows)
        p = tk.StringVar(value=product)
        d = tk.StringVar()
        q = tk.StringVar(value="1")
        c = tk.StringVar(value="0")
        widgets = []
        row_bg = ROW_BLUE if r % 2 == 0 else ROW_WHITE

        num_lbl = tk.Label(self.table, text=str(r + 1), bg=row_bg, fg="#667085",
                           font=("Segoe UI", 8), width=4)
        num_lbl.grid(row=r, column=0, padx=2, pady=2, sticky="nsew")

        for j, var in enumerate([p, d, q, c], start=1):
            justify = "center" if j == 2 else ("right" if j == 4 else "left")
            # Product name is intentionally bold for quick visual scanning.
            entry_font = ("Segoe UI", 9, "bold") if j == 1 else ("Segoe UI", 9)
            e = tk.Entry(self.table, textvariable=var, justify=justify,
                         font=entry_font, bg=row_bg, fg="#17324D",
                         insertbackground="#075EAA", relief="solid", bd=1,
                         highlightthickness=1, highlightbackground="#C8D6E5",
                         highlightcolor="#0878D1")
            e.grid(row=r, column=j, padx=2, pady=2, sticky="ew", ipady=3)
            widgets.append(e)
            e.bind("<KeyRelease>", lambda e: self.recalc())
            if j == 2:
                e.bind("<Return>", lambda event, widget=e: self.focus_next_row_field(widget, 1))
            elif j == 4:
                e.bind("<Return>", lambda event, widget=e: self.focus_next_row_field(widget, 3))

        btn = tk.Button(self.table, text="✕", width=4,
                        command=lambda rr=r: self.remove_row(rr),
                        bg="#FFF4F4", fg="#D92D20", activebackground="#FEE4E2",
                        activeforeground="#B42318", font=("Segoe UI", 9, "bold"),
                        relief="solid", bd=1, cursor="hand2")
        btn.grid(row=r, column=5, padx=2, pady=2, sticky="nsew")
        self.rows.append((p, d, q, c, widgets, btn, num_lbl))
        # Keep the hand-drawn layout proportions: narrow # / Qty / Cost / Remove,
        # medium Product, and the widest Description column.
        for col, (weight, minsize) in enumerate(zip(
                [0, 20, 52, 8, 14, 0],
                [36, 150, 260, 70, 110, 70])):
            self.table.columnconfigure(col, weight=weight, minsize=minsize)
        if not silent:
            self.recalc()

    def focus_next_row_field(self, widget, column):
        """Move Enter-key focus to the same field in the next visible row."""
        current_idx = None
        for idx, row in enumerate(self.rows):
            if widget in row[4]:
                current_idx = idx
                break

        if current_idx is None:
            return "break"

        next_idx = current_idx + 1
        if next_idx < len(self.rows):
            self.rows[next_idx][4][column].focus_set()
            self.rows[next_idx][4][column].selection_range(0, tk.END)
        return "break"

    def remove_row(self, idx):
        if idx >= len(self.rows):
            return
        for w in self.rows[idx][4]:
            w.destroy()
        self.rows[idx][5].destroy()
        self.rows[idx][6].destroy()
        self.rows.pop(idx)

        for r, row in enumerate(self.rows):
            row_bg = ROW_BLUE if r % 2 == 0 else ROW_WHITE
            row[6].configure(text=str(r + 1), bg=row_bg)
            row[6].grid_configure(row=r, column=0)
            for j, w in enumerate(row[4], start=1):
                w.configure(bg=row_bg)
                w.grid_configure(row=r, column=j)
            row[5].grid_configure(row=r, column=5)
        self.recalc()

    def num(self, x):
        try:
            return float(str(x).replace(",", "").replace("LKR", "").strip() or 0)
        except Exception:
            return 0

    def recalc(self):
        cost = 0
        for p, d, q, c, *_ in self.rows:
            qty = self.num(q.get())
            cost += qty * self.num(c.get())

        profit = self.num(self.profit.get())
        final90 = cost + profit
        final180 = final90 * 1.35

        weight = self.num(self.weight.get())
        first_kg = self.num(get_setting("cod_first_kg", "450"))
        additional_kg = self.num(get_setting("cod_additional_kg", "100"))
        commission_pct = self.num(get_setting("cod_commission", "2.5"))
        commission_min = self.num(get_setting("cod_min_amount", "20000"))
        if weight <= 0:
            cod_charge = 0
        else:
            import math
            extra_kg = max(0, math.ceil(weight - 1))
            cod_charge = first_kg + extra_kg * additional_kg
        cod_subtotal = final90 + cod_charge
        cod_commission = cod_subtotal * commission_pct / 100.0 if cod_subtotal > commission_min else 0
        cod_final = cod_subtotal + cod_commission

        self.total_cost.set(money(cost))
        self.final90.set(money(final90))
        self.final180.set(money(final180))
        self.cod_charge.set(money(cod_charge))
        self.cod_subtotal.set(money(cod_subtotal))
        self.cod_commission.set(money(cod_commission))
        self.cod_final.set(money(cod_final))

    def collect_items(self):
        out = []
        for p, d, q, c, *_ in self.rows:
            if p.get().strip() and self.num(q.get()) > 0:
                out.append((
                    p.get().strip(),
                    d.get().strip(),
                    self.num(q.get()),
                    self.num(c.get())
                ))
        return out

    def save_quote(self, show_message=True):
        items = self.collect_items()
        if not self.customer.get().strip():
            messagebox.showwarning("Customer", "Enter customer name.")
            return None

        qno = self.qno.get().strip()
        if not qno:
            qno = next_qno()
            self.qno.set(qno)

        self.recalc()
        c = db()

        duplicate = c.execute(
            "SELECT id FROM quotations WHERE qno=? AND id!=?",
            (qno, self.editing_id or -1)
        ).fetchone()
        if duplicate:
            c.close()
            messagebox.showerror(
                "Duplicate Quotation No.",
                f"Quotation number {qno} already exists.\n"
                "Please use a different quotation number."
            )
            return None

        values = (
            qno,
            self.customer.get().strip(),
            self.phone.get().strip(),
            self.qdate.get().strip(),
            self.num(self.profit.get()),
            self.num(self.final90.get()),
            self.num(self.final180.get()),
            self.num(self.weight.get()),
            self.prepared_by.get().strip(),
            datetime.now().isoformat()
        )

        if self.editing_id is not None:
            c.execute(
                """UPDATE quotations
                   SET qno=?,customer=?,phone=?,date=?,profit=?,
                       warranty90=?,warranty180=?,weight=?,prepared_by=?,created_at=?
                   WHERE id=?""",
                values + (self.editing_id,)
            )
            c.execute("DELETE FROM items WHERE quotation_id=?", (self.editing_id,))
            qid = self.editing_id
            action = "updated"
        else:
            c.execute(
                """INSERT INTO quotations
                   (qno,customer,phone,date,profit,warranty90,warranty180,weight,prepared_by,created_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?)""",
                values
            )
            qid = c.execute("SELECT last_insert_rowid()").fetchone()[0]
            action = "saved"

        c.executemany(
            """INSERT INTO items
               (quotation_id,product,description,qty,cost)
               VALUES(?,?,?,?,?)""",
            [(qid, *x) for x in items]
        )
        c.commit()
        c.close()

        self.editing_id = qid

        if show_message:
            messagebox.showinfo(
                "Saved",
                f"Quotation {self.qno.get()} {action}."
            )
        return qid

    def save_as_new_quote(self):
        """Save the current quotation as a brand-new quotation record."""
        if not self.customer.get().strip():
            messagebox.showwarning("Customer", "Enter customer name.")
            return None

        # Detach from any existing quotation so the original record is never updated.
        self.editing_id = None
        self.qno.set(next_qno())
        self.qdate.set(datetime.now().strftime("%Y-%m-%d"))
        return self.save_quote(show_message=True)

    def save_pdf(self, silent=False):
        self.recalc()
        items = self.collect_items()
        if not self.customer.get().strip():
            messagebox.showwarning("Customer", "Enter customer name.")
            return None

        filename = os.path.join(get_pdf_dir(), f"{self.qno.get()}.pdf")

        styles = getSampleStyleSheet()
        title = ParagraphStyle(
            "title", parent=styles["Title"], fontName="Helvetica-Bold",
            fontSize=24, leading=27, textColor=colors.HexColor(DARK_BLUE),
            alignment=TA_LEFT, spaceAfter=2
        )
        logo_font = "DeadlyAdvance" if DEADLY_ADVANCE_AVAILABLE else "Helvetica-Bold"

        subtitle = ParagraphStyle(
            "subtitle", parent=styles["BodyText"], fontSize=8.5, leading=10,
            textColor=colors.HexColor(DARK_BLUE), alignment=TA_LEFT
        )
        small = ParagraphStyle(
            "small", parent=styles["BodyText"], fontSize=7.5, leading=9.5,
            textColor=colors.HexColor(GREY)
        )
        normal = ParagraphStyle(
            "normal", parent=styles["BodyText"], fontSize=8.5, leading=11,
            textColor=colors.HexColor(DARK_BLUE)
        )
        info_style = ParagraphStyle(
            "info", parent=styles["BodyText"], fontSize=8.5, leading=12,
            textColor=colors.HexColor(DARK_BLUE)
        )
        customer_style = ParagraphStyle(
            "customer", parent=info_style, fontName="Helvetica-Bold"
        )

        doc = SimpleDocTemplate(
            filename, pagesize=A4,
            rightMargin=12 * mm, leftMargin=12 * mm,
            topMargin=10 * mm, bottomMargin=10 * mm,
            title=f"Bluetech Computers - Quotation {self.qno.get()}",
            author="Bluetech Computers",
            subject="Quotation"
        )

        story = []

        logo_style = ParagraphStyle(
            "logo", parent=title, fontName=logo_font,
            fontSize=24, leading=25,
            textColor=colors.HexColor(DARK_BLUE), alignment=TA_LEFT
        )

        header_left = [
            Paragraph("BLUETECH COMPUTERS", logo_style),
            Paragraph("Computer Sales | Repairs | Upgrades", subtitle)
        ]

        contact = Paragraph(
            "<b>077 633 7942</b><br/>"
            "<b>074 394 6233</b><br/>"
            "230,<br/>1st Floor, Lakyanya Plaza,<br/>"
            "Highlevel Road, Maharagama",
            info_style
        )

        header = Table(
            [[header_left, contact]],
            colWidths=[112 * mm, 68 * mm]
        )
        header.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LINEBELOW", (0, 0), (-1, -1), 1.1, colors.HexColor(BLUE)),
        ]))
        story.append(header)
        story.append(Spacer(1, 6))

        customer_name = self.customer.get().strip() or "-"
        qinfo = Paragraph(
            f"<b>Quotation No</b> : {self.qno.get()}<br/>"
            f"<b>Date</b> : {self.qdate.get()}<br/>"
            f"<b>Customer</b> : <font name='Helvetica-Bold'>{customer_name}</font><br/>"
            f"<b>Phone / WhatsApp</b> : {self.phone.get()}<br/>"
            f"<font size='7.5'>Prepared By : {self.prepared_by.get()}</font>",
            info_style
        )

        qtitle = Table(
            [[Paragraph("QUOTATION", title), qinfo]],
            colWidths=[105 * mm, 75 * mm]
        )
        qtitle.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#F6FAFF")),
            ("BOX", (1, 0), (1, 0), 0.7, colors.HexColor("#B8D8F5")),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(qtitle)

        story.append(Paragraph(
            "BUILD YOUR IDEAL PC WITH US",
            ParagraphStyle(
                "tag", parent=subtitle, fontSize=7.5, leading=9,
                textColor=colors.HexColor(BLUE)
            )
        ))
        story.append(Spacer(1, 6))

        data = [["#", "PRODUCT", "PRODUCT DESCRIPTION", "QTY"]]
        for i, (p, d, q, c) in enumerate(items, start=1):
            data.append([
                str(i), p, d,
                str(int(q) if float(q).is_integer() else q)
            ])

        t = Table(
            data,
            colWidths=[10 * mm, 49 * mm, 103 * mm, 18 * mm],
            repeatRows=1
        )
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(BLUE)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            # Product descriptions are slightly larger for print readability.
            ("FONTSIZE", (0, 0), (-1, -1), 9.2),
            ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#162A43")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#B7C3D0")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#F3F7FB")]),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("ALIGN", (2, 1), (2, -1), "CENTER"),
            ("ALIGN", (-1, 0), (-1, -1), "CENTER"),
        ]))
        story.append(t)
        story.append(Spacer(1, 7))

        p90 = self.num(self.final90.get())
        p180 = self.num(self.final180.get())

        # Main selling option: 3 months.
        warranty90_style = ParagraphStyle(
            "warranty90", parent=styles["BodyText"],
            fontName="Helvetica-Bold", fontSize=11.5, leading=13.5,
            textColor=colors.HexColor(DARK_BLUE), alignment=TA_LEFT
        )
        warranty180_style = ParagraphStyle(
            "warranty180", parent=styles["BodyText"],
            fontName="Helvetica-Bold", fontSize=8.2, leading=9.5,
            textColor=colors.HexColor(GREEN), alignment=TA_LEFT
        )
        price90_style = ParagraphStyle(
            "price90", parent=styles["BodyText"],
            fontName="Helvetica-Bold", fontSize=15, leading=17,
            textColor=colors.white, alignment=TA_CENTER
        )
        price180_style = ParagraphStyle(
            "price180", parent=styles["BodyText"],
            fontName="Helvetica-Bold", fontSize=9.8, leading=11,
            textColor=colors.white, alignment=TA_CENTER
        )

        w90 = Table(
            [[
                Paragraph("3 MONTHS<br/>HARDWARE WARRANTY", warranty90_style),
                Paragraph(money(p90), price90_style)
            ]],
            colWidths=[66 * mm, 46 * mm]
        )
        w90.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(LIGHT_BLUE)),
            ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#B9DBF8")),
            ("BACKGROUND", (1, 0), (1, 0), colors.HexColor(BLUE)),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))

        # Wider price cell prevents amounts such as LKR 70,132.50 wrapping.
        w180 = Table(
            [[
                Paragraph("6 MONTHS<br/>HARDWARE WARRANTY", warranty180_style),
                Paragraph(money(p180), price180_style)
            ]],
            colWidths=[38 * mm, 30 * mm]
        )
        w180.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(LIGHT_GREEN)),
            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#BEE7CB")),
            ("BACKGROUND", (1, 0), (1, 0), colors.HexColor(GREEN)),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]))

        warranty_row = Table(
            [[w90, w180]],
            colWidths=[112 * mm, 68 * mm]
        )
        warranty_row.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(warranty_row)
        story.append(Spacer(1, 7))

        terms = Paragraph(
            "<b>Terms & Conditions</b><br/>"
            "• Quotation Validity: Prices are valid for 2 days from the quotation date and time.<br/>"
            "• Warranty: Warranty covers MANUFACTURER FAULTS ONLY. Physical damage, burns, liquid damage, and other external damages are not covered.<br/>"
            "• Stock Availability: Product availability is subject to change without prior notice.<br/>"
            "• Support: For further information or assistance, please contact us by phone or WhatsApp.<br/>"
            "• COD: Courier charges must be paid to our bank account before dispatch.",
            small
        )

        terms_box = Table(
            [[
                terms,
                Paragraph(
                    "<b>Thank you<br/>for your business!</b>",
                    ParagraphStyle(
                        "thanks", parent=normal, fontSize=11,
                        leading=14, alignment=TA_CENTER
                    )
                )
            ]],
            colWidths=[126 * mm, 54 * mm]
        )
        terms_box.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F9FE")),
            ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#C7D8EA")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]))
        story.append(terms_box)
        story.append(Spacer(1, 7))

        footer = Paragraph(
            "Facebook  |  TikTok  |  Google Reviews<br/>"
            "QUALITY PARTS  |  TRUSTED SERVICE  |  BETTER COMPUTING",
            ParagraphStyle(
                "footer", parent=small, alignment=TA_CENTER,
                fontSize=7.3, leading=9
            )
        )
        story.append(footer)

        doc.build(story)

        try:
            if sys.platform.startswith("win"):
                os.startfile(filename)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", filename])
            else:
                subprocess.Popen(["xdg-open", filename])
        except Exception:
            pass

        if not silent:
            messagebox.showinfo(
                "PDF Created",
                f"PDF created:\n{filename}\n\n"
                "You can use the WhatsApp button to open the customer's chat."
            )
        return filename

    def convert_to_invoice(self):
        """Open a separate invoice window populated from the current quotation."""
        self.recalc()
        items = self.collect_items()
        if not self.customer.get().strip():
            messagebox.showwarning("Customer", "Enter customer name.")
            return
        if not items:
            messagebox.showwarning("Invoice", "Add at least one product before converting to invoice.")
            return

        # Allocate the 3-month quotation total across lines in proportion to internal cost.
        # The user can edit each invoice unit price before printing.
        quote_total = self.num(self.final90.get())
        raw_costs = [max(0.0, self.num(c)) for _, _, q, c in items]
        cost_total = sum(self.num(q) * c for (_, _, q, c) in items)
        if cost_total <= 0:
            shares = [quote_total / len(items)] * len(items)
        else:
            shares = [(self.num(q) * c / cost_total) * quote_total for (_, _, q, c) in items]

        win = tk.Toplevel(self.root)
        win.title("Bluetech Computers - Invoice")
        win.geometry("1050x700")
        win.minsize(900, 600)

        top = ttk.Frame(win, padding=12)
        top.pack(fill="x")
        ttk.Label(top, text="BLUETECH COMPUTERS", font=("Segoe UI", 18, "bold")).pack(side="left")
        ttk.Label(top, text="INVOICE", font=("Segoe UI", 18, "bold")).pack(side="right")

        info = ttk.LabelFrame(win, text="Invoice Details", padding=10)
        info.pack(fill="x", padx=12, pady=5)
        invoice_no = tk.StringVar(value=next_invoice_no())
        invoice_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        customer = tk.StringVar(value=self.customer.get().strip())
        phone = tk.StringVar(value=self.phone.get().strip())
        fields = [
            ("Invoice No.", invoice_no), ("Customer Name", customer),
            ("WhatsApp / Phone", phone), ("Date", invoice_date)
        ]
        for i, (lab, var) in enumerate(fields):
            ttk.Label(info, text=lab).grid(row=0, column=i*2, sticky="w", padx=5)
            ttk.Entry(info, textvariable=var, width=24).grid(row=0, column=i*2+1, sticky="ew", padx=5)
            info.columnconfigure(i*2+1, weight=1)

        box = ttk.LabelFrame(win, text="Invoice Items", padding=8)
        box.pack(fill="both", expand=True, padx=12, pady=5)
        headers = ["PRODUCT", "DESCRIPTION", "QTY", "UNIT PRICE", "AMOUNT"]
        for j, h in enumerate(headers):
            ttk.Label(box, text=h, font=("Segoe UI", 9, "bold")).grid(row=0, column=j, padx=4, pady=4, sticky="ew")
        for j, w in enumerate((24, 38, 10, 18, 18)):
            box.columnconfigure(j, weight=1, minsize=w*7)

        invoice_rows = []
        for idx, ((prod, desc, qty, _cost), share) in enumerate(zip(items, shares), start=1):
            pv, dv = tk.StringVar(value=prod), tk.StringVar(value=desc)
            qv = tk.StringVar(value=str(int(qty) if float(qty).is_integer() else qty))
            unit_default = share / self.num(qty) if self.num(qty) else share
            uv = tk.StringVar(value=f"{unit_default:.2f}")
            av = tk.StringVar(value="LKR 0.00")
            ttk.Entry(box, textvariable=pv).grid(row=idx, column=0, padx=2, pady=2, sticky="ew")
            ttk.Entry(box, textvariable=dv, justify="center").grid(row=idx, column=1, padx=2, pady=2, sticky="ew")
            ttk.Entry(box, textvariable=qv).grid(row=idx, column=2, padx=2, pady=2, sticky="ew")
            ttk.Entry(box, textvariable=uv, justify="right").grid(row=idx, column=3, padx=2, pady=2, sticky="ew")
            ttk.Label(box, textvariable=av, anchor="e").grid(row=idx, column=4, padx=5, pady=2, sticky="ew")
            invoice_rows.append((pv, dv, qv, uv, av))

        total_var = tk.StringVar(value="LKR 0.00")
        ttk.Label(box, text="TOTAL", font=("Segoe UI", 10, "bold")).grid(row=len(invoice_rows)+1, column=3, sticky="e", padx=5, pady=10)
        ttk.Label(box, textvariable=total_var, font=("Segoe UI", 11, "bold"), anchor="e").grid(row=len(invoice_rows)+1, column=4, sticky="ew", padx=5, pady=10)

        def calc_invoice(*_):
            total = 0.0
            for pv, dv, qv, uv, av in invoice_rows:
                qty = self.num(qv.get())
                unit = self.num(uv.get())
                amount = qty * unit
                total += amount
                av.set(money(amount))
            total_var.set(money(total))

        for row in invoice_rows:
            row[2].trace_add("write", calc_invoice)
            row[3].trace_add("write", calc_invoice)
        calc_invoice()

        note = ttk.Label(win, text="English invoice format for Epson LQ-300+ continuous paper.", foreground=GREY)
        note.pack(anchor="w", padx=15, pady=(0, 4))

        actions = ttk.Frame(win, padding=10)
        actions.pack(fill="x")

        def print_invoice():
            calc_invoice()
            if win32print is None:
                messagebox.showerror("Printer", "Windows printer support is not available. Install the app with the updated requirements first.", parent=win)
                return
            printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            if not printers:
                messagebox.showerror("Printer", "No Windows printers were found.", parent=win)
                return
            printer_win = tk.Toplevel(win)
            printer_win.title("Select Printer")
            printer_win.geometry("520x180")
            ttk.Label(printer_win, text="Printer").pack(anchor="w", padx=15, pady=(15, 5))
            try:
                default_printer = win32print.GetDefaultPrinter()
            except Exception:
                default_printer = printers[0]
            printer_var = tk.StringVar(value=default_printer if default_printer in printers else printers[0])
            combo = ttk.Combobox(printer_win, textvariable=printer_var, values=printers, state="readonly", width=58)
            combo.pack(padx=15, fill="x")

            def do_print():
                printer_name = printer_var.get()
                lines = []
                def line(txt=""):
                    lines.append(str(txt)[:80])
                line("BLUETECH COMPUTERS")
                line("COMPUTER SALES | REPAIRS | UPGRADES")
                line("230, 1st Floor, Lakyanya Plaza, Highlevel Road, Maharagama")
                line("077 633 7942 / 074 394 6233")
                line("=" * 80)
                line(f"INVOICE: {invoice_no.get():<22} DATE: {invoice_date.get()}")
                line(f"CUSTOMER: {customer.get()[:60]}")
                line(f"PHONE   : {phone.get()[:60]}")
                line("-" * 80)
                line(f"{'PRODUCT':<22} {'DESCRIPTION':<27} {'QTY':>5} {'UNIT PRICE':>11} {'AMOUNT':>12}")
                line("-" * 80)
                total = 0.0
                for pv, dv, qv, uv, av in invoice_rows:
                    qty = self.num(qv.get()); unit = self.num(uv.get()); amount = qty * unit; total += amount
                    product_text = pv.get().strip()[:22]
                    desc_text = dv.get().strip()[:27]
                    line(f"{product_text:<22} {desc_text:<27} {qty:>5g} {unit:>11.2f} {amount:>12.2f}")
                line("-" * 80)
                line(f"{'TOTAL':>68} {total:>12.2f}")
                line("=" * 80)
                line("Warranty: As stated on the quotation / invoice.")
                line("Thank you for your business!")
                data = "\r\n".join(lines) + "\r\n\f"
                try:
                    h = win32print.OpenPrinter(printer_name)
                    try:
                        win32print.StartDocPrinter(h, 1, (invoice_no.get(), None, "RAW"))
                        win32print.StartPagePrinter(h)
                        win32print.WritePrinter(h, data.encode("cp437", errors="replace"))
                        win32print.EndPagePrinter(h)
                        win32print.EndDocPrinter(h)
                    finally:
                        win32print.ClosePrinter(h)
                    messagebox.showinfo("Invoice", f"Invoice sent to {printer_name}.", parent=printer_win)
                    printer_win.destroy()
                except Exception as e:
                    messagebox.showerror("Printer", f"Could not print invoice:\n{e}", parent=printer_win)

            ttk.Button(printer_win, text="PRINT", command=do_print).pack(pady=18)

        ttk.Button(actions, text="PRINT INVOICE", command=print_invoice).pack(side="right", padx=5)
        ttk.Button(actions, text="CLOSE", command=win.destroy).pack(side="right", padx=5)

    def whatsapp_quotation(self):
        if not self.customer.get().strip():
            messagebox.showwarning("Customer", "Enter customer name.")
            return

        phone = "".join(ch for ch in self.phone.get() if ch.isdigit())
        if phone.startswith("0"):
            phone = "94" + phone[1:]
        elif phone.startswith("94"):
            pass

        if not phone:
            messagebox.showwarning("WhatsApp", "Enter the customer's WhatsApp / phone number.")
            return

        self.recalc()
        if self.editing_id is None:
            if self.save_quote(show_message=False) is None:
                return

        p90 = self.num(self.final90.get())
        p180 = self.num(self.final180.get())
        message = (
            f"Hello {self.customer.get().strip()},\n\n"
            f"Quotation No: {self.qno.get()}\n"
            f"Date: {self.qdate.get()}\n\n"
            f"3 Months Hardware Warranty: {money(p90)}\n"
            f"6 Months Hardware Warranty: {money(p180)}\n\n"
            "Thank you for choosing Bluetech Computers.\n"
            "Computer Sales | Repairs | Upgrades\n"
            "077 633 7942 / 074 394 6233"
        )

        url = f"https://wa.me/{phone}?text={quote(message)}"
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("WhatsApp", f"Could not open WhatsApp:\n{e}")

    def settings(self):
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.geometry("760x520")
        win.resizable(False, False)

        ttk.Label(win, text="PDF / Quotation Save Location", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=18, pady=(18, 8))
        row = ttk.Frame(win)
        row.pack(fill="x", padx=18)
        path_var = tk.StringVar(value=get_pdf_dir())
        entry = ttk.Entry(row, textvariable=path_var)
        entry.pack(side="left", fill="x", expand=True)

        def choose():
            folder = filedialog.askdirectory(
                title="Choose quotation save folder",
                initialdir=path_var.get() if os.path.isdir(path_var.get()) else APP_DIR,
                parent=win
            )
            if folder:
                folder = os.path.normpath(os.path.abspath(folder))
                os.makedirs(folder, exist_ok=True)
                path_var.set(folder)
                entry.delete(0, "end")
                entry.insert(0, folder)
                status_var.set("Selected folder: " + folder)

        ttk.Button(row, text="Browse...", command=choose).pack(side="left", padx=(8, 0))
        status_var = tk.StringVar(value="Current save folder: " + get_pdf_dir())
        ttk.Label(win, textvariable=status_var, foreground=GREY, wraplength=700).pack(anchor="w", padx=18, pady=8)

        ttk.Separator(win).pack(fill="x", padx=18, pady=8)
        ttk.Label(win, text="Manage Users", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=18, pady=(4, 8))
        user_frame = ttk.Frame(win)
        user_frame.pack(fill="x", padx=18)
        user_list = tk.Listbox(user_frame, height=6)
        user_list.pack(side="left", fill="x", expand=True)
        for _, name in get_users():
            user_list.insert("end", name)

        def refresh_users():
            user_list.delete(0, "end")
            for _, name in get_users():
                user_list.insert("end", name)

        def add_user():
            name = simpledialog.askstring("Add User", "User name:", parent=win)
            if name and name.strip():
                try:
                    c = db(); c.execute("INSERT INTO users(name,active) VALUES(?,1)", (name.strip(),)); c.commit(); c.close()
                    refresh_users()
                    status_var.set("User added: " + name.strip())
                    self.refresh_prepared_users()
                except sqlite3.IntegrityError:
                    messagebox.showwarning("Users", "That user already exists.", parent=win)

        def delete_user():
            sel = user_list.curselection()
            if not sel:
                messagebox.showwarning("Users", "Select a user first.", parent=win); return
            name = user_list.get(sel[0])
            if name == self.prepared_by.get():
                messagebox.showwarning("Users", "Select another Prepared By user before deleting this user.", parent=win); return
            c = db(); c.execute("UPDATE users SET active=0 WHERE name=?", (name,)); c.commit(); c.close()
            refresh_users(); self.refresh_prepared_users()

        ub = ttk.Frame(win); ub.pack(pady=6)
        ttk.Button(ub, text="ADD USER", command=add_user).pack(side="left", padx=4)
        ttk.Button(ub, text="DELETE USER", command=delete_user).pack(side="left", padx=4)

        ttk.Separator(win).pack(fill="x", padx=18, pady=8)
        ttk.Label(win, text="COD Settings", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=18, pady=(4, 8))
        cod = ttk.Frame(win); cod.pack(fill="x", padx=18)
        first_var = tk.StringVar(value=get_setting("cod_first_kg", "450"))
        add_var = tk.StringVar(value=get_setting("cod_additional_kg", "100"))
        comm_var = tk.StringVar(value=get_setting("cod_commission", "2.5"))
        min_var = tk.StringVar(value=get_setting("cod_min_amount", "20000"))
        for i, (label, var) in enumerate([("1st KG Charge", first_var), ("Additional KG Charge", add_var), ("COD Commission %", comm_var), ("Commission Minimum Amount", min_var)]):
            ttk.Label(cod, text=label).grid(row=0, column=i, padx=5, sticky="w")
            ttk.Entry(cod, textvariable=var, width=18).grid(row=1, column=i, padx=5, sticky="ew")
        ttk.Label(win, text="COD is calculated internally only; it is not shown on customer PDFs.", foreground=GREY).pack(anchor="w", padx=18, pady=8)

        buttons = ttk.Frame(win); buttons.pack(pady=10)
        def save():
            try:
                if self.num(first_var.get()) < 0 or self.num(add_var.get()) < 0 or self.num(comm_var.get()) < 0 or self.num(min_var.get()) < 0:
                    raise ValueError("COD values cannot be negative.")
                selected_folder = os.path.normpath(os.path.abspath(path_var.get().strip()))
                if not selected_folder:
                    raise ValueError("Please choose a PDF save folder.")
                os.makedirs(selected_folder, exist_ok=True)
                set_pdf_dir(selected_folder)
                set_setting("cod_first_kg", self.num(first_var.get()))
                set_setting("cod_additional_kg", self.num(add_var.get()))
                set_setting("cod_commission", self.num(comm_var.get()))
                set_setting("cod_min_amount", self.num(min_var.get()))
                self.recalc()
                self.refresh_prepared_users()
                messagebox.showinfo("Settings", "Settings saved.", parent=win)
                win.destroy()
            except Exception as e:
                messagebox.showerror("Settings", f"Could not save settings:\n{e}", parent=win)
        ttk.Button(buttons, text="SAVE", command=save).pack(side="left", padx=5)
        ttk.Button(buttons, text="CANCEL", command=win.destroy).pack(side="left", padx=5)

    def refresh_prepared_users(self):
        if not hasattr(self, "prepared_combo"):
            return
        users = [name for _, name in get_users()]
        self.prepared_combo["values"] = users
        if self.prepared_by.get() not in users and users:
            self.prepared_by.set(users[0])

    def new_quote(self):
        self.editing_id = None
        for w in self.root.winfo_children():
            w.destroy()
        self.rows = []
        self.build()

    def history(self):
        win = tk.Toplevel(self.root)
        win.title("Quotation History")
        win.geometry("1160x650")

        search_var = tk.StringVar()
        search_row = ttk.Frame(win, padding=10)
        search_row.pack(fill="x")

        ttk.Label(search_row, text="Search:").pack(side="left", padx=(0, 6))
        search_entry = ttk.Entry(
            search_row, textvariable=search_var, width=55
        )
        search_entry.pack(side="left", fill="x", expand=True)
        ttk.Label(
            search_row,
            text="Name / Quotation No. / Phone / Date"
        ).pack(side="left", padx=10)

        tree = ttk.Treeview(
            win,
            columns=("q", "customer", "phone", "date", "profit", "p90", "p180"),
            show="headings"
        )
        headings = (
            "Quotation No.", "Customer", "Phone", "Date",
            "Requested Profit", "3 Months", "6 Months"
        )
        widths = (155, 190, 135, 105, 135, 135, 135)

        for col, h, width in zip(tree["columns"], headings, widths):
            tree.heading(col, text=h)
            tree.column(col, width=width)

        tree.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        c = db()
        rows = c.execute(
            """SELECT id,qno,customer,phone,date,profit,warranty90,warranty180,prepared_by
               FROM quotations ORDER BY id DESC"""
        ).fetchall()
        c.close()

        def refresh(*_):
            term = search_var.get().strip().lower()
            for item in tree.get_children():
                tree.delete(item)

            for row in rows:
                qid, qno, customer, phone, date, profit, p90, p180, prepared_by = row
                hay = " ".join([
                    str(qno or ""), str(customer or ""),
                    str(phone or ""), str(date or "")
                ]).lower()

                if term and term not in hay:
                    continue

                tree.insert(
                    "", "end", iid=str(qid),
                    values=(
                        qno, customer, phone, date,
                        money(profit), money(p90), money(p180)
                    )
                )

        search_var.trace_add("write", refresh)
        refresh()
        search_entry.focus_set()

        ttk.Label(
            win,
            text="Double-click a quotation to open and edit it."
        ).pack(pady=(0, 4))

        btns = ttk.Frame(win)
        btns.pack(pady=6)

        ttk.Button(
            btns, text="OPEN / EDIT SELECTED",
            command=lambda: self.load_history_item(tree, win)
        ).pack(side="left", padx=5)

        ttk.Button(
            btns, text="REPRINT PDF",
            command=lambda: self.reprint_history_item(tree, win)
        ).pack(side="left", padx=5)

        ttk.Button(
            btns, text="WHATSAPP",
            command=lambda: self.whatsapp_history_item(tree, win)
        ).pack(side="left", padx=5)

        ttk.Button(
            btns, text="Close",
            command=win.destroy
        ).pack(side="left", padx=5)

        tree.bind("<Double-1>", lambda e: self.load_history_item(tree, win))

    def get_history_record(self, tree, win):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning(
                "History", "Select a quotation first.", parent=win
            )
            return None

        qid = int(selected[0])
        c = db()
        q = c.execute(
            """SELECT id,qno,customer,phone,date,profit,
                      warranty90,warranty180,weight,prepared_by
               FROM quotations WHERE id=?""",
            (qid,)
        ).fetchone()
        items = c.execute(
            """SELECT product,description,qty,cost
               FROM items WHERE quotation_id=? ORDER BY id""",
            (qid,)
        ).fetchall()
        c.close()

        if not q:
            messagebox.showerror(
                "History", "Quotation could not be loaded.", parent=win
            )
            return None

        return q, items

    def load_history_item(self, tree, win):
        record = self.get_history_record(tree, win)
        if not record:
            return

        q, items = record
        self.editing_id = q[0]
        self.qno.set(q[1])
        self.customer.set(q[2])
        self.phone.set(q[3])
        self.qdate.set(q[4])
        self.profit.set(str(q[5] or 0))
        self.weight.set(str(q[8] or 0))
        self.prepared_by.set(q[9] or self.prepared_by.get())
        self.refresh_prepared_users()

        for row in self.rows:
            for w in row[4]:
                w.destroy()
            row[5].destroy()
        self.rows = []

        for p, d, qty, cost in items:
            self.add_row(p, silent=True)
            row = self.rows[-1]
            row[1].set(d or "")
            row[2].set(str(int(qty) if float(qty).is_integer() else qty))
            row[3].set(str(cost or 0))

        if not items:
            self.add_row("", silent=True)

        self.recalc()
        win.destroy()
        self.root.lift()
        self.root.focus_force()

    def save_history_item_as_new(self, tree, win):
        record = self.get_history_record(tree, win)
        if not record:
            return
        q, items = record
        # Load selected quotation into the editor, but deliberately detach it from the old DB id.
        self.editing_id = None
        self.qno.set(next_qno())
        self.customer.set(q[2] or "")
        self.phone.set(q[3] or "")
        self.qdate.set(datetime.now().strftime("%Y-%m-%d"))
        self.profit.set(str(q[5] or 0))
        self.weight.set(str(q[8] or 0))
        self.prepared_by.set(q[9] or self.prepared_by.get())
        self.refresh_prepared_users()

        for row in self.rows:
            for w in row[4]:
                w.destroy()
            row[5].destroy()
        self.rows = []
        for p, d, qty, cost in items:
            self.add_row(p, silent=True)
            row = self.rows[-1]
            row[1].set(d or "")
            row[2].set(str(int(qty) if float(qty).is_integer() else qty))
            row[3].set(str(cost or 0))
        if not items:
            self.add_row("", silent=True)
        self.recalc()
        win.destroy()
        self.root.lift()
        self.root.focus_force()
        messagebox.showinfo("Save As New", f"New quotation {self.qno.get()} is ready.\nEdit the details if needed, then click SAVE QUOTATION.")

    def reprint_history_item(self, tree, win):
        record = self.get_history_record(tree, win)
        if not record:
            return
        q, items = record

        self.editing_id = q[0]
        self.qno.set(q[1])
        self.customer.set(q[2])
        self.phone.set(q[3])
        self.qdate.set(q[4])
        self.profit.set(str(q[5] or 0))
        self.weight.set(str(q[8] or 0))
        self.prepared_by.set(q[9] or self.prepared_by.get())
        self.refresh_prepared_users()

        for row in self.rows:
            for w in row[4]:
                w.destroy()
            row[5].destroy()
        self.rows = []

        for p, d, qty, cost in items:
            self.add_row(p, silent=True)
            row = self.rows[-1]
            row[1].set(d or "")
            row[2].set(str(int(qty) if float(qty).is_integer() else qty))
            row[3].set(str(cost or 0))

        if not items:
            self.add_row("", silent=True)

        self.recalc()
        win.destroy()
        self.save_pdf()

    def whatsapp_history_item(self, tree, win):
        record = self.get_history_record(tree, win)
        if not record:
            return
        q, items = record

        self.editing_id = q[0]
        self.qno.set(q[1])
        self.customer.set(q[2])
        self.phone.set(q[3])
        self.qdate.set(q[4])
        self.profit.set(str(q[5] or 0))
        self.weight.set(str(q[8] or 0))
        self.prepared_by.set(q[9] or self.prepared_by.get())
        self.refresh_prepared_users()

        for row in self.rows:
            for w in row[4]:
                w.destroy()
            row[5].destroy()
        self.rows = []

        for p, d, qty, cost in items:
            self.add_row(p, silent=True)
            row = self.rows[-1]
            row[1].set(d or "")
            row[2].set(str(int(qty) if float(qty).is_integer() else qty))
            row[3].set(str(cost or 0))

        if not items:
            self.add_row("", silent=True)

        self.recalc()
        win.destroy()
        self.whatsapp_quotation()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
