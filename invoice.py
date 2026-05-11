"""
Professional Invoice Generator - Full Featured
Libraries Required:
  pip install reportlab pillow openpyxl matplotlib
  tkinter is built into Python standard library
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser
import sqlite3
import json
import os
import io
import datetime
import shutil
from PIL import Image, ImageTk, ImageDraw
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.pdfgen import canvas as rl_canvas

# ─── CONSTANTS ──────────────────────────────────────────────────────────────
DB_PATH = os.path.join(os.path.expanduser("~"), "invoices.db")
LOGO_PATH = os.path.join(os.path.expanduser("~"), "company_logo.png")
APP_BG = "#0f1117"
CARD_BG = "#1a1d27"
ACCENT = "#6c63ff"
ACCENT2 = "#f5a623"
SUCCESS = "#2ecc71"
DANGER = "#e74c3c"
TEXT = "#e8e8f0"
SUBTEXT = "#8888aa"
BORDER = "#2a2d3e"

# ─── DATABASE ────────────────────────────────────────────────────────────────
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT, address TEXT, phone TEXT, email TEXT,
        website TEXT, tax_id TEXT, logo_path TEXT,
        bank_name TEXT, account_no TEXT, currency TEXT DEFAULT 'USD'
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, email TEXT, phone TEXT,
        address TEXT, city TEXT, country TEXT, notes TEXT
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_no TEXT UNIQUE, customer_id INTEGER,
        issue_date TEXT, due_date TEXT,
        subtotal REAL, tax_rate REAL, tax_amount REAL,
        discount REAL, total REAL,
        status TEXT DEFAULT 'Unpaid',
        notes TEXT, terms TEXT,
        created_at TEXT,
        FOREIGN KEY(customer_id) REFERENCES customers(id)
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS invoice_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_id INTEGER, description TEXT,
        quantity REAL, unit_price REAL, total REAL,
        FOREIGN KEY(invoice_id) REFERENCES invoices(id)
    )""")
    conn.commit()
    conn.close()

def get_conn():
    return sqlite3.connect(DB_PATH)

# ─── HELPERS ─────────────────────────────────────────────────────────────────
def make_round_logo(src_path, size=120):
    """Crop image into circle, return PIL Image."""
    img = Image.open(src_path).convert("RGBA").resize((size, size), Image.LANCZOS)
    mask = Image.new("L", (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse((0, 0, size, size), fill=255)
    result = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    result.paste(img, (0, 0), mask)
    return result

def next_invoice_number():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM invoices")
    n = c.fetchone()[0] + 1
    conn.close()
    return f"INV-{datetime.datetime.now().year}-{n:04d}"

def styled_btn(parent, text, cmd, bg=ACCENT, fg="white", width=18, **kw):
    b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                  font=("Segoe UI", 10, "bold"), relief="flat",
                  activebackground=bg, activeforeground=fg,
                  cursor="hand2", width=width, pady=6, **kw)
    b.bind("<Enter>", lambda e: b.config(bg=_lighten(bg)))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b

def _lighten(hex_color):
    r = min(255, int(hex_color[1:3], 16) + 30)
    g = min(255, int(hex_color[3:5], 16) + 30)
    b = min(255, int(hex_color[5:7], 16) + 30)
    return f"#{r:02x}{g:02x}{b:02x}"

def lbl(parent, text, size=10, bold=False, color=TEXT, **kw):
    weight = "bold" if bold else "normal"
    return tk.Label(parent, text=text, font=("Segoe UI", size, weight),
                    bg=parent["bg"] if hasattr(parent, "winfo_class") else CARD_BG,
                    fg=color, **kw)

def entry(parent, width=30, **kw):
    e = tk.Entry(parent, font=("Segoe UI", 10), bg="#22253a", fg=TEXT,
                 insertbackground=TEXT, relief="flat", bd=0,
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT, width=width, **kw)
    return e

def card(parent, **kw):
    f = tk.Frame(parent, bg=CARD_BG, bd=0, relief="flat", **kw)
    return f

# ─── PDF GENERATOR ───────────────────────────────────────────────────────────
def generate_pdf(invoice_id, save_path):
    conn = get_conn()
    c = conn.cursor()
    c.execute("""SELECT i.*, cu.name, cu.email, cu.phone, cu.address, cu.city, cu.country
                 FROM invoices i JOIN customers cu ON i.customer_id=cu.id
                 WHERE i.id=?""", (invoice_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return
    cols = [d[0] for d in c.description]
    inv = dict(zip(cols, row))

    c.execute("SELECT * FROM invoice_items WHERE invoice_id=?", (invoice_id,))
    items = c.fetchall()

    c.execute("SELECT * FROM companies LIMIT 1")
    co_row = c.fetchone()
    co_cols = [d[0] for d in c.description]
    company = dict(zip(co_cols, co_row)) if co_row else {}
    conn.close()

    doc = SimpleDocTemplate(save_path, pagesize=A4,
                            rightMargin=15*mm, leftMargin=15*mm,
                            topMargin=15*mm, bottomMargin=15*mm)
    styles = getSampleStyleSheet()
    story = []

    accent_color = colors.HexColor("#6c63ff")
    dark_color   = colors.HexColor("#0f1117")
    light_gray   = colors.HexColor("#f0f0f8")
    mid_gray     = colors.HexColor("#ccccdd")
    white        = colors.white

    title_style = ParagraphStyle("Title2", fontSize=28, textColor=accent_color,
                                 fontName="Helvetica-Bold", spaceAfter=2)
    sub_style   = ParagraphStyle("Sub", fontSize=10, textColor=colors.HexColor("#888888"))
    bold_style  = ParagraphStyle("Bold", fontSize=11, fontName="Helvetica-Bold")
    normal_style= ParagraphStyle("Normal2", fontSize=10)
    right_style = ParagraphStyle("Right", fontSize=10, alignment=TA_RIGHT)
    big_total   = ParagraphStyle("BigTotal", fontSize=16, fontName="Helvetica-Bold",
                                 textColor=accent_color, alignment=TA_RIGHT)

    # ── HEADER ──
    header_data = [[
        Paragraph(f"<b>INVOICE</b>", title_style),
        Paragraph(f"<b>{company.get('name','Your Company')}</b>", bold_style)
    ],[
        Paragraph(f"#{inv['invoice_no']}", sub_style),
        Paragraph(company.get('address',''), normal_style)
    ],[
        Paragraph("", normal_style),
        Paragraph(f"{company.get('phone','')}  |  {company.get('email','')}", sub_style)
    ],[
        Paragraph("", normal_style),
        Paragraph(f"{company.get('website','')}", sub_style)
    ]]

    # Add logo if exists
    logo_cell = ""
    logo_path = company.get("logo_path","") or LOGO_PATH
    if logo_path and os.path.exists(logo_path):
        try:
            round_img = make_round_logo(logo_path, 100)
            buf = io.BytesIO()
            round_img.save(buf, format="PNG")
            buf.seek(0)
            logo_cell = RLImage(buf, width=25*mm, height=25*mm)
        except:
            logo_cell = ""

    # Company + logo side by side
    co_block = [
        [Paragraph(f"<b>{company.get('name','Your Company')}</b>", bold_style), logo_cell],
        [Paragraph(company.get('address',''), normal_style), ""],
        [Paragraph(f"{company.get('phone','')} | {company.get('email','')}", sub_style), ""],
        [Paragraph(company.get('website',''), sub_style), ""],
    ]
    co_table = Table(co_block, colWidths=[100*mm, 30*mm])
    co_table.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("TOPPADDING", (0,0), (-1,-1), 2),
    ]))

    top_table = Table([
        [Paragraph("<b>INVOICE</b>", title_style), co_table],
        [Paragraph(f"<font color='#888888'>#{inv['invoice_no']}</font>", ParagraphStyle("x", fontSize=11)), ""],
    ], colWidths=[80*mm, 130*mm])
    top_table.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
    ]))
    story.append(top_table)
    story.append(HRFlowable(width="100%", thickness=2, color=accent_color, spaceAfter=8))

    # ── BILL TO / DATES ──
    status_color = colors.HexColor("#2ecc71") if inv["status"]=="Paid" else colors.HexColor("#e74c3c")
    bill_data = [
        [Paragraph("<b>BILL TO</b>", ParagraphStyle("bt", fontSize=9, textColor=accent_color, fontName="Helvetica-Bold")),
         "", "",
         Paragraph(f"<b>STATUS: {inv['status'].upper()}</b>",
                   ParagraphStyle("st", fontSize=11, textColor=status_color, fontName="Helvetica-Bold", alignment=TA_RIGHT))],
        [Paragraph(f"<b>{inv['name']}</b>", bold_style), "", "",
         Paragraph(f"Issue Date: {inv['issue_date']}", right_style)],
        [Paragraph(inv.get('email',''), normal_style), "", "",
         Paragraph(f"Due Date: {inv['due_date']}", right_style)],
        [Paragraph(inv.get('address',''), normal_style), "", "", ""],
        [Paragraph(f"{inv.get('city','')} {inv.get('country','')}", normal_style), "", "", ""],
    ]
    bill_table = Table(bill_data, colWidths=[80*mm, 30*mm, 20*mm, 80*mm])
    bill_table.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("TOPPADDING",(0,0),(-1,-1),2)]))
    story.append(bill_table)
    story.append(Spacer(1, 6*mm))

    # ── ITEMS TABLE ──
    item_header = ["#", "Description", "Qty", "Unit Price", "Total"]
    item_rows = [item_header]
    for i, it in enumerate(items, 1):
        _, _, desc, qty, price, total = it
        cur = company.get("currency","USD")
        item_rows.append([str(i), desc, f"{qty:g}", f"{cur} {price:,.2f}", f"{cur} {total:,.2f}"])

    col_widths = [10*mm, 90*mm, 20*mm, 30*mm, 30*mm]
    item_table = Table(item_rows, colWidths=col_widths, repeatRows=1)
    item_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), accent_color),
        ("TEXTCOLOR",  (0,0), (-1,0), white),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,0), 10),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
        ("ALIGN",      (1,0), (1,-1), "LEFT"),
        ("ALIGN",      (3,0), (-1,-1), "RIGHT"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [white, light_gray]),
        ("GRID",       (0,0), (-1,-1), 0.3, mid_gray),
        ("TOPPADDING", (0,0), (-1,-1), 7),
        ("BOTTOMPADDING",(0,0),(-1,-1),7),
        ("FONTSIZE",   (0,1),(-1,-1),9),
    ]))
    story.append(item_table)
    story.append(Spacer(1, 6*mm))

    # ── TOTALS ──
    cur = company.get("currency","USD")
    totals_data = [
        ["", Paragraph("Subtotal:", right_style), Paragraph(f"{cur} {inv['subtotal']:,.2f}", right_style)],
        ["", Paragraph(f"Tax ({inv['tax_rate']:.1f}%):", right_style), Paragraph(f"{cur} {inv['tax_amount']:,.2f}", right_style)],
        ["", Paragraph(f"Discount:", right_style), Paragraph(f"- {cur} {inv['discount']:,.2f}", right_style)],
        ["", Paragraph(f"<b>TOTAL DUE</b>", ParagraphStyle("tt",fontSize=13,fontName="Helvetica-Bold",alignment=TA_RIGHT)),
         Paragraph(f"<b>{cur} {inv['total']:,.2f}</b>", ParagraphStyle("tv",fontSize=13,fontName="Helvetica-Bold",textColor=accent_color,alignment=TA_RIGHT))],
    ]
    totals_table = Table(totals_data, colWidths=[100*mm, 50*mm, 30*mm])
    totals_table.setStyle(TableStyle([
        ("LINEABOVE", (1,3),(2,3), 1.5, accent_color),
        ("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("BACKGROUND",(1,3),(2,3), light_gray),
    ]))
    story.append(totals_table)
    story.append(Spacer(1, 8*mm))

    # ── NOTES / TERMS ──
    if inv.get("notes"):
        story.append(HRFlowable(width="100%", thickness=0.5, color=mid_gray, spaceAfter=4))
        story.append(Paragraph("<b>Notes:</b>", bold_style))
        story.append(Paragraph(inv["notes"], normal_style))
        story.append(Spacer(1, 3*mm))
    if inv.get("terms"):
        story.append(Paragraph("<b>Terms & Conditions:</b>", bold_style))
        story.append(Paragraph(inv["terms"], normal_style))

    # ── BANK INFO ──
    if company.get("bank_name") or company.get("account_no"):
        story.append(Spacer(1, 6*mm))
        story.append(HRFlowable(width="100%", thickness=1, color=accent_color, spaceAfter=4))
        story.append(Paragraph("<b>Payment Information</b>", ParagraphStyle("pi",fontSize=10,fontName="Helvetica-Bold",textColor=accent_color)))
        story.append(Paragraph(f"Bank: {company.get('bank_name','')}  |  Account: {company.get('account_no','')}", normal_style))

    story.append(Spacer(1, 8*mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=mid_gray))
    story.append(Paragraph(f"<font color='#aaaacc'>Thank you for your business! Generated on {datetime.datetime.now().strftime('%Y-%m-%d')}</font>",
                            ParagraphStyle("footer", fontSize=8, alignment=TA_CENTER)))

    doc.build(story)

# ─── EXCEL EXPORT ─────────────────────────────────────────────────────────────
def export_unpaid_excel(save_path):
    conn = get_conn()
    c = conn.cursor()
    c.execute("""SELECT i.invoice_no, cu.name, cu.email, cu.phone,
                        i.issue_date, i.due_date, i.total, i.status
                 FROM invoices i JOIN customers cu ON i.customer_id=cu.id
                 WHERE i.status='Unpaid' ORDER BY i.due_date""")
    rows = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Unpaid Invoices"

    accent_fill = PatternFill("solid", fgColor="6c63ff")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style='thin', color="CCCCDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["Invoice No", "Customer", "Email", "Phone", "Issue Date", "Due Date", "Amount Due", "Status"]
    ws.append(headers)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col)
        cell.fill = accent_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    ws.row_dimensions[1].height = 25

    red_fill   = PatternFill("solid", fgColor="FFCCCC")
    alt_fill   = PatternFill("solid", fgColor="F0F0F8")

    for i, row in enumerate(rows, 2):
        ws.append(list(row))
        fill = red_fill if i % 2 == 0 else alt_fill
        for col in range(1, 9):
            cell = ws.cell(i, col)
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

    # Column widths
    for col, width in zip("ABCDEFGH", [15,25,28,15,14,14,14,12]):
        ws.column_dimensions[col].width = width

    # Total row
    ws.append([])
    total_row = ws.max_row + 1
    ws.cell(total_row, 6, "TOTAL DUE:")
    ws.cell(total_row, 6).font = Font(bold=True)
    ws.cell(total_row, 7, f"=SUM(G2:G{ws.max_row-1})")
    ws.cell(total_row, 7).font = Font(bold=True, color="E74C3C")

    wb.save(save_path)

# ─── MAIN APP ─────────────────────────────────────────────────────────────────
class InvoiceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        init_db()
        self.title("💼 ProInvoice — Professional Invoice Manager")
        self.geometry("1280x780")
        self.minsize(1100, 700)
        self.configure(bg=APP_BG)
        self.logo_img = None
        self._build_ui()
        self.show_frame("Dashboard")

    def _build_ui(self):
        # ── SIDEBAR ──
        self.sidebar = tk.Frame(self, bg="#13151f", width=220)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Logo area
        logo_frame = tk.Frame(self.sidebar, bg="#13151f", pady=20)
        logo_frame.pack(fill="x")
        tk.Label(logo_frame, text="💼", font=("Segoe UI", 32), bg="#13151f", fg=ACCENT).pack()
        tk.Label(logo_frame, text="ProInvoice", font=("Segoe UI", 16, "bold"),
                 bg="#13151f", fg=TEXT).pack()
        tk.Label(logo_frame, text="Invoice Manager", font=("Segoe UI", 9),
                 bg="#13151f", fg=SUBTEXT).pack()

        tk.Frame(self.sidebar, bg=BORDER, height=1).pack(fill="x", padx=15, pady=5)

        # Nav items
        nav_items = [
            ("🏠  Dashboard",    "Dashboard"),
            ("➕  New Invoice",  "NewInvoice"),
            ("📋  All Invoices", "Invoices"),
            ("👥  Customers",    "Customers"),
            ("📊  Analytics",    "Analytics"),
            ("🏢  Company",      "Company"),
        ]
        self.nav_btns = {}
        for label, name in nav_items:
            btn = tk.Button(self.sidebar, text=label, font=("Segoe UI", 11),
                            bg="#13151f", fg=SUBTEXT, relief="flat",
                            activebackground=ACCENT, activeforeground="white",
                            anchor="w", padx=20, pady=10, cursor="hand2",
                            command=lambda n=name: self.show_frame(n))
            btn.pack(fill="x")
            self.nav_btns[name] = btn

        # ── MAIN AREA ──
        self.main = tk.Frame(self, bg=APP_BG)
        self.main.pack(side="left", fill="both", expand=True)

        self.frames = {}
        for FClass, name in [
            (DashboardFrame, "Dashboard"),
            (NewInvoiceFrame, "NewInvoice"),
            (InvoicesFrame, "Invoices"),
            (CustomersFrame, "Customers"),
            (AnalyticsFrame, "Analytics"),
            (CompanyFrame, "Company"),
        ]:
            f = FClass(self.main, self)
            f.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.frames[name] = f

    def show_frame(self, name):
        for n, btn in self.nav_btns.items():
            btn.config(bg="#13151f", fg=SUBTEXT)
        if name in self.nav_btns:
            self.nav_btns[name].config(bg=ACCENT, fg="white")
        frame = self.frames[name]
        frame.tkraise()
        if hasattr(frame, "refresh"):
            frame.refresh()

# ─── DASHBOARD ────────────────────────────────────────────────────────────────
class DashboardFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=APP_BG)
        self.app = app
        self._build()

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg=APP_BG, pady=20, padx=30)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Dashboard", font=("Segoe UI", 22, "bold"),
                 bg=APP_BG, fg=TEXT).pack(side="left")
        tk.Label(hdr, text=datetime.datetime.now().strftime("%A, %B %d %Y"),
                 font=("Segoe UI", 11), bg=APP_BG, fg=SUBTEXT).pack(side="right", padx=10)

        # Stats row
        self.stats_row = tk.Frame(self, bg=APP_BG, padx=20)
        self.stats_row.pack(fill="x")

        # Recent + quick actions
        content = tk.Frame(self, bg=APP_BG, padx=20, pady=15)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=2)
        content.columnconfigure(1, weight=1)

        self.recent_frame = card(content)
        self.recent_frame.grid(row=0, column=0, sticky="nsew", padx=(0,10))

        quick_frame = card(content)
        quick_frame.grid(row=0, column=1, sticky="nsew")
        tk.Label(quick_frame, text="Quick Actions", font=("Segoe UI", 13, "bold"),
                 bg=CARD_BG, fg=TEXT, pady=15).pack()
        styled_btn(quick_frame, "➕  New Invoice", lambda: self.app.show_frame("NewInvoice")).pack(pady=5)
        styled_btn(quick_frame, "👥  Add Customer", lambda: self.app.show_frame("Customers"), bg="#2ecc71").pack(pady=5)
        styled_btn(quick_frame, "📊  Analytics", lambda: self.app.show_frame("Analytics"), bg=ACCENT2).pack(pady=5)
        styled_btn(quick_frame, "📋  All Invoices", lambda: self.app.show_frame("Invoices"), bg="#3498db").pack(pady=5)

    def refresh(self):
        for w in self.stats_row.winfo_children():
            w.destroy()

        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM invoices")
        total_inv = c.fetchone()[0]
        c.execute("SELECT COALESCE(SUM(total),0) FROM invoices WHERE status='Unpaid'")
        unpaid_amt = c.fetchone()[0]
        c.execute("SELECT COALESCE(SUM(total),0) FROM invoices WHERE status='Paid'")
        paid_amt = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM customers")
        cust_count = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM invoices WHERE status='Unpaid'")
        unpaid_count = c.fetchone()[0]
        conn.close()

        stats = [
            ("📄", "Total Invoices", str(total_inv), ACCENT),
            ("💰", "Total Earned",   f"${paid_amt:,.0f}", SUCCESS),
            ("⚠️", "Unpaid",          f"${unpaid_amt:,.0f}", DANGER),
            ("👥", "Customers",       str(cust_count), "#3498db"),
        ]
        for icon, label, val, color in stats:
            f = tk.Frame(self.stats_row, bg=CARD_BG, padx=20, pady=15, relief="flat")
            f.pack(side="left", expand=True, fill="x", padx=8, pady=8)
            tk.Label(f, text=icon, font=("Segoe UI", 24), bg=CARD_BG).pack()
            tk.Label(f, text=val, font=("Segoe UI", 20, "bold"), bg=CARD_BG, fg=color).pack()
            tk.Label(f, text=label, font=("Segoe UI", 9), bg=CARD_BG, fg=SUBTEXT).pack()

        # Recent invoices
        for w in self.recent_frame.winfo_children():
            w.destroy()
        tk.Label(self.recent_frame, text="Recent Invoices", font=("Segoe UI", 13, "bold"),
                 bg=CARD_BG, fg=TEXT, pady=10, padx=15, anchor="w").pack(fill="x")

        conn = get_conn()
        c = conn.cursor()
        c.execute("""SELECT i.invoice_no, cu.name, i.total, i.status, i.due_date
                     FROM invoices i JOIN customers cu ON i.customer_id=cu.id
                     ORDER BY i.id DESC LIMIT 8""")
        rows = c.fetchall()
        conn.close()

        hdr_f = tk.Frame(self.recent_frame, bg="#22253a")
        hdr_f.pack(fill="x", padx=10, pady=(0,2))
        for col, w in [("Invoice#",14),("Customer",20),("Amount",12),("Status",10),("Due Date",14)]:
            tk.Label(hdr_f, text=col, font=("Segoe UI", 9, "bold"),
                     bg="#22253a", fg=SUBTEXT, width=w, anchor="w").pack(side="left", padx=3, pady=4)

        for i, (inv_no, name, total, status, due) in enumerate(rows):
            row_f = tk.Frame(self.recent_frame, bg=CARD_BG if i%2==0 else "#1e2133")
            row_f.pack(fill="x", padx=10)
            s_color = SUCCESS if status=="Paid" else DANGER
            for val, w, color in [
                (inv_no, 14, ACCENT), (name, 20, TEXT), (f"${total:,.2f}", 12, TEXT),
                (status, 10, s_color), (due, 14, SUBTEXT)
            ]:
                tk.Label(row_f, text=val, font=("Segoe UI", 9), bg=row_f["bg"],
                         fg=color, width=w, anchor="w").pack(side="left", padx=3, pady=3)

# ─── NEW INVOICE ─────────────────────────────────────────────────────────────
class NewInvoiceFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=APP_BG)
        self.app = app
        self.items = []
        self._build()

    def _build(self):
        # Scrollable canvas
        canvas = tk.Canvas(self, bg=APP_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scroll_frame = tk.Frame(canvas, bg=APP_BG)
        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        sf = self.scroll_frame
        tk.Label(sf, text="Create New Invoice", font=("Segoe UI", 20, "bold"),
                 bg=APP_BG, fg=TEXT, pady=20, padx=30).pack(fill="x")

        body = tk.Frame(sf, bg=APP_BG, padx=20, pady=5)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        # ── Customer selection ──
        cust_card = card(body, padx=20, pady=15)
        cust_card.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        lbl(cust_card, "Customer", 13, bold=True).pack(anchor="w", pady=(0,10))

        row_f = tk.Frame(cust_card, bg=CARD_BG)
        row_f.pack(fill="x")
        lbl(row_f, "Select Customer:").pack(side="left")
        self.cust_var = tk.StringVar()
        self.cust_combo = ttk.Combobox(row_f, textvariable=self.cust_var, width=28,
                                        font=("Segoe UI", 10))
        self.cust_combo.pack(side="left", padx=8)
        styled_btn(row_f, "+ New", self._new_customer, width=8).pack(side="left")

        # ── Invoice Info ──
        info_card = card(body, padx=20, pady=15)
        info_card.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        lbl(info_card, "Invoice Details", 13, bold=True).pack(anchor="w", pady=(0,10))

        self.inv_no_var = tk.StringVar(value=next_invoice_number())
        self.issue_var  = tk.StringVar(value=datetime.date.today().strftime("%Y-%m-%d"))
        self.due_var    = tk.StringVar(value=(datetime.date.today()+datetime.timedelta(30)).strftime("%Y-%m-%d"))

        for lbl_text, var in [("Invoice No:", self.inv_no_var),
                               ("Issue Date:", self.issue_var),
                               ("Due Date:",   self.due_var)]:
            r = tk.Frame(info_card, bg=CARD_BG)
            r.pack(fill="x", pady=3)
            tk.Label(r, text=lbl_text, font=("Segoe UI", 10), bg=CARD_BG, fg=SUBTEXT, width=12, anchor="w").pack(side="left")
            entry(r, width=22, textvariable=var).pack(side="left", padx=5)

        # ── Items ──
        items_card = card(body, padx=20, pady=15)
        items_card.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        lbl(items_card, "Line Items", 13, bold=True).pack(anchor="w", pady=(0,10))

        # Items header
        hdr_f = tk.Frame(items_card, bg="#22253a")
        hdr_f.pack(fill="x")
        for col, w in [("Description",45),("Qty",8),("Unit Price",14),("Total",14),("",5)]:
            tk.Label(hdr_f, text=col, font=("Segoe UI", 9,"bold"),
                     bg="#22253a", fg=SUBTEXT, width=w, anchor="w").pack(side="left", padx=5, pady=5)

        self.items_frame = tk.Frame(items_card, bg=CARD_BG)
        self.items_frame.pack(fill="x")
        self.item_rows = []

        add_r = tk.Frame(items_card, bg=CARD_BG, pady=8)
        add_r.pack(fill="x")
        styled_btn(add_r, "+ Add Item", self._add_item_row, width=14, bg="#2ecc71").pack(side="left")

        # Totals
        totals_card = card(body, padx=20, pady=15)
        totals_card.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)

        self.tax_var      = tk.DoubleVar(value=10.0)
        self.discount_var = tk.DoubleVar(value=0.0)

        for lbl_text, var in [("Tax Rate (%):", self.tax_var),
                               ("Discount ($):", self.discount_var)]:
            r = tk.Frame(totals_card, bg=CARD_BG)
            r.pack(fill="x", pady=4)
            tk.Label(r, text=lbl_text, bg=CARD_BG, fg=SUBTEXT, font=("Segoe UI",10), width=14, anchor="w").pack(side="left")
            entry(r, width=14, textvariable=var).pack(side="left")

        self.total_lbl = tk.Label(totals_card, text="Total: $0.00",
                                   font=("Segoe UI", 16, "bold"), bg=CARD_BG, fg=ACCENT, pady=10)
        self.total_lbl.pack()

        # Notes/Terms
        notes_card = card(body, padx=20, pady=15)
        notes_card.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        lbl(notes_card, "Notes", 11, bold=True).pack(anchor="w")
        self.notes_txt = tk.Text(notes_card, height=4, font=("Segoe UI",9),
                                  bg="#22253a", fg=TEXT, relief="flat",
                                  highlightthickness=1, highlightbackground=BORDER,
                                  insertbackground=TEXT)
        self.notes_txt.pack(fill="x", pady=5)
        lbl(notes_card, "Terms & Conditions", 11, bold=True).pack(anchor="w")
        self.terms_txt = tk.Text(notes_card, height=3, font=("Segoe UI",9),
                                  bg="#22253a", fg=TEXT, relief="flat",
                                  highlightthickness=1, highlightbackground=BORDER,
                                  insertbackground=TEXT)
        self.terms_txt.pack(fill="x", pady=5)
        self.terms_txt.insert("end", "Payment due within 30 days. Late fees may apply.")

        # Action buttons
        action_row = tk.Frame(sf, bg=APP_BG, pady=15, padx=20)
        action_row.pack(fill="x")
        styled_btn(action_row, "💾 Save Invoice", self._save_invoice, width=20).pack(side="left", padx=5)
        styled_btn(action_row, "🖨️  Save as PDF", self._save_pdf, bg=ACCENT2, width=18).pack(side="left", padx=5)
        styled_btn(action_row, "🔄 Reset", self._reset, bg="#555", width=12).pack(side="left", padx=5)

        self._add_item_row()
        self._refresh_customers()

    def _refresh_customers(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT id, name FROM customers ORDER BY name")
        self._customers = {f"{r[1]} (ID:{r[0]})": r[0] for r in c.fetchall()}
        conn.close()
        self.cust_combo["values"] = list(self._customers.keys())

    def _new_customer(self):
        self.app.show_frame("Customers")

    def _add_item_row(self, desc="", qty=1, price=0.0):
        f = tk.Frame(self.items_frame, bg=CARD_BG)
        f.pack(fill="x", pady=2)

        desc_e  = entry(f, width=44); desc_e.pack(side="left", padx=3); desc_e.insert(0, desc)
        qty_e   = entry(f, width=7);  qty_e.pack(side="left", padx=3);  qty_e.insert(0, str(qty))
        price_e = entry(f, width=13); price_e.pack(side="left", padx=3); price_e.insert(0, str(price))
        total_l = tk.Label(f, text="$0.00", font=("Segoe UI",10), bg=CARD_BG, fg=SUCCESS, width=14, anchor="w")
        total_l.pack(side="left", padx=3)

        def update_total(*_):
            try:
                t = float(qty_e.get()) * float(price_e.get())
                total_l.config(text=f"${t:,.2f}")
            except:
                total_l.config(text="$0.00")
            self._recalc_total()

        qty_e.bind("<KeyRelease>", update_total)
        price_e.bind("<KeyRelease>", update_total)

        del_btn = tk.Button(f, text="✕", bg=DANGER, fg="white", font=("Segoe UI",9,"bold"),
                            relief="flat", cursor="hand2", padx=6, pady=2,
                            command=lambda: [f.destroy(), self._recalc_total()])
        del_btn.pack(side="left", padx=3)

        self.item_rows.append((f, desc_e, qty_e, price_e, total_l))

    def _recalc_total(self):
        subtotal = 0
        for _, _, qty_e, price_e, _ in self.item_rows:
            if qty_e.winfo_exists():
                try: subtotal += float(qty_e.get()) * float(price_e.get())
                except: pass
        try: tax = subtotal * self.tax_var.get() / 100
        except: tax = 0
        try: disc = self.discount_var.get()
        except: disc = 0
        total = subtotal + tax - disc
        self.total_lbl.config(text=f"Total: ${total:,.2f}")

    def _get_items(self):
        result = []
        for _, desc_e, qty_e, price_e, _ in self.item_rows:
            if not desc_e.winfo_exists(): continue
            try:
                d = desc_e.get().strip()
                q = float(qty_e.get())
                p = float(price_e.get())
                if d:
                    result.append((d, q, p, q*p))
            except: pass
        return result

    def _save_invoice(self):
        cust_key = self.cust_var.get()
        if not cust_key or cust_key not in self._customers:
            messagebox.showerror("Error", "Please select a customer first.")
            return
        items = self._get_items()
        if not items:
            messagebox.showerror("Error", "Add at least one line item.")
            return

        subtotal = sum(i[3] for i in items)
        tax = subtotal * self.tax_var.get() / 100
        disc = self.discount_var.get()
        total = subtotal + tax - disc

        conn = get_conn()
        c = conn.cursor()
        try:
            c.execute("""INSERT INTO invoices
                (invoice_no, customer_id, issue_date, due_date,
                 subtotal, tax_rate, tax_amount, discount, total, notes, terms, created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                (self.inv_no_var.get(), self._customers[cust_key],
                 self.issue_var.get(), self.due_var.get(),
                 subtotal, self.tax_var.get(), tax, disc, total,
                 self.notes_txt.get("1.0","end").strip(),
                 self.terms_txt.get("1.0","end").strip(),
                 datetime.datetime.now().isoformat()))
            inv_id = c.lastrowid
            for desc, qty, price, tot in items:
                c.execute("INSERT INTO invoice_items (invoice_id,description,quantity,unit_price,total) VALUES (?,?,?,?,?)",
                          (inv_id, desc, qty, price, tot))
            conn.commit()
            messagebox.showinfo("Saved", f"Invoice {self.inv_no_var.get()} saved!")
            self._reset()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Invoice number already exists.")
        finally:
            conn.close()

    def _save_pdf(self):
        # Save first if not already saved
        cust_key = self.cust_var.get()
        if not cust_key or cust_key not in self._customers:
            messagebox.showerror("Error", "Select a customer first.")
            return
        self._save_invoice()
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT id FROM invoices WHERE invoice_no=?", (self.inv_no_var.get(),))
        row = c.fetchone()
        conn.close()
        if not row:
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF","*.pdf")],
                                             initialfile=f"invoice_{self.inv_no_var.get()}.pdf")
        if path:
            generate_pdf(row[0], path)
            messagebox.showinfo("PDF Saved", f"Invoice PDF saved to:\n{path}")

    def _reset(self):
        self.inv_no_var.set(next_invoice_number())
        self.cust_var.set("")
        for f, *_ in self.item_rows:
            if f.winfo_exists():
                f.destroy()
        self.item_rows.clear()
        self.notes_txt.delete("1.0", "end")
        self.terms_txt.delete("1.0", "end")
        self.terms_txt.insert("end", "Payment due within 30 days.")
        self._add_item_row()
        self._refresh_customers()

    def refresh(self):
        self._refresh_customers()

# ─── INVOICES LIST ────────────────────────────────────────────────────────────
class InvoicesFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=APP_BG)
        self.app = app
        self._build()

    def _build(self):
        # Header + search
        top = tk.Frame(self, bg=APP_BG, padx=20, pady=15)
        top.pack(fill="x")
        tk.Label(top, text="All Invoices", font=("Segoe UI",20,"bold"), bg=APP_BG, fg=TEXT).pack(side="left")

        btn_r = tk.Frame(top, bg=APP_BG)
        btn_r.pack(side="right")
        styled_btn(btn_r, "📥 Export Unpaid Excel", self._export_excel, bg=SUCCESS, width=22).pack(side="left", padx=4)
        styled_btn(btn_r, "🔄 Refresh", self.refresh, bg="#3498db", width=12).pack(side="left", padx=4)

        # Filter bar
        flt = tk.Frame(self, bg=APP_BG, padx=20)
        flt.pack(fill="x")
        self.filter_var = tk.StringVar(value="All")
        for opt in ["All","Paid","Unpaid"]:
            rb = tk.Radiobutton(flt, text=opt, variable=self.filter_var, value=opt,
                                bg=APP_BG, fg=TEXT, selectcolor=ACCENT,
                                font=("Segoe UI",10), command=self.refresh)
            rb.pack(side="left", padx=8)

        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *_: self.refresh())
        tk.Label(flt, text="Search:", bg=APP_BG, fg=SUBTEXT, font=("Segoe UI",10)).pack(side="left", padx=(20,5))
        entry(flt, width=25, textvariable=self.search_var).pack(side="left")

        # Table
        cols = ("Invoice#", "Customer", "Issue Date", "Due Date", "Total", "Status")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Treeview",
                         background=CARD_BG, fieldbackground=CARD_BG,
                         foreground=TEXT, rowheight=32,
                         font=("Segoe UI",10))
        style.configure("Custom.Treeview.Heading",
                         background="#22253a", foreground=SUBTEXT,
                         font=("Segoe UI",10,"bold"), relief="flat")
        style.map("Custom.Treeview", background=[("selected", ACCENT)])

        frame = tk.Frame(self, bg=APP_BG, padx=20)
        frame.pack(fill="both", expand=True, pady=10)

        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                  style="Custom.Treeview", selectmode="browse")
        for col, w in zip(cols, [140,180,120,120,120,100]):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")
        self.tree.tag_configure("paid",   background="#1a2e22", foreground=SUCCESS)
        self.tree.tag_configure("unpaid", background="#2e1a1a", foreground=DANGER)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Context menu
        self.menu = tk.Menu(self, tearoff=0, bg=CARD_BG, fg=TEXT,
                            activebackground=ACCENT, font=("Segoe UI",10))
        self.menu.add_command(label="📄 Export PDF",   command=self._export_pdf)
        self.menu.add_command(label="✅ Mark as Paid", command=self._mark_paid)
        self.menu.add_command(label="🗑️ Delete",        command=self._delete)
        self.tree.bind("<Button-3>", lambda e: self.menu.post(e.x_root, e.y_root))
        self.tree.bind("<Double-1>", lambda e: self._export_pdf())

        # Bottom hint
        tk.Label(self, text="Right-click or double-click an invoice for actions",
                 font=("Segoe UI",9), bg=APP_BG, fg=SUBTEXT, pady=6).pack()

    def refresh(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        flt = self.filter_var.get()
        srch = self.search_var.get().lower()
        conn = get_conn()
        c = conn.cursor()
        q = """SELECT i.id, i.invoice_no, cu.name, i.issue_date, i.due_date, i.total, i.status
               FROM invoices i JOIN customers cu ON i.customer_id=cu.id"""
        params = []
        if flt != "All":
            q += " WHERE i.status=?"
            params.append(flt)
        q += " ORDER BY i.id DESC"
        c.execute(q, params)
        for row in c.fetchall():
            iid, inv_no, name, iss, due, tot, status = row
            if srch and srch not in inv_no.lower() and srch not in name.lower():
                continue
            tag = "paid" if status=="Paid" else "unpaid"
            self.tree.insert("", "end", iid=str(iid),
                             values=(inv_no, name, iss, due, f"${tot:,.2f}", status),
                             tags=(tag,))
        conn.close()

    def _selected_id(self):
        sel = self.tree.focus()
        if not sel:
            messagebox.showwarning("Select", "Please select an invoice first.")
            return None
        return int(sel)

    def _export_pdf(self):
        inv_id = self._selected_id()
        if not inv_id: return
        vals = self.tree.item(str(inv_id), "values")
        path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF","*.pdf")],
                                             initialfile=f"invoice_{vals[0]}.pdf")
        if path:
            generate_pdf(inv_id, path)
            messagebox.showinfo("Done", f"PDF exported to:\n{path}")

    def _mark_paid(self):
        inv_id = self._selected_id()
        if not inv_id: return
        if messagebox.askyesno("Confirm", "Mark this invoice as Paid and delete it?"):
            conn = get_conn()
            conn.execute("DELETE FROM invoice_items WHERE invoice_id=?", (inv_id,))
            conn.execute("DELETE FROM invoices WHERE id=?", (inv_id,))
            conn.commit()
            conn.close()
            self.refresh()
            messagebox.showinfo("Done", "Invoice marked as paid and removed.")

    def _delete(self):
        inv_id = self._selected_id()
        if not inv_id: return
        if messagebox.askyesno("Delete", "Permanently delete this invoice?"):
            conn = get_conn()
            conn.execute("DELETE FROM invoice_items WHERE invoice_id=?", (inv_id,))
            conn.execute("DELETE FROM invoices WHERE id=?", (inv_id,))
            conn.commit()
            conn.close()
            self.refresh()

    def _export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel","*.xlsx")],
                                             initialfile="unpaid_invoices.xlsx")
        if path:
            export_unpaid_excel(path)
            messagebox.showinfo("Exported", f"Unpaid invoices exported to:\n{path}")

# ─── CUSTOMERS ────────────────────────────────────────────────────────────────
class CustomersFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=APP_BG)
        self.app = app
        self._build()

    def _build(self):
        top = tk.Frame(self, bg=APP_BG, padx=20, pady=15)
        top.pack(fill="x")
        tk.Label(top, text="Customer Management", font=("Segoe UI",20,"bold"), bg=APP_BG, fg=TEXT).pack(side="left")
        styled_btn(top, "+ Add Customer", self._add_dialog, bg=SUCCESS, width=16).pack(side="right", padx=5)

        frame = tk.Frame(self, bg=APP_BG, padx=20)
        frame.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Cust.Treeview",
                         background=CARD_BG, fieldbackground=CARD_BG,
                         foreground=TEXT, rowheight=30, font=("Segoe UI",10))
        style.configure("Cust.Treeview.Heading",
                         background="#22253a", foreground=SUBTEXT,
                         font=("Segoe UI",10,"bold"))

        cols = ("ID","Name","Email","Phone","City","Country","Invoices")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", style="Cust.Treeview")
        for col, w in zip(cols, [50,180,200,120,120,100,80]):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        menu = tk.Menu(self, tearoff=0, bg=CARD_BG, fg=TEXT, activebackground=ACCENT, font=("Segoe UI",10))
        menu.add_command(label="✏️ Edit",   command=self._edit_dialog)
        menu.add_command(label="🗑️ Delete", command=self._delete)
        self.tree.bind("<Button-3>", lambda e: menu.post(e.x_root, e.y_root))

    def refresh(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        conn = get_conn()
        c = conn.cursor()
        c.execute("""SELECT cu.id, cu.name, cu.email, cu.phone, cu.city, cu.country,
                            COUNT(i.id) as inv_count
                     FROM customers cu LEFT JOIN invoices i ON i.customer_id=cu.id
                     GROUP BY cu.id ORDER BY cu.name""")
        for row in c.fetchall():
            self.tree.insert("", "end", iid=str(row[0]), values=row)
        conn.close()

    def _customer_dialog(self, title, defaults=None):
        d = defaults or {}
        win = tk.Toplevel(self, bg=CARD_BG)
        win.title(title)
        win.geometry("420x460")
        win.grab_set()

        fields = {}
        for i, (lbl_text, key) in enumerate([("Full Name*","name"),("Email","email"),
                                              ("Phone","phone"),("Address","address"),
                                              ("City","city"),("Country","country"),("Notes","notes")]):
            tk.Label(win, text=lbl_text, font=("Segoe UI",10), bg=CARD_BG, fg=SUBTEXT).pack(anchor="w", padx=20, pady=(8,0))
            if key == "notes":
                e = tk.Text(win, height=3, font=("Segoe UI",10), bg="#22253a", fg=TEXT,
                            relief="flat", insertbackground=TEXT,
                            highlightthickness=1, highlightbackground=BORDER)
                e.pack(fill="x", padx=20)
                if d.get(key): e.insert("end", d[key])
            else:
                e = entry(win, width=40)
                e.pack(padx=20)
                if d.get(key): e.insert(0, d[key])
            fields[key] = e

        result = {}
        def submit():
            n = fields["name"].get().strip()
            if not n:
                messagebox.showerror("Error", "Name is required.", parent=win)
                return
            for k in ["name","email","phone","address","city","country"]:
                result[k] = fields[k].get().strip()
            result["notes"] = fields["notes"].get("1.0","end").strip()
            win.destroy()

        styled_btn(win, "Save Customer", submit).pack(pady=15)
        win.wait_window()
        return result if result else None

    def _add_dialog(self):
        data = self._customer_dialog("Add Customer")
        if data:
            conn = get_conn()
            conn.execute("""INSERT INTO customers (name,email,phone,address,city,country,notes)
                           VALUES (?,?,?,?,?,?,?)""",
                         (data["name"],data["email"],data["phone"],
                          data["address"],data["city"],data["country"],data["notes"]))
            conn.commit()
            conn.close()
            self.refresh()

    def _edit_dialog(self):
        sel = self.tree.focus()
        if not sel: return
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT * FROM customers WHERE id=?", (int(sel),))
        row = c.fetchone()
        conn.close()
        if not row: return
        _, name, email, phone, address, city, country, notes = row
        data = self._customer_dialog("Edit Customer",
                                     {"name":name,"email":email,"phone":phone,
                                      "address":address,"city":city,"country":country,"notes":notes or ""})
        if data:
            conn = get_conn()
            conn.execute("""UPDATE customers SET name=?,email=?,phone=?,address=?,city=?,country=?,notes=?
                           WHERE id=?""",
                         (data["name"],data["email"],data["phone"],
                          data["address"],data["city"],data["country"],data["notes"],int(sel)))
            conn.commit()
            conn.close()
            self.refresh()

    def _delete(self):
        sel = self.tree.focus()
        if not sel: return
        if messagebox.askyesno("Delete", "Delete this customer and ALL their invoices?"):
            conn = get_conn()
            c = conn.cursor()
            c.execute("SELECT id FROM invoices WHERE customer_id=?", (int(sel),))
            for (iid,) in c.fetchall():
                conn.execute("DELETE FROM invoice_items WHERE invoice_id=?", (iid,))
            conn.execute("DELETE FROM invoices WHERE customer_id=?", (int(sel),))
            conn.execute("DELETE FROM customers WHERE id=?", (int(sel),))
            conn.commit()
            conn.close()
            self.refresh()

# ─── ANALYTICS ────────────────────────────────────────────────────────────────
class AnalyticsFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=APP_BG)
        self.app = app
        self._build()

    def _build(self):
        top = tk.Frame(self, bg=APP_BG, padx=20, pady=15)
        top.pack(fill="x")
        tk.Label(top, text="Analytics & Reports", font=("Segoe UI",20,"bold"), bg=APP_BG, fg=TEXT).pack(side="left")
        styled_btn(top, "🔄 Refresh", self.refresh, bg="#3498db", width=12).pack(side="right")

        self.chart_frame = tk.Frame(self, bg=APP_BG)
        self.chart_frame.pack(fill="both", expand=True, padx=20, pady=10)

    def refresh(self):
        for w in self.chart_frame.winfo_children():
            w.destroy()

        conn = get_conn()
        c = conn.cursor()

        # Monthly revenue
        c.execute("""SELECT strftime('%Y-%m', issue_date) as month, SUM(total)
                     FROM invoices WHERE status='Paid'
                     GROUP BY month ORDER BY month DESC LIMIT 12""")
        monthly = c.fetchall()

        # Revenue by customer
        c.execute("""SELECT cu.name, SUM(i.total) as rev
                     FROM invoices i JOIN customers cu ON i.customer_id=cu.id
                     WHERE i.status='Paid'
                     GROUP BY cu.id ORDER BY rev DESC LIMIT 8""")
        by_customer = c.fetchall()

        # Paid vs Unpaid
        c.execute("SELECT status, COUNT(*), SUM(total) FROM invoices GROUP BY status")
        status_data = {r[0]: (r[1], r[2]) for r in c.fetchall()}
        conn.close()

        plt.style.use("dark_background")
        fig = plt.figure(figsize=(14, 9), facecolor=APP_BG)
        fig.patch.set_facecolor("#0f1117")

        # ── Chart 1: Monthly line chart ──
        ax1 = fig.add_subplot(221)
        ax1.set_facecolor(CARD_BG)
        if monthly:
            months = [m[0] for m in reversed(monthly)]
            revenues = [m[1] for m in reversed(monthly)]
            ax1.plot(months, revenues, color=ACCENT, linewidth=2.5, marker="o", markersize=6)
            ax1.fill_between(range(len(months)), revenues, alpha=0.2, color=ACCENT)
            ax1.set_xticks(range(len(months)))
            ax1.set_xticklabels(months, rotation=45, ha="right", fontsize=8)
        ax1.set_title("Monthly Revenue (Paid)", color=TEXT, fontsize=11, pad=10)
        ax1.tick_params(colors=SUBTEXT)
        ax1.spines[:].set_color(BORDER)
        ax1.yaxis.label.set_color(SUBTEXT)

        # ── Chart 2: Bar chart by customer ──
        ax2 = fig.add_subplot(222)
        ax2.set_facecolor(CARD_BG)
        if by_customer:
            names = [r[0][:12] for r in by_customer]
            vals  = [r[1] for r in by_customer]
            bars  = ax2.bar(names, vals, color=[ACCENT, ACCENT2, SUCCESS, DANGER,
                                                 "#3498db","#9b59b6","#1abc9c","#e67e22"][:len(vals)])
            for bar, val in zip(bars, vals):
                ax2.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.5,
                         f"${val:,.0f}", ha="center", va="bottom", fontsize=7, color=TEXT)
            ax2.set_xticklabels(names, rotation=35, ha="right", fontsize=8)
        ax2.set_title("Revenue by Customer", color=TEXT, fontsize=11, pad=10)
        ax2.tick_params(colors=SUBTEXT)
        ax2.spines[:].set_color(BORDER)

        # ── Chart 3: Pie - paid vs unpaid ──
        ax3 = fig.add_subplot(223)
        ax3.set_facecolor(CARD_BG)
        paid_amt   = status_data.get("Paid",   (0,0))[1] or 0
        unpaid_amt = status_data.get("Unpaid", (0,0))[1] or 0
        if paid_amt + unpaid_amt > 0:
            wedges, texts, autotexts = ax3.pie(
                [paid_amt, unpaid_amt],
                labels=["Paid", "Unpaid"],
                colors=[SUCCESS, DANGER],
                autopct="%1.1f%%",
                startangle=90,
                wedgeprops=dict(edgecolor=CARD_BG, linewidth=2)
            )
            for t in texts+autotexts: t.set_color(TEXT)
        ax3.set_title("Paid vs Unpaid", color=TEXT, fontsize=11, pad=10)

        # ── Chart 4: Invoice count timeline ──
        ax4 = fig.add_subplot(224)
        ax4.set_facecolor(CARD_BG)
        conn2 = get_conn()
        c2 = conn2.cursor()
        c2.execute("""SELECT strftime('%Y-%m', issue_date) as month, COUNT(*)
                      FROM invoices GROUP BY month ORDER BY month DESC LIMIT 12""")
        inv_monthly = c2.fetchall()
        conn2.close()
        if inv_monthly:
            months2 = [m[0] for m in reversed(inv_monthly)]
            counts  = [m[1] for m in reversed(inv_monthly)]
            ax4.bar(months2, counts, color=ACCENT2, alpha=0.85)
            ax4.set_xticks(range(len(months2)))
            ax4.set_xticklabels(months2, rotation=45, ha="right", fontsize=8)
        ax4.set_title("Invoices Created per Month", color=TEXT, fontsize=11, pad=10)
        ax4.tick_params(colors=SUBTEXT)
        ax4.spines[:].set_color(BORDER)

        plt.tight_layout(pad=2.5)
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)

# ─── COMPANY SETTINGS ─────────────────────────────────────────────────────────
class CompanyFrame(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=APP_BG)
        self.app = app
        self.logo_preview = None
        self._build()

    def _build(self):
        tk.Label(self, text="Company Settings", font=("Segoe UI",20,"bold"),
                 bg=APP_BG, fg=TEXT, pady=20, padx=30).pack(fill="x")

        body = tk.Frame(self, bg=APP_BG, padx=30)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        # Left: Company fields
        left = card(body, padx=25, pady=20)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,10), pady=5)
        lbl(left, "Company Information", 14, bold=True).pack(anchor="w", pady=(0,15))

        self.fields = {}
        for label, key in [("Company Name","name"),("Address","address"),
                            ("Phone","phone"),("Email","email"),
                            ("Website","website"),("Tax ID / VAT","tax_id"),
                            ("Bank Name","bank_name"),("Account Number","account_no"),
                            ("Currency (USD/EUR/NPR...)","currency")]:
            r = tk.Frame(left, bg=CARD_BG)
            r.pack(fill="x", pady=4)
            tk.Label(r, text=label, font=("Segoe UI",10), bg=CARD_BG, fg=SUBTEXT, width=24, anchor="w").pack(side="left")
            e = entry(r, width=28)
            e.pack(side="left")
            self.fields[key] = e

        styled_btn(left, "💾 Save Settings", self._save).pack(pady=15)

        # Right: Logo
        right = card(body, padx=25, pady=20)
        right.grid(row=0, column=1, sticky="nsew", padx=(10,0), pady=5)
        lbl(right, "Company Logo", 14, bold=True).pack(anchor="w", pady=(0,15))
        lbl(right, "Upload a round company logo (any format)", 10, color=SUBTEXT).pack(anchor="w")
        lbl(right, "It will be shown as a circle on invoices", 9, color=SUBTEXT).pack(anchor="w", pady=(0,10))

        self.logo_canvas = tk.Canvas(right, width=150, height=150, bg="#22253a",
                                      highlightthickness=1, highlightbackground=BORDER)
        self.logo_canvas.pack(pady=10)
        self._draw_logo_placeholder()

        styled_btn(right, "📁 Upload Logo", self._upload_logo, bg=ACCENT2, width=18).pack(pady=5)
        styled_btn(right, "❌ Remove Logo", self._remove_logo, bg=DANGER, width=18).pack(pady=5)

        self.refresh()

    def _draw_logo_placeholder(self):
        self.logo_canvas.delete("all")
        self.logo_canvas.create_oval(10,10,140,140, fill="#2a2d3e", outline=ACCENT, width=2)
        self.logo_canvas.create_text(75,75, text="LOGO", fill=SUBTEXT, font=("Segoe UI",14,"bold"))

    def _upload_logo(self):
        path = filedialog.askopenfilename(filetypes=[("Images","*.png *.jpg *.jpeg *.gif *.bmp *.webp")])
        if not path: return
        shutil.copy(path, LOGO_PATH)
        self._show_logo(LOGO_PATH)
        conn = get_conn()
        conn.execute("UPDATE companies SET logo_path=?", (LOGO_PATH,))
        conn.commit()
        conn.close()
        messagebox.showinfo("Logo Saved", "Company logo updated!")

    def _remove_logo(self):
        if os.path.exists(LOGO_PATH):
            os.remove(LOGO_PATH)
        conn = get_conn()
        conn.execute("UPDATE companies SET logo_path=''")
        conn.commit()
        conn.close()
        self._draw_logo_placeholder()

    def _show_logo(self, path):
        try:
            round_img = make_round_logo(path, 130)
            tk_img = ImageTk.PhotoImage(round_img)
            self.logo_canvas.delete("all")
            self.logo_canvas.create_image(75, 75, image=tk_img)
            self.logo_canvas._img = tk_img  # prevent GC
        except:
            self._draw_logo_placeholder()

    def refresh(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT * FROM companies LIMIT 1")
        row = c.fetchone()
        conn.close()
        if row:
            _, name, address, phone, email, website, tax_id, logo_path, bank_name, account_no, currency = row
            for key, val in [("name",name),("address",address),("phone",phone),
                              ("email",email),("website",website),("tax_id",tax_id),
                              ("bank_name",bank_name),("account_no",account_no),("currency",currency)]:
                self.fields[key].delete(0,"end")
                if val: self.fields[key].insert(0, val)
            if logo_path and os.path.exists(logo_path):
                self._show_logo(logo_path)

    def _save(self):
        data = {k: e.get().strip() for k, e in self.fields.items()}
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT id FROM companies LIMIT 1")
        exists = c.fetchone()
        logo = LOGO_PATH if os.path.exists(LOGO_PATH) else ""
        if exists:
            conn.execute("""UPDATE companies SET name=?,address=?,phone=?,email=?,website=?,
                           tax_id=?,bank_name=?,account_no=?,currency=? WHERE id=?""",
                         (data["name"],data["address"],data["phone"],data["email"],
                          data["website"],data["tax_id"],data["bank_name"],data["account_no"],
                          data["currency"] or "USD", exists[0]))
        else:
            conn.execute("""INSERT INTO companies (name,address,phone,email,website,tax_id,logo_path,bank_name,account_no,currency)
                           VALUES (?,?,?,?,?,?,?,?,?,?)""",
                         (data["name"],data["address"],data["phone"],data["email"],
                          data["website"],data["tax_id"],logo,data["bank_name"],
                          data["account_no"],data["currency"] or "USD"))
        conn.commit()
        conn.close()
        messagebox.showinfo("Saved", "Company settings saved successfully!")

# ─── RUN ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = InvoiceApp()
    app.mainloop()