"""
FACTORY MANAGEMENT SYSTEM PRO v11
- Works on its own database (factory_pro_v11.db). On first run it COPIES
  factory_pro_v10.db, so v10.4 and its data are never modified.
- Articles catalog, edit operations, soft delete + restore, weeks history,
  safe week closing (files first, then PAID), archive filters,
  real per-year analysis, weighted average price, automatic backups,
  Excel export (needs: pip install openpyxl).
"""
import customtkinter as ctk
from tkinter import ttk, messagebox
import sqlite3
import os
import re
import sys
import shutil
from datetime import datetime, date

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

# ---------------------------------------------------------------- settings
APP_TITLE = "FACTORY MANAGEMENT SYSTEM PRO v11"
COMPANY_NAME = ""          # put the company name here to print it on reports
BACKUPS_TO_KEEP = 30

if getattr(sys, "frozen", False):          # running as .exe (PyInstaller)
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(BASE_DIR, "factory_pro_v11.db")
OLD_DB_PATH = os.path.join(BASE_DIR, "factory_pro_v10.db")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
REPORTS_DIR = os.path.join(BASE_DIR, "Reports")

GREEN, BLUE, RED, ORANGE, GREY = "#1e8449", "#1f6aa5", "#c0392b", "#b34700", "#3a3a3a"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


# ================================================================ database
def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def columns(conn, table):
    return [r[1] for r in conn.execute(f"PRAGMA table_info({table})")]


def setup_database():
    """Create / migrate the v11 database. Never touches the v10 file."""
    if not os.path.exists(DB_PATH) and os.path.exists(OLD_DB_PATH):
        shutil.copy2(OLD_DB_PATH, DB_PATH)

    conn = connect()
    conn.execute("""CREATE TABLE IF NOT EXISTS operations
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, item TEXT, qty REAL,
                     price REAL, total REAL, date_op DATE, status TEXT DEFAULT 'PENDING')""")
    cols = columns(conn, "operations")
    if "week_id" not in cols:
        conn.execute("ALTER TABLE operations ADD COLUMN week_id INTEGER")
    if "deleted_at" not in cols:
        conn.execute("ALTER TABLE operations ADD COLUMN deleted_at TEXT")
    if "prev_status" not in cols:
        conn.execute("ALTER TABLE operations ADD COLUMN prev_status TEXT")

    conn.execute("""CREATE TABLE IF NOT EXISTS articles
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL,
                     price REAL NOT NULL, active INTEGER NOT NULL DEFAULT 1)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS weeks
                    (id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT NOT NULL,
                     start_date TEXT, end_date TEXT, closed_at TEXT,
                     ops INTEGER, meters REAL, amount REAL)""")

    # Normalise old article names and seed the catalog from past operations
    conn.execute("UPDATE operations SET item = UPPER(TRIM(item)) WHERE item IS NOT NULL")
    if conn.execute("SELECT COUNT(*) FROM articles").fetchone()[0] == 0:
        rows = conn.execute("""SELECT o.item, o.price FROM operations o
                               WHERE o.id = (SELECT MAX(id) FROM operations WHERE item = o.item)""").fetchall()
        for name, price in rows:
            if name:
                conn.execute("INSERT OR IGNORE INTO articles (name, price) VALUES (?, ?)", (name, price or 0))
    conn.commit()
    conn.close()


def make_backup():
    if not os.path.exists(DB_PATH):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    dest = os.path.join(BACKUP_DIR, f"factory_v11_{datetime.now():%Y%m%d_%H%M%S}.db")
    src = sqlite3.connect(DB_PATH)
    dst = sqlite3.connect(dest)
    src.backup(dst)                       # safe copy even while in use
    dst.close(); src.close()
    files = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith("factory_v11_") and f.endswith(".db"))
    for old in files[:-BACKUPS_TO_KEEP]:
        try:
            os.remove(os.path.join(BACKUP_DIR, old))
        except OSError:
            pass
    return dest


def fmt(n, d=2):
    return f"{(n or 0):,.{d}f}"


def valid_date(s):
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def week_code(d_str):
    y, w, _ = datetime.strptime(d_str, "%Y-%m-%d").date().isocalendar()
    return f"{y}-W{w:02d}"


def safe_filename(name):
    name = re.sub(r'[\\/:*?"<>|]+', "-", name).strip(" .-")
    return name or "report"


def unique_path(folder, base, ext):
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{base}{ext}")
    i = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{base}_{i}{ext}")
        i += 1
    return path


def summary_rows(conn, where, params=()):
    """Per-article summary with WEIGHTED average price."""
    return conn.execute(f"""SELECT item, COUNT(*), SUM(qty),
                                   CASE WHEN SUM(qty) > 0 THEN SUM(total) / SUM(qty) ELSE 0 END,
                                   SUM(total)
                            FROM operations WHERE {where}
                            GROUP BY item ORDER BY SUM(total) DESC""", params).fetchall()


# ================================================================ reports
def build_pdf(path, title, period_txt, rows):
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(path, pagesize=A4)
    el = []
    if COMPANY_NAME:
        el.append(Paragraph(COMPANY_NAME, styles["Heading2"]))
    el.append(Paragraph(f"PRODUCTION REPORT: {title}", styles["Title"]))
    el.append(Paragraph(period_txt, styles["Normal"]))
    el.append(Paragraph(f"Generated: {datetime.now():%d/%m/%Y %H:%M}", styles["Normal"]))
    el.append(Spacer(1, 18))
    data = [["ARTICLE", "OPS", "TOTAL QTY", "AVG PRICE / m", "TOTAL (DA)"]]
    t_ops = t_q = t_da = 0
    for item, ops, q, avg, tot in rows:
        data.append([item, str(ops), f"{fmt(q, 1)} m", fmt(avg), fmt(tot)])
        t_ops += ops; t_q += q or 0; t_da += tot or 0
    data.append(["TOTAL", str(t_ops), f"{fmt(t_q, 1)} m", "", fmt(t_da)])
    t = Table(data, colWidths=[150, 50, 100, 100, 120])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.darkblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("BACKGROUND", (0, -1), (-1, -1), colors.lightgrey),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    el.append(t)
    el.append(Spacer(1, 24))
    el.append(Paragraph(f"GRAND TOTAL: {fmt(t_da)} DA", styles["Heading1"]))
    doc.build(el)


def build_xlsx(path, title, period_txt, rows, detail):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append([f"PRODUCTION REPORT: {title}"]); ws["A1"].font = Font(bold=True, size=14)
    ws.append([period_txt]); ws.append([])
    ws.append(["Article", "Operations", "Total Qty (m)", "Avg Price / m", "Total (DA)"])
    head = PatternFill("solid", fgColor="1F3A68")
    for c in ws[4]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = head; c.alignment = Alignment(horizontal="center")
    for item, ops, q, avg, tot in rows:
        ws.append([item, ops, round(q or 0, 2), round(avg or 0, 2), round(tot or 0, 2)])
    last = ws.max_row
    ws.append(["TOTAL", f"=SUM(B5:B{last})", f"=SUM(C5:C{last})", "", f"=SUM(E5:E{last})"])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    for col, w in zip("ABCDE", (22, 12, 16, 16, 18)):
        ws.column_dimensions[col].width = w

    ws2 = wb.create_sheet("Operations")
    ws2.append(["ID", "Date", "Article", "Qty (m)", "Price / m", "Total (DA)"])
    for c in ws2[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = head
    for r in detail:
        ws2.append(list(r))
    for col, w in zip("ABCDEF", (8, 12, 22, 12, 12, 16)):
        ws2.column_dimensions[col].width = w
    wb.save(path)


# ================================================================ app
class FactoryV11:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1350x900")
        self.edit_id = None

        setup_database()
        self.last_backup = make_backup()

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#2b2b2b", foreground="white",
                        fieldbackground="#2b2b2b", rowheight=32, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background="#3a3a3a", foreground="white",
                        font=("Segoe UI", 10, "bold"))
        style.map("Treeview", background=[("selected", BLUE)])

        self.tabs = ctk.CTkTabview(self.root, corner_radius=15, border_width=1, command=self.refresh_all)
        self.tabs.pack(fill="both", expand=True, padx=20, pady=(15, 5))
        self.t_entry = self.tabs.add("Daily Entry")
        self.t_sum = self.tabs.add("Weekly Summary")
        self.t_weeks = self.tabs.add("Weeks History")
        self.t_arch = self.tabs.add("General Archive")
        self.t_stats = self.tabs.add("Annual Analysis")
        self.t_arts = self.tabs.add("Articles")

        self.build_entry(); self.build_summary(); self.build_weeks()
        self.build_archive(); self.build_stats(); self.build_articles()

        bk = os.path.basename(self.last_backup) if self.last_backup else "none"
        ctk.CTkLabel(self.root, text=f"Database: {DB_PATH}    ·    Last backup: {bk}",
                     text_color="#b5b5b5", font=("Segoe UI", 11)).pack(pady=(0, 8))
        self.refresh_all()

    # ------------------------------------------------------------ helpers
    def tree(self, parent, cols, widths=None, height=None, select="browse"):
        f = ctk.CTkFrame(parent)
        f.pack(fill="both", expand=True, padx=10, pady=8)
        t = ttk.Treeview(f, columns=cols, show="headings", selectmode=select, height=height or 10)
        for i, c in enumerate(cols):
            t.heading(c, text=c)
            t.column(c, anchor="center", width=(widths[i] if widths else 120))
        sb = ttk.Scrollbar(f, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        t.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        t.tag_configure("deleted", foreground="#8a8a8a")
        t.tag_configure("total", background="#3a3a3a", font=("Segoe UI", 10, "bold"))
        return t

    def card(self, parent, title, color):
        f = ctk.CTkFrame(parent, border_width=2, border_color=color)
        f.pack(side="left", expand=True, fill="both", padx=8)
        ctk.CTkLabel(f, text=title, font=("Segoe UI", 13, "bold"), text_color="#b5b5b5").pack(pady=(10, 2))
        lbl = ctk.CTkLabel(f, text="0", font=("Segoe UI", 22, "bold"), text_color=color)
        lbl.pack(pady=(0, 10))
        return lbl

    @staticmethod
    def clear(t):
        for i in t.get_children():
            t.delete(i)

    def active_articles(self):
        conn = connect()
        rows = conn.execute("SELECT name, price FROM articles WHERE active = 1 ORDER BY name").fetchall()
        conn.close()
        return {n: p for n, p in rows}

    def refresh_all(self, *_):
        self.refresh_entry(); self.refresh_summary(); self.refresh_weeks()
        self.refresh_archive(); self.refresh_stats(); self.refresh_articles()

    # ------------------------------------------------------------ 1. daily entry
    def build_entry(self):
        form = ctk.CTkFrame(self.t_entry)
        form.pack(fill="x", padx=10, pady=8)
        self.lbl_form = ctk.CTkLabel(form, text="New Operation", font=("Segoe UI", 15, "bold"))
        self.lbl_form.grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 0))

        labels = ["Date (YYYY-MM-DD)", "Article", "Meters", "Price / m (auto)", "Total"]
        for i, txt in enumerate(labels):
            ctk.CTkLabel(form, text=txt, text_color="#b5b5b5").grid(row=1, column=i, sticky="w", padx=10)

        self.ent_date = ctk.CTkEntry(form, width=150)
        self.ent_date.grid(row=2, column=0, padx=10, pady=(0, 12))
        self.ent_date.insert(0, date.today().isoformat())
        self.cmb_article = ctk.CTkComboBox(form, width=200, values=[], command=lambda _: self.update_total(),
                                           state="readonly")
        self.cmb_article.grid(row=2, column=1, padx=10, pady=(0, 12))
        self.ent_qty = ctk.CTkEntry(form, width=120, placeholder_text="0.0")
        self.ent_qty.grid(row=2, column=2, padx=10, pady=(0, 12))
        self.ent_qty.bind("<KeyRelease>", lambda e: self.update_total())
        self.ent_qty.bind("<Return>", lambda e: self.save_op())
        self.lbl_price = ctk.CTkLabel(form, text="0.00 DA", width=130, anchor="w")
        self.lbl_price.grid(row=2, column=3, padx=10, pady=(0, 12))
        self.lbl_total = ctk.CTkLabel(form, text="0.00 DA", width=150, anchor="w",
                                      font=("Segoe UI", 15, "bold"), text_color="#ffd27a")
        self.lbl_total.grid(row=2, column=4, padx=10, pady=(0, 12))

        btns = ctk.CTkFrame(form, fg_color="transparent")
        btns.grid(row=3, column=0, columnspan=6, sticky="w", padx=5, pady=(0, 12))
        self.btn_save = ctk.CTkButton(btns, text="Add Operation", fg_color=GREEN, command=self.save_op)
        self.btn_save.pack(side="left", padx=5)
        self.btn_cancel = ctk.CTkButton(btns, text="Cancel edit", fg_color=GREY, command=self.cancel_edit)
        ctk.CTkButton(btns, text="Edit Selected", fg_color=BLUE, command=self.start_edit).pack(side="left", padx=5)
        ctk.CTkButton(btns, text="Delete Selected", fg_color=RED, command=self.delete_op).pack(side="left", padx=5)

        cards = ctk.CTkFrame(self.t_entry, fg_color="transparent")
        cards.pack(fill="x", padx=10, pady=4)
        self.c_ops = self.card(cards, "Operations this week", "#8ec5ff")
        self.c_m = self.card(cards, "Meters this week", "#8ec5ff")
        self.c_da = self.card(cards, "Amount this week", "#ffd27a")

        self.tree_daily = self.tree(self.t_entry, ("ID", "Date", "Article", "Qty (m)", "Price / m", "Total (DA)"))
        self.tree_daily.bind("<Double-1>", lambda e: self.start_edit())

    def update_total(self):
        arts = self.active_articles()
        price = arts.get(self.cmb_article.get(), 0)
        try:
            q = float(self.ent_qty.get().replace(",", "."))
        except ValueError:
            q = 0
        self.lbl_price.configure(text=f"{fmt(price)} DA")
        self.lbl_total.configure(text=f"{fmt(q * price)} DA")

    def refresh_entry(self):
        arts = self.active_articles()
        names = list(arts.keys())
        self.cmb_article.configure(values=names)
        if self.cmb_article.get() not in names:
            self.cmb_article.set(names[0] if names else "")
        self.update_total()

        self.clear(self.tree_daily)
        conn = connect()
        for r in conn.execute("SELECT id, date_op, item, qty, price, total FROM operations "
                              "WHERE status='PENDING' ORDER BY date_op, id"):
            self.tree_daily.insert("", "end", values=(r[0], r[1], r[2], fmt(r[3], 1), fmt(r[4]), fmt(r[5])))
        n, m, da = conn.execute("SELECT COUNT(*), SUM(qty), SUM(total) FROM operations "
                                "WHERE status='PENDING'").fetchone()
        conn.close()
        self.c_ops.configure(text=str(n)); self.c_m.configure(text=f"{fmt(m, 1)} m")
        self.c_da.configure(text=f"{fmt(da)} DA")

    def save_op(self):
        d = self.ent_date.get().strip()
        art = self.cmb_article.get()
        arts = self.active_articles()
        if not valid_date(d):
            messagebox.showerror("Error", "Date must be YYYY-MM-DD, e.g. 2026-11-05."); return
        if art not in arts:
            messagebox.showerror("Error", "Choose an article (add it in the Articles tab first)."); return
        try:
            q = float(self.ent_qty.get().replace(",", "."))
            if q <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Meters must be a number greater than 0."); return
        p = arts[art]
        conn = connect()
        if self.edit_id is None:
            conn.execute("INSERT INTO operations (item, qty, price, total, date_op) VALUES (?,?,?,?,?)",
                         (art, q, p, q * p, d))
        else:
            conn.execute("UPDATE operations SET item=?, qty=?, price=?, total=?, date_op=? "
                         "WHERE id=? AND status='PENDING'", (art, q, p, q * p, d, self.edit_id))
        conn.commit(); conn.close()
        self.cancel_edit()
        self.refresh_all()

    def start_edit(self):
        sel = self.tree_daily.selection()
        if not sel:
            messagebox.showwarning("Warning", "Select an operation to edit."); return
        op_id = int(self.tree_daily.item(sel[0])["values"][0])
        conn = connect()
        r = conn.execute("SELECT date_op, item, qty FROM operations WHERE id=?", (op_id,)).fetchone()
        conn.close()
        if not r:
            return
        if r[1] not in self.active_articles():
            messagebox.showwarning("Warning", f"Article {r[1]} is hidden in the catalog. Show it first."); return
        self.edit_id = op_id
        self.ent_date.delete(0, "end"); self.ent_date.insert(0, r[0])
        self.cmb_article.set(r[1])
        self.ent_qty.delete(0, "end"); self.ent_qty.insert(0, str(r[2]))
        self.lbl_form.configure(text=f"Editing operation #{op_id}", text_color="#ffd27a")
        self.btn_save.configure(text="Save Changes", fg_color=BLUE)
        self.btn_cancel.pack(side="left", padx=5, after=self.btn_save)
        self.update_total()

    def cancel_edit(self):
        self.edit_id = None
        self.ent_qty.delete(0, "end")
        self.lbl_form.configure(text="New Operation", text_color=("black", "white"))
        self.btn_save.configure(text="Add Operation", fg_color=GREEN)
        self.btn_cancel.pack_forget()
        self.update_total()

    def delete_op(self):
        sel = self.tree_daily.selection()
        if not sel:
            messagebox.showwarning("Warning", "Select an operation to delete."); return
        op_id = int(self.tree_daily.item(sel[0])["values"][0])
        if messagebox.askyesno("Confirm", f"Delete operation #{op_id}?\nIt goes to the Archive and can be restored."):
            conn = connect()
            conn.execute("UPDATE operations SET prev_status=status, status='DELETED', deleted_at=? "
                         "WHERE id=? AND status='PENDING'", (datetime.now().isoformat(timespec="seconds"), op_id))
            conn.commit(); conn.close()
            if self.edit_id == op_id:
                self.cancel_edit()
            self.refresh_all()

    # ------------------------------------------------------------ 2. weekly summary
    def build_summary(self):
        self.tree_sum = self.tree(self.t_sum, ("Article", "Operations", "Total Qty", "Avg Price / m (weighted)",
                                               "Total Amount (DA)"), widths=[180, 100, 140, 190, 170])
        box = ctk.CTkFrame(self.t_sum, border_width=1, border_color=ORANGE)
        box.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(box, text="Close the week", font=("Segoe UI", 15, "bold")).grid(row=0, column=0, sticky="w",
                                                                                     padx=10, pady=(10, 4))
        ctk.CTkLabel(box, text="Period name", text_color="#b5b5b5").grid(row=1, column=0, sticky="w", padx=10)
        self.ent_period = ctk.CTkEntry(box, width=260)
        self.ent_period.grid(row=2, column=0, padx=10, pady=(0, 12))
        self.lbl_range = ctk.CTkLabel(box, text="", text_color="#b5b5b5")
        self.lbl_range.grid(row=2, column=1, padx=10, pady=(0, 12))
        self.var_pdf = ctk.BooleanVar(value=True)
        self.var_xlsx = ctk.BooleanVar(value=HAS_XLSX)
        ctk.CTkCheckBox(box, text="PDF report", variable=self.var_pdf).grid(row=2, column=2, padx=10, pady=(0, 12))
        ctk.CTkCheckBox(box, text="Excel (.xlsx)" if HAS_XLSX else "Excel (pip install openpyxl)",
                        variable=self.var_xlsx, state="normal" if HAS_XLSX else "disabled"
                        ).grid(row=2, column=3, padx=10, pady=(0, 12))
        ctk.CTkButton(box, text="Close Week & Generate", fg_color=ORANGE, height=40,
                      command=self.close_week).grid(row=2, column=4, padx=10, pady=(0, 12))

    def refresh_summary(self):
        self.clear(self.tree_sum)
        conn = connect()
        rows = summary_rows(conn, "status='PENDING'")
        rng = conn.execute("SELECT MIN(date_op), MAX(date_op) FROM operations WHERE status='PENDING'").fetchone()
        conn.close()
        t_ops = t_q = t_da = 0
        for item, ops, q, avg, tot in rows:
            self.tree_sum.insert("", "end", values=(item, ops, f"{fmt(q, 1)} m", fmt(avg), fmt(tot)))
            t_ops += ops; t_q += q or 0; t_da += tot or 0
        if rows:
            self.tree_sum.insert("", "end", values=("TOTAL", t_ops, f"{fmt(t_q, 1)} m", "—", fmt(t_da)),
                                 tags=("total",))
        if rng[0]:
            self.lbl_range.configure(text=f"{rng[0]}  →  {rng[1]}")
            if not self.ent_period.get().strip() or re.fullmatch(r"\d{4}-W\d{2}", self.ent_period.get().strip()):
                self.ent_period.delete(0, "end"); self.ent_period.insert(0, week_code(rng[1]))
        else:
            self.lbl_range.configure(text="No open operations")

    def close_week(self):
        conn = connect()
        rows = summary_rows(conn, "status='PENDING'")
        detail = conn.execute("SELECT id, date_op, item, qty, price, total FROM operations "
                              "WHERE status='PENDING' ORDER BY date_op, id").fetchall()
        rng = conn.execute("SELECT MIN(date_op), MAX(date_op), COUNT(*), SUM(qty), SUM(total) "
                           "FROM operations WHERE status='PENDING'").fetchone()
        conn.close()
        if not detail:
            messagebox.showinfo("Info", "There are no open operations to close."); return
        name = self.ent_period.get().strip() or week_code(rng[1])
        if not messagebox.askyesno("Confirm",
                                   f"Close '{name}'?\n\n{rng[2]} operations · {fmt(rng[3], 1)} m · {fmt(rng[4])} DA "
                                   f"will be marked PAID.\nFiles are created first; if they fail nothing changes."):
            return

        period_txt = f"Period: {rng[0]} to {rng[1]}  ·  {rng[2]} operations"
        made = []
        try:                                   # 1) files first
            base = safe_filename(name)
            if self.var_pdf.get():
                p = unique_path(REPORTS_DIR, base, ".pdf"); build_pdf(p, name, period_txt, rows); made.append(p)
            if self.var_xlsx.get() and HAS_XLSX:
                x = unique_path(REPORTS_DIR, base, ".xlsx"); build_xlsx(x, name, period_txt, rows, detail); made.append(x)
        except Exception as e:
            for f in made:
                try:
                    os.remove(f)
                except OSError:
                    pass
            messagebox.showerror("Error", f"Report could not be created, the week was NOT closed.\n\n{e}")
            return

        conn = connect()                       # 2) then mark PAID, in one transaction
        try:
            ids = [r[0] for r in detail]
            cur = conn.execute("INSERT INTO weeks (code, start_date, end_date, closed_at, ops, meters, amount) "
                               "VALUES (?,?,?,?,?,?,?)",
                               (name, rng[0], rng[1], datetime.now().isoformat(timespec="seconds"),
                                rng[2], rng[3], rng[4]))
            wid = cur.lastrowid
            conn.executemany("UPDATE operations SET status='PAID', week_id=? WHERE id=? AND status='PENDING'",
                             [(wid, i) for i in ids])
            conn.commit()
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Error", f"Database error, nothing was changed.\n\n{e}")
            return
        finally:
            conn.close()

        make_backup()
        self.ent_period.delete(0, "end")
        self.refresh_all()
        messagebox.showinfo("Done", f"'{name}' closed and saved in Weeks History."
                            + ("\n\n" + "\n".join(made) if made else ""))
        if made:
            os.startfile(REPORTS_DIR)

    # ------------------------------------------------------------ 3. weeks history
    def build_weeks(self):
        self.tree_weeks = self.tree(self.t_weeks, ("ID", "Week", "From", "To", "Operations", "Meters",
                                                   "Amount (DA)", "Closed on"),
                                    widths=[50, 150, 110, 110, 100, 130, 150, 160], select="extended")
        bar = ctk.CTkFrame(self.t_weeks, fg_color="transparent")
        bar.pack(fill="x", padx=10, pady=4)
        ctk.CTkButton(bar, text="Reprint PDF", fg_color=RED, command=lambda: self.reprint("pdf")).pack(side="left", padx=5)
        ctk.CTkButton(bar, text="Export Excel", fg_color=GREEN, state="normal" if HAS_XLSX else "disabled",
                      command=lambda: self.reprint("xlsx")).pack(side="left", padx=5)
        ctk.CTkButton(bar, text="Compare 2 selected weeks", fg_color=BLUE, command=self.compare).pack(side="left", padx=5)
        ctk.CTkLabel(bar, text="Tip: hold Ctrl to select two weeks", text_color="#b5b5b5").pack(side="left", padx=10)
        self.lbl_compare = ctk.CTkLabel(self.t_weeks, text="", font=("Segoe UI", 14), justify="left")
        self.lbl_compare.pack(fill="x", padx=20, pady=(4, 10))

    def refresh_weeks(self):
        self.clear(self.tree_weeks)
        conn = connect()
        for r in conn.execute("SELECT id, code, start_date, end_date, ops, meters, amount, closed_at "
                              "FROM weeks ORDER BY id DESC"):
            self.tree_weeks.insert("", "end", values=(r[0], r[1], r[2], r[3], r[4], f"{fmt(r[5], 1)} m",
                                                      fmt(r[6]), (r[7] or "").replace("T", " ")))
        conn.close()

    def selected_weeks(self):
        return [self.tree_weeks.item(i)["values"] for i in self.tree_weeks.selection()]

    def reprint(self, kind):
        sel = self.selected_weeks()
        if len(sel) != 1:
            messagebox.showwarning("Warning", "Select one week."); return
        wid, name = int(sel[0][0]), str(sel[0][1])
        conn = connect()
        rows = summary_rows(conn, "week_id=? AND status='PAID'", (wid,))
        detail = conn.execute("SELECT id, date_op, item, qty, price, total FROM operations "
                              "WHERE week_id=? AND status='PAID' ORDER BY date_op, id", (wid,)).fetchall()
        w = conn.execute("SELECT start_date, end_date, ops FROM weeks WHERE id=?", (wid,)).fetchone()
        conn.close()
        period_txt = f"Period: {w[0]} to {w[1]}  ·  {w[2]} operations  ·  reprint"
        try:
            if kind == "pdf":
                path = unique_path(REPORTS_DIR, safe_filename(name) + "_reprint", ".pdf")
                build_pdf(path, name, period_txt, rows)
            else:
                path = unique_path(REPORTS_DIR, safe_filename(name), ".xlsx")
                build_xlsx(path, name, period_txt, rows, detail)
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def compare(self):
        sel = self.selected_weeks()
        if len(sel) != 2:
            messagebox.showwarning("Warning", "Select exactly two weeks (hold Ctrl)."); return
        conn = connect()
        ws = [conn.execute("SELECT code, meters, amount FROM weeks WHERE id=?", (int(s[0]),)).fetchone() for s in sel]
        conn.close()
        ws.sort(key=lambda x: x[0])
        (a, am, ad), (b, bm, bd) = ws
        dm, dd = (bm or 0) - (am or 0), (bd or 0) - (ad or 0)
        pct = (dm / am * 100) if am else 0
        self.lbl_compare.configure(
            text=f"{a}: {fmt(am, 1)} m · {fmt(ad)} DA        {b}: {fmt(bm, 1)} m · {fmt(bd)} DA\n"
                 f"Difference: {dm:+,.1f} m ({pct:+.1f} %)  ·  {dd:+,.2f} DA",
            text_color="#8ec5ff" if dm >= 0 else "#ffb27a")

    # ------------------------------------------------------------ 4. archive
    def build_archive(self):
        f = ctk.CTkFrame(self.t_arch)
        f.pack(fill="x", padx=10, pady=8)
        for i, txt in enumerate(["Search article", "From (YYYY-MM-DD)", "To (YYYY-MM-DD)", "Status"]):
            ctk.CTkLabel(f, text=txt, text_color="#b5b5b5").grid(row=0, column=i, sticky="w", padx=10, pady=(8, 0))
        self.f_q = ctk.CTkEntry(f, width=180); self.f_q.grid(row=1, column=0, padx=10, pady=(0, 10))
        self.f_from = ctk.CTkEntry(f, width=140); self.f_from.grid(row=1, column=1, padx=10, pady=(0, 10))
        self.f_to = ctk.CTkEntry(f, width=140); self.f_to.grid(row=1, column=2, padx=10, pady=(0, 10))
        self.f_status = ctk.CTkComboBox(f, width=130, values=["ALL", "PENDING", "PAID", "DELETED"], state="readonly")
        self.f_status.set("ALL"); self.f_status.grid(row=1, column=3, padx=10, pady=(0, 10))
        ctk.CTkButton(f, text="Apply", fg_color=BLUE, width=90, command=self.refresh_archive).grid(row=1, column=4, padx=5)
        ctk.CTkButton(f, text="Clear", fg_color=GREY, width=90, command=self.clear_filters).grid(row=1, column=5, padx=5)
        self.f_q.bind("<Return>", lambda e: self.refresh_archive())

        self.lbl_arch = ctk.CTkLabel(self.t_arch, text="", text_color="#ffd27a", font=("Segoe UI", 13, "bold"))
        self.lbl_arch.pack(anchor="w", padx=20)
        self.tree_arch = self.tree(self.t_arch, ("ID", "Date", "Article", "Qty (m)", "Total (DA)", "Week", "Status"))
        ctk.CTkButton(self.t_arch, text="Restore deleted operation", fg_color=BLUE,
                      command=self.restore_op).pack(pady=(0, 8))

    def clear_filters(self):
        for e in (self.f_q, self.f_from, self.f_to):
            e.delete(0, "end")
        self.f_status.set("ALL")
        self.refresh_archive()

    def refresh_archive(self):
        where, params = ["1=1"], []
        q, d1, d2, st = self.f_q.get().strip(), self.f_from.get().strip(), self.f_to.get().strip(), self.f_status.get()
        if q:
            where.append("o.item LIKE ?"); params.append(f"%{q.upper()}%")
        if d1 and valid_date(d1):
            where.append("o.date_op >= ?"); params.append(d1)
        if d2 and valid_date(d2):
            where.append("o.date_op <= ?"); params.append(d2)
        if st != "ALL":
            where.append("o.status = ?"); params.append(st)
        self.clear(self.tree_arch)
        conn = connect()
        rows = conn.execute(f"""SELECT o.id, o.date_op, o.item, o.qty, o.total, COALESCE(w.code, '—'), o.status
                                FROM operations o LEFT JOIN weeks w ON w.id = o.week_id
                                WHERE {' AND '.join(where)} ORDER BY o.date_op DESC, o.id DESC""", params).fetchall()
        total_all = conn.execute("SELECT COUNT(*) FROM operations").fetchone()[0]
        conn.close()
        s_m = s_da = 0
        for r in rows:
            tag = ("deleted",) if r[6] == "DELETED" else ()
            if r[6] != "DELETED":
                s_m += r[3] or 0; s_da += r[4] or 0
            self.tree_arch.insert("", "end", values=(r[0], r[1], r[2], fmt(r[3], 1), fmt(r[4]), r[5], r[6]), tags=tag)
        self.lbl_arch.configure(text=f"Showing {len(rows)} of {total_all} operations   ·   "
                                     f"Total (without deleted): {fmt(s_m, 1)} m · {fmt(s_da)} DA")

    def restore_op(self):
        sel = self.tree_arch.selection()
        if not sel:
            messagebox.showwarning("Warning", "Select a DELETED operation."); return
        v = self.tree_arch.item(sel[0])["values"]
        if v[6] != "DELETED":
            messagebox.showinfo("Info", "Only DELETED operations can be restored."); return
        conn = connect()
        conn.execute("UPDATE operations SET status=COALESCE(prev_status, 'PENDING'), deleted_at=NULL, "
                     "prev_status=NULL WHERE id=?", (int(v[0]),))
        conn.commit(); conn.close()
        self.refresh_all()

    # ------------------------------------------------------------ 5. annual analysis
    def build_stats(self):
        top = ctk.CTkFrame(self.t_stats, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(top, text="Year").pack(side="left", padx=(5, 8))
        self.cmb_year = ctk.CTkComboBox(top, width=110, values=[str(date.today().year)], state="readonly",
                                        command=lambda _: self.refresh_stats())
        self.cmb_year.set(str(date.today().year)); self.cmb_year.pack(side="left")
        ctk.CTkButton(top, text="Export year to Excel", fg_color=GREEN, state="normal" if HAS_XLSX else "disabled",
                      command=self.export_year).pack(side="right", padx=5)

        cards = ctk.CTkFrame(self.t_stats, fg_color="transparent")
        cards.pack(fill="x", padx=10, pady=4)
        self.s_m = self.card(cards, "Total meters", "#8ec5ff")
        self.s_da = self.card(cards, "Total amount", "#ffd27a")
        self.s_avg = self.card(cards, "Avg price / m (weighted)", "#e6e6e6")
        self.s_best = self.card(cards, "Best month", "#e6e6e6")

        self.chart = ctk.CTkCanvas(self.t_stats, height=230, bg="#242424", highlightthickness=0)
        self.chart.pack(fill="x", padx=18, pady=8)
        self.chart.bind("<Configure>", lambda e: self.draw_chart())
        self.month_vals = [0] * 12
        self.tree_year = self.tree(self.t_stats, ("Article", "Annual Qty", "Revenue (DA)", "Share of revenue"), height=6)

    def refresh_stats(self):
        conn = connect()
        years = [r[0] for r in conn.execute("SELECT DISTINCT substr(date_op,1,4) FROM operations "
                                            "WHERE status!='DELETED' AND date_op IS NOT NULL ORDER BY 1 DESC")]
        if str(date.today().year) not in years:
            years.insert(0, str(date.today().year))
        self.cmb_year.configure(values=years)
        y = self.cmb_year.get() if self.cmb_year.get() in years else years[0]
        self.cmb_year.set(y)

        flt = "status!='DELETED' AND substr(date_op,1,4)=?"
        m, da = conn.execute(f"SELECT SUM(qty), SUM(total) FROM operations WHERE {flt}", (y,)).fetchone()
        m, da = m or 0, da or 0
        self.month_vals = [0] * 12
        for mon, q in conn.execute(f"SELECT CAST(substr(date_op,6,2) AS INTEGER), SUM(qty) FROM operations "
                                   f"WHERE {flt} GROUP BY 1", (y,)):
            if mon and 1 <= mon <= 12:
                self.month_vals[mon - 1] = q or 0
        rows = conn.execute(f"SELECT item, SUM(qty), SUM(total) FROM operations WHERE {flt} "
                            f"GROUP BY item ORDER BY SUM(total) DESC", (y,)).fetchall()
        conn.close()

        self.s_m.configure(text=f"{fmt(m, 1)} m"); self.s_da.configure(text=f"{fmt(da)} DA")
        self.s_avg.configure(text=f"{fmt(da / m if m else 0)} DA")
        best = max(range(12), key=lambda i: self.month_vals[i])
        self.s_best.configure(text=f"{MONTHS[best]} · {fmt(self.month_vals[best], 0)} m" if m else "—")
        self.clear(self.tree_year)
        for item, q, tot in rows:
            self.tree_year.insert("", "end", values=(item, f"{fmt(q, 1)} m", fmt(tot),
                                                     f"{(tot / da * 100) if da else 0:.1f} %"))
        self.draw_chart()

    def draw_chart(self):
        c = self.chart
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 50:
            return
        mx = max(self.month_vals) or 1
        best = max(range(12), key=lambda i: self.month_vals[i])
        slot = (w - 40) / 12
        base = h - 28
        for i, v in enumerate(self.month_vals):
            x0 = 20 + i * slot + slot * 0.2
            x1 = 20 + (i + 1) * slot - slot * 0.2
            bh = (v / mx) * (base - 30)
            color = "#ffb347" if (i == best and v) else "#3a86c8"
            if v:
                c.create_rectangle(x0, base - bh, x1, base, fill=color, outline="")
                c.create_text((x0 + x1) / 2, base - bh - 10, text=f"{v:,.0f}", fill="#d0d0d0", font=("Segoe UI", 9))
            c.create_text((x0 + x1) / 2, base + 14, text=MONTHS[i], fill="#b5b5b5", font=("Segoe UI", 10))
        c.create_line(20, base, w - 20, base, fill="#4a4a4a")

    def export_year(self):
        y = self.cmb_year.get()
        conn = connect()
        rows = summary_rows(conn, "status!='DELETED' AND substr(date_op,1,4)=?", (y,))
        detail = conn.execute("SELECT id, date_op, item, qty, price, total FROM operations "
                              "WHERE status!='DELETED' AND substr(date_op,1,4)=? ORDER BY date_op, id", (y,)).fetchall()
        conn.close()
        try:
            path = unique_path(REPORTS_DIR, f"Year_{y}", ".xlsx")
            build_xlsx(path, f"YEAR {y}", f"Year {y}  ·  {len(detail)} operations", rows, detail)
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ------------------------------------------------------------ 6. articles
    def build_articles(self):
        f = ctk.CTkFrame(self.t_arts)
        f.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(f, text="Article name", text_color="#b5b5b5").grid(row=0, column=0, sticky="w", padx=10, pady=(8, 0))
        ctk.CTkLabel(f, text="Price per meter (DA)", text_color="#b5b5b5").grid(row=0, column=1, sticky="w", padx=10, pady=(8, 0))
        self.a_name = ctk.CTkEntry(f, width=220, placeholder_text="e.g. SKAI BEIGE")
        self.a_name.grid(row=1, column=0, padx=10, pady=(0, 10))
        self.a_price = ctk.CTkEntry(f, width=140, placeholder_text="0.00")
        self.a_price.grid(row=1, column=1, padx=10, pady=(0, 10))
        ctk.CTkButton(f, text="Add article", fg_color=GREEN, command=self.add_article).grid(row=1, column=2, padx=5)
        ctk.CTkButton(f, text="Set price of selected", fg_color=BLUE, command=self.set_price).grid(row=1, column=3, padx=5)
        ctk.CTkButton(f, text="Show / Hide selected", fg_color=GREY, command=self.toggle_article).grid(row=1, column=4, padx=5)
        ctk.CTkLabel(self.t_arts, text="A new price applies to new operations only. Closed weeks keep their price.",
                     text_color="#b5b5b5").pack(anchor="w", padx=20)
        self.tree_arts = self.tree(self.t_arts, ("Article", "Price / m (DA)", "Used in operations", "In entry list"))
        self.tree_arts.bind("<<TreeviewSelect>>", self.fill_article_form)

    def refresh_articles(self):
        self.clear(self.tree_arts)
        conn = connect()
        for name, price, active, used in conn.execute(
                "SELECT a.name, a.price, a.active, (SELECT COUNT(*) FROM operations o "
                "WHERE o.item = a.name AND o.status != 'DELETED') FROM articles a ORDER BY a.name"):
            self.tree_arts.insert("", "end", values=(name, fmt(price), used, "Shown" if active else "Hidden"),
                                  tags=() if active else ("deleted",))
        conn.close()

    def fill_article_form(self, _=None):
        sel = self.tree_arts.selection()
        if sel:
            v = self.tree_arts.item(sel[0])["values"]
            self.a_name.delete(0, "end"); self.a_name.insert(0, v[0])
            self.a_price.delete(0, "end"); self.a_price.insert(0, str(v[1]).replace(",", ""))

    def read_price(self):
        try:
            p = float(self.a_price.get().replace(",", "."))
            if p <= 0:
                raise ValueError
            return p
        except ValueError:
            messagebox.showerror("Error", "Price must be a number greater than 0.")
            return None

    def add_article(self):
        name = self.a_name.get().strip().upper()
        if not name:
            messagebox.showerror("Error", "Enter an article name."); return
        p = self.read_price()
        if p is None:
            return
        conn = connect()
        try:
            conn.execute("INSERT INTO articles (name, price) VALUES (?, ?)", (name, p))
            conn.commit()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"{name} is already in the catalog.")
        finally:
            conn.close()
        self.a_name.delete(0, "end"); self.a_price.delete(0, "end")
        self.refresh_all()

    def set_price(self):
        sel = self.tree_arts.selection()
        if not sel:
            messagebox.showwarning("Warning", "Select an article."); return
        p = self.read_price()
        if p is None:
            return
        name = self.tree_arts.item(sel[0])["values"][0]
        conn = connect()
        conn.execute("UPDATE articles SET price=? WHERE name=?", (p, name))
        conn.commit(); conn.close()
        self.refresh_all()

    def toggle_article(self):
        sel = self.tree_arts.selection()
        if not sel:
            messagebox.showwarning("Warning", "Select an article."); return
        name = self.tree_arts.item(sel[0])["values"][0]
        conn = connect()
        conn.execute("UPDATE articles SET active = 1 - active WHERE name=?", (name,))
        conn.commit(); conn.close()
        self.refresh_all()


if __name__ == "__main__":
    app = ctk.CTk()
    FactoryV11(app)
    app.mainloop()
