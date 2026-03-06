import customtkinter as ctk
from tkinter import ttk, messagebox, simpledialog
import sqlite3
from datetime import datetime
import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class GlobalFactoryFinal:
    def __init__(self, root):
        self.root = root
        self.root.title("GLOBAL FACTORY MANAGER Pro v1.1")
        self.root.geometry("1300x850")
        self.setup_database()

        self.tabs = ctk.CTkTabview(self.root, corner_radius=15)
        self.tabs.pack(fill="both", expand=True, padx=20, pady=20)

        self.tab_input = self.tabs.add("📝 Data Entry")
        self.tab_filter = self.tabs.add("🔍 Inventory Summary")
        self.tab_archive = self.tabs.add("📚 Records")
        self.tab_stats = self.tabs.add("📊 Global Analytics")

        self.build_input_page()
        self.build_filter_page()
        self.build_archive_page()
        self.build_stats_page()
        self.refresh_all_data()

    def setup_database(self):
        conn = sqlite3.connect('global_factory_v1.db')
        conn.execute('''CREATE TABLE IF NOT EXISTS operations 
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, item TEXT, qty REAL, unit TEXT,
                       price REAL, currency TEXT, total REAL, date_op DATE, status TEXT DEFAULT 'PENDING')''')
        conn.commit(); conn.close()

    def build_input_page(self):
        frame = ctk.CTkFrame(self.tab_input)
        frame.pack(fill="x", padx=20, pady=10)
        
        self.ent_item = ctk.CTkEntry(frame, placeholder_text="Item Name", width=180)
        self.ent_item.grid(row=0, column=0, padx=5, pady=20)
        
        self.ent_qty = ctk.CTkEntry(frame, placeholder_text="Qty", width=80)
        self.ent_qty.grid(row=0, column=1, padx=5)
        
        self.unit_var = ctk.StringVar(value="Meters")
        self.unit_menu = ctk.CTkOptionMenu(frame, values=["Meters", "Pieces", "Kg"], variable=self.unit_var, width=100)
        self.unit_menu.grid(row=0, column=2, padx=5)
        
        self.ent_price = ctk.CTkEntry(frame, placeholder_text="Price", width=80)
        self.ent_price.grid(row=0, column=3, padx=5)
        
        self.curr_var = ctk.StringVar(value="USD")
        self.curr_menu = ctk.CTkOptionMenu(frame, values=["DA", "USD", "EUR", "TND", "MAD"], variable=self.curr_var, width=80)
        self.curr_menu.grid(row=0, column=4, padx=5)
        
        ctk.CTkButton(frame, text="Add Entry ✅", command=self.add_op, fg_color="#27ae60").grid(row=0, column=5, padx=5)
        ctk.CTkButton(frame, text="Delete 🗑️", command=lambda: self.delete_op(self.tree_daily), fg_color="#c0392b").grid(row=0, column=6, padx=5)

        self.tree_daily = self.create_styled_tree(self.tab_input, ("ID", "Date", "Item", "Qty", "Unit", "Price", "Currency", "Total"))

    def build_filter_page(self):
        self.tree_filter = self.create_styled_tree(self.tab_filter, ("Item", "Total Qty", "Unit", "Avg Price", "Currency", "Final Total"))
        ctk.CTkButton(self.tab_filter, text="Generate Professional Invoice (PDF) 📄", command=self.close_week, height=45, fg_color="#d35400").pack(pady=10)

    def build_archive_page(self):
        self.tree_archive = self.create_styled_tree(self.tab_archive, ("ID", "Date", "Item", "Qty", "Total", "Status"))
        ctk.CTkButton(self.tab_archive, text="Delete Record ⚠️", command=lambda: self.delete_op(self.tree_archive), fg_color="#c0392b", width=200).pack(pady=5)

    def build_stats_page(self):
        f = ctk.CTkFrame(self.tab_stats); f.pack(fill="x", padx=20, pady=10)
        self.lbl_annual_qty = self.create_card(f, "Global Volume", "0 Units", "#3498db")
        self.lbl_annual_money = self.create_card(f, "Global Revenue", "0.00", "#f1c40f")
        
        # إضافة عنوان للجدول السنوي
        ctk.CTkLabel(self.tab_stats, text="📊 Annual Performance Analysis", font=("Arial", 16, "bold")).pack(pady=10)
        self.tree_stats_yearly = self.create_styled_tree(self.tab_stats, ("Product", "Annual Qty", "Revenue Share", "% Contribution"))

    def create_card(self, parent, title, val, color):
        card = ctk.CTkFrame(parent, border_width=2, border_color=color); card.pack(side="left", expand=True, padx=10, fill="both")
        ctk.CTkLabel(card, text=title, font=("Arial", 14, "bold")).pack(pady=10)
        l = ctk.CTkLabel(card, text=val, font=("Arial", 25, "bold"), text_color=color); l.pack(pady=10)
        return l

    def create_styled_tree(self, parent, cols):
        f = ctk.CTkFrame(parent); f.pack(fill="both", expand=True, padx=10, pady=10)
        t = ttk.Treeview(f, columns=cols, show='headings')
        for c in cols:
            t.heading(c, text=c)
            t.column(c, anchor="center", width=120)
        t.pack(fill="both", expand=True)
        return t

    def refresh_all_data(self):
        # تصفير الجداول قبل التحديث
        for t in [self.tree_daily, self.tree_filter, self.tree_archive, self.tree_stats_yearly]:
            for i in t.get_children(): t.delete(i)
            
        conn = sqlite3.connect('global_factory_v1.db')
        
        # 1. تحديث جدول البيانات اليومية
        for r in conn.execute("SELECT id, date_op, item, qty, unit, price, currency, total FROM operations WHERE status='PENDING'"):
            self.tree_daily.insert("", "end", values=r)
            
        # 2. تحديث جدول الملخص
        for r in conn.execute("SELECT UPPER(item), SUM(qty), unit, AVG(price), currency, SUM(total) FROM operations WHERE status='PENDING' GROUP BY UPPER(item), unit, currency"):
            self.tree_filter.insert("", "end", values=(r[0], f"{r[1]:,.1f}", r[2], f"{r[3]:,.1f}", r[4], f"{r[5]:,.2f}"))
            
        # 3. تحديث السجل العام
        for r in conn.execute("SELECT id, date_op, item, qty, total, status FROM operations ORDER BY id DESC LIMIT 50"):
            self.tree_archive.insert("", "end", values=r)
            
        # 4. تحديث الإحصائيات السنوية والجدول التحليلي
        st = conn.execute("SELECT SUM(qty), SUM(total) FROM operations").fetchone()
        total_q = st[0] or 0
        total_m = st[1] or 0
        self.lbl_annual_qty.configure(text=f"{total_q:,.1f} Units")
        self.lbl_annual_money.configure(text=f"{total_m:,.2f} (Mix)")
        
        # ملء جدول التحليل السنوي (المشكلة التي ذكرتها)
        for r in conn.execute("SELECT UPPER(item), SUM(qty), SUM(total) FROM operations GROUP BY UPPER(item) ORDER BY SUM(total) DESC"):
            percent = (r[2] / total_m * 100) if total_m > 0 else 0
            self.tree_stats_yearly.insert("", "end", values=(r[0], f"{r[1]:,.1f}", f"{r[2]:,.2f}", f"{percent:.1f}%"))
            
        conn.close()

    def add_op(self):
        try:
            it, q, u, p, c = self.ent_item.get().strip(), float(self.ent_qty.get()), self.unit_var.get(), float(self.ent_price.get()), self.curr_var.get()
            if not it: return
            conn = sqlite3.connect('global_factory_v1.db')
            conn.execute("INSERT INTO operations (item, qty, unit, price, currency, total, date_op) VALUES (?,?,?,?,?,?,?)",
                         (it, q, u, p, c, q*p, datetime.now().strftime("%Y-%m-%d")))
            conn.commit(); conn.close()
            
            # مسح الخانات بعد الإضافة (التعديل المطلوب)
            self.ent_item.delete(0, 'end')
            self.ent_qty.delete(0, 'end')
            self.ent_price.delete(0, 'end')
            
            self.refresh_all_data()
        except: messagebox.showerror("Error", "Please check your numbers!")

    def delete_op(self, target_tree):
        sel = target_tree.selection()
        if not sel: return
        if messagebox.askyesno("Confirm", "Are you sure you want to delete this?"):
            row_id = target_tree.item(sel[0])['values'][0]
            conn = sqlite3.connect('global_factory_v1.db')
            conn.execute("DELETE FROM operations WHERE id=?", (row_id,))
            conn.commit(); conn.close()
            self.refresh_all_data()

    def close_week(self):
        period = simpledialog.askstring("Invoice", "Enter Client or Period Name:")
        if not period: return
        self.gen_pdf(period)
        conn = sqlite3.connect('global_factory_v1.db')
        conn.execute("UPDATE operations SET status='PAID' WHERE status='PENDING'")
        conn.commit(); conn.close()
        self.refresh_all_data()

    def gen_pdf(self, period):
        fname = f"Invoice_{period}.pdf"
        doc = SimpleDocTemplate(fname, pagesize=letter)
        styles = getSampleStyleSheet()
        elements = [Paragraph(f"INVOICE REPORT: {period}", styles['Title']), Spacer(1,20)]
        data = [["ARTICLE", "QTY", "UNIT", "AVG PRICE", "CURRENCY", "TOTAL"]]
        conn = sqlite3.connect('global_factory_v1.db')
        gt = 0
        for r in conn.execute("SELECT UPPER(item), SUM(qty), unit, AVG(price), currency, SUM(total) FROM operations WHERE status='PENDING' GROUP BY UPPER(item), unit, currency"):
            data.append([r[0], f"{r[1]:,.1f}", r[2], f"{r[3]:,.1f}", r[4], f"{r[5]:,.2f}"]); gt += r[5]
        t = Table(data, colWidths=[120, 80, 80, 80, 80, 100])
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.darkblue),('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),('GRID',(0,0),(-1,-1),1,colors.black),('ALIGN',(0,0),(-1,-1),'CENTER')]))
        elements.append(t); elements.append(Spacer(1,30))
        elements.append(Paragraph(f"GRAND TOTAL: {gt:,.2f}", styles['Heading1']))
        doc.build(elements); conn.close()
        os.startfile(os.getcwd())

if __name__ == "__main__":
    app = ctk.CTk(); gui = GlobalFactoryFinal(app); app.mainloop()