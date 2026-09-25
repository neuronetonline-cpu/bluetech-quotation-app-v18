"""Bluetech PC-build job sheet window; bundled alongside quotation_app.py."""
import json
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from xml.sax.saxutils import escape
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

CHECKS = ["Motherboard / CPU / cooler installed", "RAM / SSD / HDD installed", "PSU / GPU / cabling checked", "BIOS / boot verified", "Windows / drivers installed", "USB / audio / network tested", "Display / peripherals tested", "Temperature / stress test", "Final cleaning / accessories checked"]
STAGES = ["Prepared By", "PC Built By", "Final Checked By", "POS Entered By"]
STATUSES = ["Pending", "Approved", "Building", "Built", "Final Checked", "POS Entered", "Completed"]


def setup_db(db):
    con = db()
    con.execute("""CREATE TABLE IF NOT EXISTS job_chits (
        id INTEGER PRIMARY KEY AUTOINCREMENT, job_no TEXT UNIQUE NOT NULL,
        quotation_id INTEGER, quotation_no TEXT, customer TEXT, phone TEXT,
        created_at TEXT, due_date TEXT, status TEXT, items_json TEXT,
        checklist_json TEXT, staff_json TEXT, timestamps_json TEXT, remarks TEXT
    )""")
    con.commit(); con.close()


def next_job_no(db):
    prefix = datetime.now().strftime('JOB-%Y%m%d-')
    con = db(); rows = con.execute('SELECT job_no FROM job_chits WHERE job_no LIKE ?', (prefix+'%',)).fetchall(); con.close()
    nums = []
    for (value,) in rows:
        try: nums.append(int(value.rsplit('-', 1)[1]))
        except (ValueError, IndexError): pass
    return prefix + f'{max(nums, default=0)+1:04d}'


def open_job_chit(app, db, get_pdf_dir, job_id=None):
    setup_db(db)
    saved = None
    if job_id is not None:
        con = db(); saved = con.execute('SELECT * FROM job_chits WHERE id=?', (job_id,)).fetchone(); con.close()
        if not saved:
            messagebox.showerror('Job Sheet', 'Job sheet not found.'); return
    else:
        if not app.customer.get().strip():
            messagebox.showwarning('Job Sheet', 'Enter the customer name first.'); return
        items = app.collect_items()
        if not items:
            messagebox.showwarning('Job Sheet', 'Add at least one product first.'); return
        # Persist quotation before creating its workshop document.
        qid = app.save_quote(show_message=False)
        if qid is None: return
        con = db(); saved = con.execute('SELECT * FROM job_chits WHERE quotation_id=? ORDER BY id DESC LIMIT 1', (qid,)).fetchone(); con.close()

    win = tk.Toplevel(app.root); win.title('Bluetech Computers - PC Build Job Sheet')
    win.geometry('1020x780'); win.minsize(780, 550)
    win.transient(app.root); win.grab_set(); win.focus_force()
    win.configure(bg='#F3F7FC')
    outer = ttk.Frame(win); outer.pack(fill='both', expand=True)
    canvas = tk.Canvas(outer, highlightthickness=0, bg='#F3F7FC')
    scrollbar = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side='left', fill='both', expand=True); scrollbar.pack(side='right', fill='y')
    body = ttk.Frame(canvas, padding=15); canvas_window = canvas.create_window((0,0), window=body, anchor='nw')
    body.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
    canvas.bind('<Configure>', lambda e: canvas.itemconfigure(canvas_window, width=e.width))
    def wheel(e):
        canvas.yview_scroll(int(-e.delta/120), 'units')
    canvas.bind('<Enter>', lambda e: canvas.bind_all('<MouseWheel>', wheel))
    canvas.bind('<Leave>', lambda e: canvas.unbind_all('<MouseWheel>'))

    if saved:
        jid, number, qid, qno, customer, phone, created, due, status, itemjson, checkjson, staffjson, timejson, notes = saved
        items = json.loads(itemjson or '[]'); checks = json.loads(checkjson or '{}')
        staff = json.loads(staffjson or '{}'); stamps = json.loads(timejson or '{}')
    else:
        jid = None; number = next_job_no(db); qid = app.editing_id; qno = app.qno.get()
        customer = app.customer.get(); phone = app.phone.get(); created = datetime.now().strftime('%Y-%m-%d %H:%M')
        due = ''; status = 'Approved'; items = [list(x[:3]) for x in app.collect_items()]
        checks = {}; staff = {'Prepared By': app.prepared_by.get()}; stamps = {'Prepared By': created}; notes = ''

    ttk.Label(body, text='PC BUILD JOB SHEET', font=('Segoe UI', 19, 'bold'), foreground='#075EAA').pack(anchor='w')
    ttk.Label(body, text='Internal workshop document — costs and selling prices are excluded.', foreground='#667085').pack(anchor='w', pady=(0,12))
    info = ttk.LabelFrame(body, text='JOB / CUSTOMER DETAILS', padding=10); info.pack(fill='x', pady=5)
    def info_row(r, label, val):
        ttk.Label(info, text=label, font=('Segoe UI',9,'bold')).grid(row=r, column=0, sticky='w', pady=4, padx=5)
        ttk.Label(info, text=val).grid(row=r, column=1, sticky='w', pady=4, padx=5)
    info_row(0, 'Job Sheet No.', number); info_row(1, 'Quotation No.', qno)
    info_row(2, 'Customer', customer); info_row(3, 'Phone', phone); info_row(4, 'Created', created)
    ttk.Label(info, text='Due Date (YYYY-MM-DD)').grid(row=5, column=0, sticky='w', padx=5)
    due_var = tk.StringVar(value=due or ''); ttk.Entry(info, textvariable=due_var, width=24).grid(row=5,column=1,sticky='w',padx=5,pady=5)
    ttk.Label(info, text='Status').grid(row=6,column=0,sticky='w',padx=5)
    status_var = tk.StringVar(value=status or 'Approved')
    ttk.Combobox(info, textvariable=status_var, values=STATUSES, state='readonly', width=22).grid(row=6,column=1,sticky='w',padx=5,pady=5)

    parts = ttk.LabelFrame(body, text='BUILD COMPONENTS / JOB SHEET ITEMS', padding=10)
    parts.pack(fill='x', pady=7)
    for col, title in enumerate(['PRODUCT', 'DESCRIPTION', 'QTY', 'REMOVE']):
        ttk.Label(parts, text=title, font=('Segoe UI',9,'bold')).grid(row=0,column=col,sticky='ew',padx=5,pady=(0,4))

    item_rows = []

    def add_item_row(product='', description='', qty='1', removable=True):
        row = len(item_rows) + 1
        pv = tk.StringVar(value=str(product or ''))
        dv = tk.StringVar(value=str(description or ''))
        qv = tk.StringVar(value=str(qty if qty is not None else '1'))
        ttk.Entry(parts, textvariable=pv).grid(row=row,column=0,sticky='ew',padx=5,pady=2)
        ttk.Entry(parts, textvariable=dv).grid(row=row,column=1,sticky='ew',padx=5,pady=2)
        ttk.Entry(parts, textvariable=qv, width=9, justify='center').grid(row=row,column=2,sticky='ew',padx=5,pady=2)
        btn = ttk.Button(parts, text='REMOVE', style='Light.TButton', command=lambda r=len(item_rows): remove_item_row(r))
        btn.grid(row=row,column=3,padx=5,pady=2)
        item_rows.append((pv,dv,qv,btn))

    def remove_item_row(index):
        if index < 0 or index >= len(item_rows):
            return
        for widget in item_rows[index]:
            if hasattr(widget, 'destroy'):
                try: widget.destroy()
                except tk.TclError: pass
        item_rows.pop(index)
        # Rebuild rows so numbering/grid positions stay clean.
        rows = [(p.get(), d.get(), q.get()) for p,d,q,_ in item_rows]
        for child in parts.grid_slaves():
            info = child.grid_info()
            r = int(info.get('row', 0))
            if r > 0 and r != 999:
                child.destroy()
        item_rows.clear()
        for p,d,q in rows:
            add_item_row(p,d,q)
        parts.update_idletasks()

    for p, d, q in items:
        add_item_row(p, d, q)
    if not item_rows:
        add_item_row('', '', '1')

    ttk.Button(parts, text='＋ ADD ITEM', style='Blue.TButton', command=lambda: add_item_row('', '', '1')).grid(
        row=999, column=0, columnspan=4, sticky='w', padx=5, pady=(7,2)
    )
    parts.columnconfigure(0, weight=1, minsize=180)
    parts.columnconfigure(1, weight=2, minsize=260)
    parts.columnconfigure(2, weight=0, minsize=70)
    parts.columnconfigure(3, weight=0, minsize=85)

    people = ttk.LabelFrame(body, text='STAFF / RESPONSIBILITY', padding=10); people.pack(fill='x', pady=7)
    names = [n for _,n in app_get_users(db)]
    if not names: names = ['Admin']
    staff_vars = {}; time_vars = {}
    for i, stage in enumerate(STAGES):
        ttk.Label(people, text=stage, font=('Segoe UI',9,'bold')).grid(row=i,column=0,sticky='w',padx=5,pady=5)
        v = tk.StringVar(value=staff.get(stage, '')); t = tk.StringVar(value=stamps.get(stage, ''))
        staff_vars[stage] = v; time_vars[stage] = t
        ttk.Combobox(people, textvariable=v, values=['']+names, width=27).grid(row=i,column=1,sticky='w',padx=5)
        ttk.Label(people, textvariable=t, width=21).grid(row=i,column=2,sticky='w',padx=5)
        def mark(s=stage, sv=v, tv=t):
            if not sv.get().strip(): messagebox.showwarning('Staff', 'Select or enter staff name.', parent=win); return
            tv.set(datetime.now().strftime('%Y-%m-%d %H:%M'))
        ttk.Button(people, text='MARK DONE', command=mark).grid(row=i,column=3,padx=5)

    check_frame = ttk.LabelFrame(body, text='BUILD / FINAL CHECKLIST', padding=10); check_frame.pack(fill='x', pady=7)
    check_vars = {}
    for i, item in enumerate(CHECKS):
        v = tk.BooleanVar(value=bool(checks.get(item,False))); check_vars[item] = v
        ttk.Checkbutton(check_frame, text=item, variable=v).grid(row=i//2,column=i%2,sticky='w',padx=8,pady=4)
    notes_frame = ttk.LabelFrame(body, text='WORKSHOP REMARKS / SERIAL NUMBERS', padding=10); notes_frame.pack(fill='x', pady=7)
    remarks = tk.Text(notes_frame, height=5, wrap='word'); remarks.insert('1.0',notes or ''); remarks.pack(fill='x')

    def current_items():
        result = []
        for pv, dv, qv, _ in item_rows:
            product = pv.get().strip()
            description = dv.get().strip()
            qty_text = qv.get().strip() or '1'
            if not product and not description:
                continue
            try:
                qty = float(qty_text)
                if qty.is_integer():
                    qty = int(qty)
                if qty <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror('Job Sheet', f'Invalid quantity: {qty_text}', parent=win)
                return None
            result.append([product or 'CUSTOM ITEM', description, qty])
        if not result:
            messagebox.showwarning('Job Sheet', 'Add at least one item.', parent=win)
            return None
        return result

    def payload(current):
        return (due_var.get().strip(), status_var.get(), json.dumps(current,ensure_ascii=False),
                json.dumps({k:v.get() for k,v in check_vars.items()}),
                json.dumps({k:v.get().strip() for k,v in staff_vars.items()}),
                json.dumps({k:v.get() for k,v in time_vars.items()}), remarks.get('1.0','end-1c').strip())
    def save(silent=False):
        nonlocal jid
        current = current_items()
        if current is None:
            return False
        due, status, item_json, checks_json, staff_json, stamps_json, notes = payload(current)
        if due:
            try: datetime.strptime(due,'%Y-%m-%d')
            except ValueError: messagebox.showerror('Due date','Use YYYY-MM-DD format.',parent=win); return False
        con = db()
        try:
            if jid:
                con.execute('''UPDATE job_chits SET due_date=?,status=?,items_json=?,checklist_json=?,staff_json=?,timestamps_json=?,remarks=? WHERE id=?''',
                            (due,status,item_json,checks_json,staff_json,stamps_json,notes,jid))
            else:
                cur = con.execute('''INSERT INTO job_chits (job_no,quotation_id,quotation_no,customer,phone,created_at,due_date,status,items_json,checklist_json,staff_json,timestamps_json,remarks)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''', (number,qid,qno,customer,phone,created,due,status,item_json,checks_json,staff_json,stamps_json,notes))
                jid = cur.lastrowid
            con.commit()
        except Exception as e:
            messagebox.showerror('Job Sheet',str(e),parent=win); return False
        finally: con.close()
        if not silent: messagebox.showinfo('Saved',f'{number} saved.',parent=win)
        return True

    def make_pdf():
        if not save(silent=True):
            return None
        folder = os.path.join(get_pdf_dir(), 'Job Sheets')
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, number + '.pdf')

        # Simple black-and-white A4 landscape workshop form.
        # Keep the page filled with useful data; avoid decorative graphics.
        page = landscape(A4)
        styles = getSampleStyleSheet()
        body = ParagraphStyle(
            'job_bw_body', parent=styles['Normal'], fontName='Helvetica',
            fontSize=8.2, leading=9.5, textColor=colors.black
        )
        body_b = ParagraphStyle(
            'job_bw_body_b', parent=body, fontName='Helvetica-Bold'
        )
        section = ParagraphStyle(
            'job_bw_section', parent=body_b, fontSize=9, leading=10,
            spaceBefore=3, spaceAfter=3, textColor=colors.black
        )
        title_style = ParagraphStyle(
            'job_bw_title', parent=styles['Title'], fontName='Helvetica-Bold',
            fontSize=13, leading=14, alignment=1, textColor=colors.black,
            spaceAfter=2
        )
        subtitle = ParagraphStyle(
            'job_bw_sub', parent=body, fontSize=7, leading=7.5,
            alignment=1, textColor=colors.black
        )

        doc = SimpleDocTemplate(
            path, pagesize=page,
            leftMargin=0.75*72, rightMargin=8*mm,
            topMargin=5*mm, bottomMargin=5*mm
        )

        def P(text, bold=False, size=None):
            st = body_b if bold else body
            if size:
                st = ParagraphStyle(
                    'tmp_%s_%s' % (id(text), size), parent=st,
                    fontSize=size, leading=size+1.2
                )
            return Paragraph(escape(str(text or '-')).replace('\n', '<br/>'), st)

        def bw_table(data, widths, header=False, padd=2.5):
            t = Table(data, colWidths=widths, repeatRows=1 if header else 0,
                      hAlign='LEFT')
            cmds = [
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.45, colors.black),
                ('LEFTPADDING', (0,0), (-1,-1), padd),
                ('RIGHTPADDING', (0,0), (-1,-1), padd),
                ('TOPPADDING', (0,0), (-1,-1), padd),
                ('BOTTOMPADDING', (0,0), (-1,-1), padd),
            ]
            if header:
                cmds += [('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold')]
            t.setStyle(TableStyle(cmds))
            return t

        story = [
            Paragraph('BLUETECH COMPUTERS', title_style),
            Paragraph('PC BUILD JOB SHEET | INTERNAL WORKSHOP COPY', subtitle),
        ]

        # Job/customer details: compact, full-width black-and-white table.
        info = [
            [P('JOB NO', True), P(number), P('QUOTATION NO', True), P(qno),
             P('STATUS', True), P(status_var.get())],
            [P('CUSTOMER', True), P(customer), P('PHONE', True), P(phone),
             P('CREATED', True), P(created)],
            [P('DUE DATE', True), P(due_var.get() or '-'), P('PREPARED BY', True),
             P(staff_vars['Prepared By'].get() or '-'), P(''), P('')],
        ]
        story += [bw_table(info, [22*mm, 62*mm, 28*mm, 55*mm, 22*mm, 73*mm], False),
                  Paragraph('BUILD COMPONENTS', section)]

        pdata = [[P('#', True), P('PRODUCT', True), P('DESCRIPTION', True), P('QTY', True)]]
        for i, (prod, desc, qty) in enumerate(current_items(), 1):
            pdata.append([P(f'{i:02d}'), P(prod), P(desc or '-'), P(qty)])
        story.append(bw_table(pdata, [10*mm, 58*mm, 168*mm, 22*mm], True, padd=1.6))

        story.append(Paragraph('STAFF / RESPONSIBILITY', section))
        sdata = [[P('STAGE', True), P('STAFF', True), P('DATE / TIME', True), P('MANUAL CHECK', True)]]
        for stage in STAGES:
            # Small empty square is intentionally printed for manual marking.
            sdata.append([P(stage), P(staff_vars[stage].get() or '-'),
                          P(time_vars[stage].get() or '-'), P('□')])
        story.append(bw_table(sdata, [65*mm, 75*mm, 70*mm, 48*mm], True, padd=1.7))

        story.append(Paragraph('BUILD / FINAL CHECKLIST', section))
        # Three columns with a small empty square for manual ticking.
        cdata = []
        for i in range(0, len(CHECKS), 3):
            row = []
            for item in CHECKS[i:i+3]:
                row.append(P('□  ' + item))
            while len(row) < 3:
                row.append(P(''))
            cdata.append(row)
        story.append(bw_table(cdata, [86*mm, 86*mm, 86*mm], False, padd=1.7))

        story.append(Paragraph('WORKSHOP REMARKS / SERIAL NUMBERS', section))
        remarks_text = remarks.get('1.0', 'end-1c').strip() or ' '
        # Give the remarks area a useful printable writing space.
        story.append(bw_table([[P(remarks_text)]], [258*mm], False, padd=2))
        story.append(Spacer(1, 2))
        story.append(bw_table([
            [P('WORKSHOP SIGNATURE', True), P(''), P('FINAL APPROVAL', True), P(''), P('DATE', True), P('')]
        ], [38*mm, 78*mm, 32*mm, 78*mm, 18*mm, 14*mm], False, padd=3))

        try:
            doc.build(story)
        except Exception as e:
            messagebox.showerror('PDF', str(e), parent=win)
            return None
        return path

    def pdf_click():
        path=make_pdf()
        if path:
            messagebox.showinfo('PDF Created',path,parent=win)
            try:
                if sys.platform.startswith('win'): os.startfile(path)
                elif sys.platform=='darwin': subprocess.Popen(['open',path])
                else: subprocess.Popen(['xdg-open',path])
            except OSError: pass
    def print_click():
        path=make_pdf()
        if not path: return
        if not sys.platform.startswith('win'):
            messagebox.showinfo('Print',f'Open the PDF and print it:\n{path}',parent=win); return
        if messagebox.askyesno('Print Job Sheet','Send the job sheet to your DEFAULT Windows printer?',parent=win):
            try: os.startfile(path,'print')
            except OSError as e: messagebox.showerror('Printer',f'Printing failed: {e}\nPDF saved at {path}',parent=win)
    def close_window():
        try: win.grab_release()
        except tk.TclError: pass
        win.destroy()
    actions = ttk.Frame(win,padding=12); actions.pack(fill='x')
    ttk.Button(actions,text='SAVE JOB SHEET',command=save).pack(side='left',padx=4)
    ttk.Button(actions,text='SAVE / PREVIEW PDF',command=pdf_click).pack(side='left',padx=4)
    ttk.Button(actions,text='PRINT JOB SHEET',command=print_click).pack(side='left',padx=4)
    ttk.Button(actions,text='CLOSE',command=close_window).pack(side='right',padx=4)


def app_get_users(db):
    con=db(); rows=con.execute('SELECT id,name FROM users WHERE active=1 ORDER BY name COLLATE NOCASE').fetchall(); con.close(); return rows


def show_job_history(app, db, get_pdf_dir):
    setup_db(db)
    win=tk.Toplevel(app.root); win.title('Job Sheet History'); win.geometry('1000x560')
    win.transient(app.root); win.grab_set(); win.focus_force()
    win.configure(bg='#F3F7FC')
    search=tk.StringVar(); ttk.Entry(win,textvariable=search).pack(fill='x',padx=12,pady=8)
    tree=ttk.Treeview(win,columns=('job','quote','customer','status','date'),show='headings')
    for col,label in [('job','Job No.'),('quote','Quotation No.'),('customer','Customer'),('status','Status'),('date','Created')]:
        tree.heading(col,text=label); tree.column(col,width=170)
    tree.pack(fill='both',expand=True,padx=12,pady=8)
    def refresh(*_):
        for item in tree.get_children(): tree.delete(item)
        con=db(); rows=con.execute('SELECT id,job_no,quotation_no,customer,status,created_at FROM job_chits ORDER BY id DESC').fetchall(); con.close()
        term=search.get().lower().strip()
        for jid,job,quote,customer,status,created in rows:
            if term in (' '.join(str(x or '') for x in (job,quote,customer,status))).lower():
                tree.insert('', 'end', iid=str(jid),values=(job,quote,customer,status,created))
    def selected():
        sel=tree.selection()
        if not sel: messagebox.showwarning('Job Sheet','Select a job sheet.',parent=win); return
        open_job_chit(app,db,get_pdf_dir,int(sel[0]))

    def delete_selected():
        sel=tree.selection()
        if not sel:
            messagebox.showwarning('Job Sheet','Select a job sheet to delete.',parent=win)
            return
        jid = int(sel[0])
        job_no = tree.item(sel[0], 'values')[0]
        if not messagebox.askyesno(
            'Delete Job Sheet',
            f'Delete Job Sheet {job_no}?\n\nThis cannot be undone.',
            parent=win
        ):
            return
        con = db()
        try:
            con.execute('DELETE FROM job_chits WHERE id=?', (jid,))
            con.commit()
        finally:
            con.close()
        refresh()
    search.trace_add('write',refresh); refresh()
    ttk.Button(win,text='OPEN / EDIT / PRINT',command=selected).pack(side='left',padx=12,pady=8)
    ttk.Button(win,text='DELETE JOB SHEET',command=delete_selected).pack(side='left',padx=12,pady=8)
    def close_history():
        try: win.grab_release()
        except tk.TclError: pass
        win.destroy()
    ttk.Button(win,text='CLOSE',command=close_history).pack(side='right',padx=12,pady=8)
    tree.bind('<Double-1>',lambda e:selected())
