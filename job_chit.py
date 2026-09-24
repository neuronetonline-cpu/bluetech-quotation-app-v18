"""Bluetech PC-build job chit window; bundled alongside quotation_app.py."""
import json
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from xml.sax.saxutils import escape
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, KeepTogether, KeepInFrame
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
            messagebox.showerror('Job Chit', 'Job chit not found.'); return
    else:
        if not app.customer.get().strip():
            messagebox.showwarning('Job Chit', 'Enter the customer name first.'); return
        items = app.collect_items()
        if not items:
            messagebox.showwarning('Job Chit', 'Add at least one product first.'); return
        # Persist quotation before creating its workshop document.
        qid = app.save_quote(show_message=False)
        if qid is None: return
        con = db(); saved = con.execute('SELECT * FROM job_chits WHERE quotation_id=? ORDER BY id DESC LIMIT 1', (qid,)).fetchone(); con.close()

    win = tk.Toplevel(app.root); win.title('Bluetech Computers - PC Build Job Chit')
    win.geometry('1020x780'); win.minsize(780, 550)
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

    ttk.Label(body, text='PC BUILD JOB CHIT', font=('Segoe UI', 19, 'bold'), foreground='#075EAA').pack(anchor='w')
    ttk.Label(body, text='Internal workshop document — costs and selling prices are excluded.', foreground='#667085').pack(anchor='w', pady=(0,12))
    info = ttk.LabelFrame(body, text='JOB / CUSTOMER DETAILS', padding=10); info.pack(fill='x', pady=5)
    def info_row(r, label, val):
        ttk.Label(info, text=label, font=('Segoe UI',9,'bold')).grid(row=r, column=0, sticky='w', pady=4, padx=5)
        ttk.Label(info, text=val).grid(row=r, column=1, sticky='w', pady=4, padx=5)
    info_row(0, 'Job Chit No.', number); info_row(1, 'Quotation No.', qno)
    info_row(2, 'Customer', customer); info_row(3, 'Phone', phone); info_row(4, 'Created', created)
    ttk.Label(info, text='Due Date (YYYY-MM-DD)').grid(row=5, column=0, sticky='w', padx=5)
    due_var = tk.StringVar(value=due or ''); ttk.Entry(info, textvariable=due_var, width=24).grid(row=5,column=1,sticky='w',padx=5,pady=5)
    ttk.Label(info, text='Status').grid(row=6,column=0,sticky='w',padx=5)
    status_var = tk.StringVar(value=status or 'Approved')
    ttk.Combobox(info, textvariable=status_var, values=STATUSES, state='readonly', width=22).grid(row=6,column=1,sticky='w',padx=5,pady=5)

    parts = ttk.LabelFrame(body, text='BUILD COMPONENTS (FROM QUOTATION)', padding=10); parts.pack(fill='x', pady=7)
    for col, title in enumerate(['PRODUCT', 'DESCRIPTION', 'QTY']):
        ttk.Label(parts, text=title, font=('Segoe UI',9,'bold')).grid(row=0,column=col,sticky='w',padx=5)

    # Product descriptions are editable here before the Job Chit is saved/printed.
    description_vars = []
    for i, (p,d,q) in enumerate(items, 1):
        ttk.Label(parts, text=str(p), font=('Segoe UI',9,'bold')).grid(row=i,column=0,sticky='w',padx=5,pady=3)
        desc_var = tk.StringVar(value=str(d or ''))
        description_vars.append(desc_var)
        ttk.Entry(parts, textvariable=desc_var, width=65).grid(row=i,column=1,sticky='ew',padx=5,pady=2)
        ttk.Label(parts, text=str(q)).grid(row=i,column=2,sticky='w',padx=5)
    parts.columnconfigure(1, weight=1)

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

    def payload():
        # Read the edited descriptions from the Job Chit window before saving.
        for idx, desc_var in enumerate(description_vars):
            if idx < len(items):
                items[idx][1] = desc_var.get().strip()
        return (due_var.get().strip(), status_var.get(), json.dumps(items,ensure_ascii=False),
                json.dumps({k:v.get() for k,v in check_vars.items()}),
                json.dumps({k:v.get().strip() for k,v in staff_vars.items()}),
                json.dumps({k:v.get() for k,v in time_vars.items()}), remarks.get('1.0','end-1c').strip())
    def save(silent=False):
        nonlocal jid
        due, status, item_json, checks_json, staff_json, stamps_json, notes = payload()
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
            messagebox.showerror('Job Chit',str(e),parent=win); return False
        finally: con.close()
        if not silent: messagebox.showinfo('Saved',f'{number} saved.',parent=win)
        return True

    def make_pdf():
        if not save(silent=True): return None
        folder = os.path.join(get_pdf_dir(), 'Job Chits'); os.makedirs(folder,exist_ok=True)
        path = os.path.join(folder,number+'.pdf')
        styles = getSampleStyleSheet()
        small = ParagraphStyle('jobsmall', parent=styles['Normal'], fontSize=7.2, leading=8.7, spaceAfter=0)
        tiny = ParagraphStyle('jobtiny', parent=small, fontSize=6.8, leading=8.0)
        title = ParagraphStyle('jobtitle',parent=styles['Title'],fontSize=15,leading=17,textColor=colors.HexColor('#075EAA'),spaceAfter=0)
        section = ParagraphStyle('jobsection',parent=small,fontSize=7.4,leading=9,fontName='Helvetica-Bold',textColor=colors.HexColor('#075EAA'),spaceAfter=2)

        doc = SimpleDocTemplate(
            path, pagesize=A4,
            leftMargin=9*mm, rightMargin=9*mm, topMargin=7*mm, bottomMargin=7*mm
        )

        def para(s, style=small):
            return Paragraph(escape(str(s or '-')).replace('\n','<br/>'), style)

        def table(data, widths, header=False, font_size=7.2, pad=2.5):
            t = Table(data, colWidths=widths, hAlign='LEFT', repeatRows=1 if header else 0)
            commands = [
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('GRID',(0,0),(-1,-1),.35,colors.HexColor('#B8C7D9')),
                ('LEFTPADDING',(0,0),(-1,-1),pad),
                ('RIGHTPADDING',(0,0),(-1,-1),pad),
                ('TOPPADDING',(0,0),(-1,-1),pad),
                ('BOTTOMPADDING',(0,0),(-1,-1),pad),
            ]
            if header:
                commands += [('BACKGROUND',(0,0),(-1,0),colors.HexColor('#DCEEFF'))]
            t.setStyle(TableStyle(commands)); return t

        story = [
            Paragraph('BLUETECH COMPUTERS', title),
            Paragraph('PC BUILD JOB CHIT  |  INTERNAL WORKSHOP COPY', tiny),
            Spacer(1,3),
        ]

        info_data = [
            [para('<b>JOB NO:</b> '+number), para('<b>QUOTATION:</b> '+str(qno)), para('<b>STATUS:</b> '+status_var.get())],
            [para('<b>CUSTOMER:</b> '+str(customer)), para('<b>PHONE:</b> '+str(phone)), para('<b>CREATED:</b> '+str(created))],
            [para('<b>DUE:</b> '+str(due_var.get() or '-')), para('<b>PREPARED BY:</b> '+str(staff_vars.get('Prepared By',tk.StringVar()).get())), para('')],
        ]
        story += [table(info_data,[64*mm,64*mm,54*mm]), Spacer(1,3), Paragraph('BUILD COMPONENTS',section)]

        pdata = [[para('<b>PRODUCT</b>'), para('<b>DESCRIPTION</b>'), para('<b>QTY</b>')]]
        for p,d,q in items:
            pdata.append([para(p), para(d or '-'), para(q)])
        story += [table(pdata,[48*mm,115*mm,19*mm],header=True,pad=2), Spacer(1,3), Paragraph('STAFF / RESPONSIBILITY',section)]

        sdata = [[para('<b>STAGE</b>'),para('<b>STAFF</b>'),para('<b>DATE / TIME</b>')]]
        for stage in STAGES:
            sdata.append([para(stage),para(staff_vars[stage].get()),para(time_vars[stage].get())])
        story += [table(sdata,[58*mm,62*mm,62*mm],header=True,pad=2), Spacer(1,3), Paragraph('BUILD / FINAL CHECKLIST',section)]

        # Compact 3-column checklist keeps the whole Job Chit on one A4 page.
        cdata = []
        for start_idx in range(0, len(CHECKS), 3):
            row = []
            for item in CHECKS[start_idx:start_idx+3]:
                mark = '✓' if check_vars[item].get() else '☐'
                row.append(para(mark+' '+item, tiny))
            while len(row) < 3:
                row.append(para('', tiny))
            cdata.append(row)
        story += [table(cdata,[60*mm,60*mm,62*mm],pad=2), Spacer(1,3), Paragraph('REMARKS / SERIAL NUMBERS',section),
                  table([[para(remarks.get('1.0','end-1c') or '-')],[para('Workshop signature: ____________________    Final approval: ____________________',tiny)]],[182*mm],pad=2)]

        # Shrink the compact layout only if unusually long descriptions/remarks would
        # otherwise push the document onto a second page.
        available_h = A4[1] - (16*mm)
        story = [KeepInFrame(A4[0] - 18*mm, available_h, story, mode='shrink')]
        try:
            doc.build(story)
        except Exception as e:
            messagebox.showerror('PDF',str(e),parent=win); return None
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
        if messagebox.askyesno('Print Job Chit','Send the job chit to your DEFAULT Windows printer?',parent=win):
            try:
                os.startfile(path, 'print')
            except OSError as e:
                # Some Windows PDF viewers do not register a 'print' shell verb
                # (WinError 1155). Open the generated PDF instead so the user
                # can print it normally with Ctrl+P / the PDF viewer's Print button.
                if getattr(e, 'winerror', None) == 1155 or getattr(e, 'errno', None) == 1155:
                    try:
                        os.startfile(path, 'open')
                        messagebox.showinfo(
                            'Print',
                            'Windows does not have a direct PDF print action configured.\n\n'
                            'The Job Chit PDF has been opened. Use Ctrl+P or the Print button in the PDF viewer.',
                            parent=win
                        )
                    except OSError:
                        messagebox.showerror('Printer', f'Printing failed.\nPDF saved at {path}', parent=win)
                else:
                    messagebox.showerror('Printer', f'Printing failed: {e}\nPDF saved at {path}', parent=win)
    actions = ttk.Frame(win,padding=12); actions.pack(fill='x')
    ttk.Button(actions,text='SAVE JOB CHIT',command=save).pack(side='left',padx=4)
    ttk.Button(actions,text='SAVE / PREVIEW PDF',command=pdf_click).pack(side='left',padx=4)
    ttk.Button(actions,text='PRINT JOB CHIT',command=print_click).pack(side='left',padx=4)
    ttk.Button(actions,text='CLOSE',command=win.destroy).pack(side='right',padx=4)


def app_get_users(db):
    con=db(); rows=con.execute('SELECT id,name FROM users WHERE active=1 ORDER BY name COLLATE NOCASE').fetchall(); con.close(); return rows


def show_job_history(app, db, get_pdf_dir):
    setup_db(db)
    win=tk.Toplevel(app.root); win.title('Job Chit History'); win.geometry('1000x560')
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
        if not sel: messagebox.showwarning('Job Chit','Select a job chit.',parent=win); return
        open_job_chit(app,db,get_pdf_dir,int(sel[0]))
    search.trace_add('write',refresh); refresh()
    ttk.Button(win,text='OPEN / EDIT / PRINT',command=selected).pack(side='left',padx=12,pady=8)
    ttk.Button(win,text='CLOSE',command=win.destroy).pack(side='right',padx=12,pady=8)
    tree.bind('<Double-1>',lambda e:selected())
