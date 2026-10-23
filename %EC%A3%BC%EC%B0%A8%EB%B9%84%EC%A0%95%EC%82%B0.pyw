# -*- coding: utf-8 -*-
"""병원 주차비 현금 정산 (Windows / Python standard library only)."""
import calendar
import datetime as dt
import math
import os
import re
import shutil
import sqlite3
import sys
import zipfile
from pathlib import Path
from tkinter import Tk, StringVar, ttk, messagebox, filedialog
from tkinter import font as tkfont
from xml.sax.saxutils import escape

APP_NAME = '병원 주차비 정산'
BASE = Path(os.getenv('LOCALAPPDATA') or Path.home()) / 'HospitalParkingCash'
DB_PATH = BASE / 'parking.sqlite3'
BACKUP_DIR = BASE / 'backups'
TIME_FMT = '%Y-%m-%d %H:%M:%S'


def now():
    return dt.datetime.now().replace(microsecond=0)


def connect():
    BASE.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(str(DB_PATH), timeout=10)
    db.row_factory = sqlite3.Row
    db.execute('PRAGMA busy_timeout=10000')
    db.execute('''CREATE TABLE IF NOT EXISTS payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        paid_at TEXT NOT NULL,
        arrival_at TEXT NOT NULL,
        recorded_departure TEXT NOT NULL,
        base_minutes INTEGER NOT NULL,
        grace_minutes INTEGER NOT NULL,
        billable_minutes INTEGER NOT NULL,
        chart_number TEXT NOT NULL,
        amount INTEGER NOT NULL,
        status TEXT NOT NULL DEFAULT '지급완료',
        cancelled_at TEXT,
        cancel_reason TEXT
    )''')
    db.execute('CREATE INDEX IF NOT EXISTS idx_paid_at ON payments(paid_at)')
    db.commit()
    return db


def get_arrival(txt, clock):
    t = txt.strip().replace('：', ':')
    if not re.fullmatch(r'\d{1,2}:\d{2}', t):
        raise ValueError('입차시간은 09:05처럼 입력해 주세요.')
    h, m = [int(v) for v in t.split(':')]
    if h > 23 or m > 59:
        raise ValueError('입차시간이 올바르지 않습니다.')
    arrival = clock.replace(hour=h, minute=m, second=0)
    if arrival > clock:
        arrival -= dt.timedelta(days=1)
    if (clock-arrival).total_seconds() >= 24*3600:
        raise ValueError('입차시간을 확인해 주세요.')
    return arrival


def calculate(arrival, clock):
    elapsed = max(0, math.ceil((clock-arrival).total_seconds() / 60))
    billed = elapsed + 20
    amount = 2000 + max(0, math.ceil((billed-30)/10))*1000
    return elapsed, billed, amount


def won(n):
    return f'{n:,}원'


def backup_db(db, force=False):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    dest = BACKUP_DIR / (dt.datetime.now().strftime('%Y%m%d') + '.sqlite3')
    if force or not dest.exists():
        with sqlite3.connect(str(dest)) as target:
            db.backup(target)
    return dest


def column_name(n):
    name = ''
    while n:
        n, r = divmod(n-1, 26)
        name = chr(65+r) + name
    return name


def sheet_xml(rows, col_widths, monetary_indices=()):
    """Generate valid simple XLSX sheet XML; leading chart-number zeros retained as text."""
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
             '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
             '<cols>']
    for i, width in enumerate(col_widths, 1):
        parts.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
    parts.append('</cols><sheetData>')
    for rnum, row in enumerate(rows, 1):
        parts.append(f'<row r="{rnum}">')
        for cnum, val in enumerate(row, 1):
            if val is None:
                continue
            addr = column_name(cnum) + str(rnum)
            style = 1 if rnum == 1 else 2 if cnum-1 in monetary_indices else 0
            if isinstance(val, (int, float)):
                parts.append(f'<c r="{addr}" s="{style}"><v>{val}</v></c>')
            else:
                s = str(val)
                # Excel formula injection protection and preserve strings like chart numbers.
                if s.startswith(('=', '+', '-', '@')):
                    s = "'" + s
                parts.append(f'<c r="{addr}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(s)}</t></is></c>')
        parts.append('</row>')
    parts.append('</sheetData></worksheet>')
    return ''.join(parts)


def save_xlsx(path, details, summary):
    names = ['지급내역', '일별집계']
    xmls = [sheet_xml(details, [10,22,19,19,19,13,14,15,13,22,30], (8,)),
            sheet_xml(summary, [16,16,18,15,19], (2, 3, 4))]
    content_types = '''<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'''
    wb = '''<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="지급내역" sheetId="1" r:id="rId1"/><sheet name="일별집계" sheetId="2" r:id="rId2"/></sheets></workbook>'''
    wbrels = '''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
    styles = '''<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="맑은 고딕"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF19365A"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>'''
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in [('[Content_Types].xml', content_types),('_rels/.rels', rels),
                           ('xl/workbook.xml',wb),('xl/_rels/workbook.xml.rels',wbrels),('xl/styles.xml',styles),
                           ('xl/worksheets/sheet1.xml',xmls[0]),('xl/worksheets/sheet2.xml',xmls[1])]:
            z.writestr(name, text)


class ParkingApp:
    def __init__(self, root):
        self.root = root
        self.db = connect()
        backup_db(self.db)
        self.root.title('병원 주차비 정산 | 현금 지급관리')
        self.root.geometry('1130x790')
        self.root.minsize(940, 650)
        self.style = ttk.Style()
        try: self.style.theme_use('clam')
        except Exception: pass
        self.style.configure('TButton', padding=8, font=('맑은 고딕', 10))
        self.style.configure('TLabel', font=('맑은 고딕', 10))
        self.style.configure('Title.TLabel', font=('맑은 고딕', 19, 'bold'), foreground='#17385d')
        self.style.configure('Amount.TLabel', font=('맑은 고딕', 23, 'bold'), foreground='#075b93')
        self.style.configure('TLabelframe.Label', font=('맑은 고딕', 11, 'bold'))
        self.style.configure('Treeview', rowheight=28, font=('맑은 고딕', 9))
        self.style.configure('Treeview.Heading', font=('맑은 고딕', 9, 'bold'))
        self.arrival_var = StringVar()
        self.chart_var = StringVar()
        self.clock_var = StringVar()
        self.result_var = StringVar(value='입차시간을 입력하면 계산됩니다.')
        self.amount_var = StringVar(value='— 원')
        self.mode = StringVar(value='일별')
        self.date_var = StringVar(value=now().strftime('%Y-%m-%d'))
        self.month_var = StringVar(value=now().strftime('%Y-%m'))
        self.summary_var = StringVar()
        self.preview = None
        self.setup_ui()
        self.refresh_records()
        self.tick()
        self.root.protocol('WM_DELETE_WINDOW', self.close)

    def setup_ui(self):
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='주차비 현금 지급 관리', style='Title.TLabel').pack(anchor='w')
        ttk.Label(frame, text='최초 30분 2,000원  ·  이후 10분당 1,000원  ·  귀가 여유시간 +20분  ·  전 요일 동일요금').pack(anchor='w', pady=(4,13))
        calc = ttk.LabelFrame(frame, text=' 1. 주차비 계산 및 지급 등록 ', padding=14)
        calc.pack(fill='x')
        first = ttk.Frame(calc)
        first.pack(fill='x')
        ttk.Label(first, text='입차시간 (시:분)').grid(row=0, column=0, sticky='w')
        self.arr_entry = ttk.Entry(first, textvariable=self.arrival_var, width=15, font=('맑은 고딕', 15))
        self.arr_entry.grid(row=1, column=0, sticky='w', padx=(0,22), pady=7)
        ttk.Label(first, text='환자 차트번호 (필수)').grid(row=0, column=1, sticky='w')
        self.chart_entry = ttk.Entry(first, textvariable=self.chart_var, width=23, font=('맑은 고딕', 15))
        self.chart_entry.grid(row=1, column=1, sticky='w', padx=(0,22), pady=7)
        ttk.Label(first, textvariable=self.clock_var).grid(row=0, column=2, columnspan=2, sticky='w')
        ttk.Button(first, text='주차비 계산', command=self.preview_amount).grid(row=1, column=2, padx=(0,9))
        ttk.Button(first, text='지급완료 및 기록 저장', command=self.pay).grid(row=1, column=3)
        ttk.Label(calc, textvariable=self.result_var).pack(anchor='w', pady=(9,0))
        ttk.Label(calc, textvariable=self.amount_var, style='Amount.TLabel').pack(anchor='w', pady=(2,0))
        ttk.Label(calc, text='※ 표시된 금액은 계산 시점 기준입니다. 지급완료 시 현재 시간을 다시 확인합니다.').pack(anchor='w', pady=(2,0))
        self.arrival_var.trace_add('write', self.reset_preview)
        self.chart_var.trace_add('write', self.reset_preview)
        hist = ttk.LabelFrame(frame, text=' 2. 지급내역 / 정산 조회 ', padding=12)
        hist.pack(fill='both', expand=True, pady=(15,0))
        bar = ttk.Frame(hist)
        bar.pack(fill='x')
        ttk.Radiobutton(bar, text='일별', value='일별', variable=self.mode, command=self.refresh_records).pack(side='left')
        ttk.Entry(bar, textvariable=self.date_var, width=13).pack(side='left', padx=(5,15))
        ttk.Radiobutton(bar, text='월별', value='월별', variable=self.mode, command=self.refresh_records).pack(side='left')
        ttk.Entry(bar, textvariable=self.month_var, width=10).pack(side='left', padx=(5,12))
        ttk.Button(bar, text='조회', command=self.refresh_records).pack(side='left', padx=(0,12))
        ttk.Button(bar, text='엑셀 다운로드 (.xlsx)', command=self.export).pack(side='right')
        ttk.Label(hist, textvariable=self.summary_var, font=('맑은 고딕', 12, 'bold')).pack(anchor='w', pady=(12,8))
        table_wrap = ttk.Frame(hist)
        table_wrap.pack(fill='both', expand=True)
        columns = ('id','paid_at','chart','arrival','duration','amount','status','cancelled')
        self.tree = ttk.Treeview(table_wrap, columns=columns, show='headings', selectmode='browse')
        heads = {'id':'번호','paid_at':'지급일시','chart':'차트번호','arrival':'입차시간','duration':'인정 주차시간','amount':'지급금액','status':'상태','cancelled':'취소일시'}
        widths = {'id':54,'paid_at':150,'chart':110,'arrival':140,'duration':114,'amount':105,'status':80,'cancelled':150}
        for c in columns:
            self.tree.heading(c, text=heads[c]); self.tree.column(c, width=widths[c], minwidth=60, anchor='center' if c not in ('chart',) else 'w')
        sy = ttk.Scrollbar(table_wrap, orient='vertical', command=self.tree.yview)
        sx = ttk.Scrollbar(table_wrap, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0,column=0,sticky='nsew'); sy.grid(row=0,column=1,sticky='ns'); sx.grid(row=1,column=0,sticky='ew')
        table_wrap.grid_columnconfigure(0,weight=1); table_wrap.grid_rowconfigure(0,weight=1)
        self.tree.tag_configure('cancelled', foreground='#9299a4')
        bottom=ttk.Frame(hist); bottom.pack(fill='x',pady=(10,0))
        ttk.Button(bottom, text='선택 지급건 취소', command=self.cancel).pack(side='left')
        ttk.Button(bottom, text='지금 데이터 백업', command=self.manual_backup).pack(side='left', padx=8)
        ttk.Label(bottom, text=f'기록 저장위치: {DB_PATH}', foreground='#666666').pack(side='right')
        self.arr_entry.focus_set()

    def reset_preview(self, *_):
        self.preview = None
        self.amount_var.set('— 원')
        self.result_var.set('입차시간을 입력하고 주차비 계산을 눌러주세요.')

    def tick(self):
        self.clock_var.set('현재시간  ' + now().strftime('%Y-%m-%d %H:%M:%S'))
        self.root.after(1000, self.tick)

    def preview_amount(self):
        clock = now()
        try:
            arrival = get_arrival(self.arrival_var.get(), clock)
            elapsed, billed, amount = calculate(arrival, clock)
        except ValueError as ex:
            messagebox.showwarning('입차시간 확인', str(ex)); return None
        if arrival.date() != clock.date():
            if not messagebox.askyesno('전날 입차 확인', f'입차 시각을 전날({arrival:%Y-%m-%d %H:%M})로 해석했습니다. 맞습니까?'):
                return None
        self.preview = (arrival, clock, elapsed, billed, amount)
        self.result_var.set(f'입차 {arrival:%m/%d %H:%M}  ·  실제 {elapsed}분  +  여유 20분  =  인정 {billed}분')
        self.amount_var.set(won(amount))
        return self.preview

    def pay(self):
        chart = self.chart_var.get().strip()
        if not chart:
            messagebox.showwarning('차트번호 필수', '차트번호를 입력한 뒤 지급완료를 눌러주세요.'); self.chart_entry.focus_set(); return
        if len(chart)>50 or not re.fullmatch(r'[0-9A-Za-z가-힣_-]+',chart):
            messagebox.showwarning('차트번호 확인','차트번호는 50자 이하의 숫자·영문·한글·하이픈·밑줄만 입력해 주세요.'); return
        if self.preview is None:
            if self.preview_amount() is None: return
        clock = now()
        try:
            arrival = get_arrival(self.arrival_var.get(),clock)
            elapsed,billed,amount = calculate(arrival,clock)
        except ValueError as ex:
            messagebox.showwarning('입차시간 확인',str(ex)); return
        if not messagebox.askyesno('현금 지급 확인', f'차트번호: {chart}\n입차: {arrival:%Y-%m-%d %H:%M}\n인정 주차시간: {billed}분\n\n현금 {won(amount)}을 지급 완료했습니까?'):
            return
        try:
            with self.db:
                self.db.execute('''INSERT INTO payments (paid_at,arrival_at,recorded_departure,base_minutes,grace_minutes,billable_minutes,chart_number,amount)
                                   VALUES (?,?,?,?,?,?,?,?)''', (clock.strftime(TIME_FMT),arrival.strftime(TIME_FMT),clock.strftime(TIME_FMT),elapsed,20,billed,chart,amount))
            backup_db(self.db, force=True)
        except Exception as ex:
            messagebox.showerror('저장 또는 백업 오류',f'기록 또는 백업에 문제가 발생했습니다. 관리자 확인이 필요합니다.\n{ex}')
            self.refresh_records(); return
        messagebox.showinfo('저장 완료',f'{chart} / {won(amount)} 지급내역이 저장되었습니다.')
        self.arrival_var.set(''); self.chart_var.set(''); self.refresh_records(); self.arr_entry.focus_set()

    def period(self):
        if self.mode.get() == '일별':
            date = dt.datetime.strptime(self.date_var.get().strip(),'%Y-%m-%d').date()
            return date.isoformat(), (date+dt.timedelta(days=1)).isoformat(),date.isoformat()
        text = self.month_var.get().strip()
        if not re.fullmatch(r'\d{4}-\d{2}',text): raise ValueError('월은 YYYY-MM 형식으로 입력해 주세요.')
        year,month = map(int,text.split('-'))
        start=dt.date(year,month,1)
        end = dt.date(year+1,1,1) if month==12 else dt.date(year,month+1,1)
        return start.isoformat(),end.isoformat(),text

    def records(self):
        start,end,label = self.period()
        items=self.db.execute('SELECT * FROM payments WHERE paid_at>=? AND paid_at<? ORDER BY paid_at DESC, id DESC',(start,end)).fetchall()
        return items,label

    def refresh_records(self):
        try: items,label = self.records()
        except ValueError as ex:
            messagebox.showwarning('조회 날짜 확인',str(ex)); return
        for it in self.tree.get_children(): self.tree.delete(it)
        active=[r for r in items if r['status']=='지급완료']
        total=sum(r['amount'] for r in active)
        self.summary_var.set(f'{label}  |  지급 {len(active)}건  |  총 지급액 {won(total)}  |  취소 {len(items)-len(active)}건')
        for r in items:
            self.tree.insert('', 'end', iid=str(r['id']), tags=('cancelled',) if r['status']!='지급완료' else (),
                values=(r['id'],r['paid_at'],r['chart_number'],r['arrival_at'][:16],f"{r['billable_minutes']}분",won(r['amount']),r['status'],r['cancelled_at'] or ''))

    def cancel(self):
        selected=self.tree.selection()
        if not selected:
            messagebox.showwarning('지급건 선택','취소할 지급건을 목록에서 선택해 주세요.'); return
        id_ = int(selected[0]); record=self.db.execute('SELECT * FROM payments WHERE id=?',(id_,)).fetchone()
        if not record or record['status']!='지급완료':
            messagebox.showwarning('취소 불가','이미 취소되었거나 존재하지 않는 기록입니다.'); return
        if not messagebox.askyesno('지급 취소 확인',f"차트번호: {record['chart_number']}\n지급금액: {won(record['amount'])}\n\n이 지급건을 취소 처리할까요?\n(기록은 남고 합계에서 제외됩니다.)"):
            return
        with self.db:
            cur=self.db.execute("UPDATE payments SET status='취소', cancelled_at=?, cancel_reason='직원 취소' WHERE id=? AND status='지급완료'",(now().strftime(TIME_FMT),id_))
        if cur.rowcount:
            try: backup_db(self.db,force=True)
            except Exception as ex: messagebox.showwarning('백업 오류',f'취소 기록은 저장됐지만 백업에 실패했습니다.\n{ex}')
            self.refresh_records()
        else: messagebox.showerror('오류','지급 취소를 완료하지 못했습니다.')

    def export(self):
        try: items,label=self.records()
        except ValueError as ex: messagebox.showwarning('조회 날짜 확인',str(ex)); return
        default=f'주차비정산_{label}.xlsx'
        dest=filedialog.asksaveasfilename(title='엑셀 정산자료 저장',initialfile=default,defaultextension='.xlsx',filetypes=[('엑셀 파일','*.xlsx')])
        if not dest: return
        details=[['번호','지급일시','차트번호','입차일시','정산시각','실제 주차(분)','추가 여유(분)','인정 주차(분)','지급액(원)','상태','취소일시']]
        for r in items:
            details.append([r['id'],r['paid_at'],r['chart_number'],r['arrival_at'],r['recorded_departure'],r['base_minutes'],r['grace_minutes'],r['billable_minutes'],r['amount'],r['status'],r['cancelled_at'] or ''])
        # Include 0-activity days within the selected window, supporting cash reconciliation.
        start,end,_=self.period()
        first=dt.date.fromisoformat(start); last=dt.date.fromisoformat(end)
        agg={}
        cur=first
        while cur<last:
            agg[cur.isoformat()]=[0,0,0,0]
            cur+=dt.timedelta(days=1)
        for r in items:
            key=r['paid_at'][:10]
            if r['status']=='지급완료':agg[key][0]+=1;agg[key][1]+=r['amount']
            else:agg[key][2]+=1;agg[key][3]+=r['amount']
        summary=[['날짜','지급건수','실제 지급액(원)','취소건수','취소된 원금(원)']]
        for day,vals in sorted(agg.items()):summary.append([day,vals[0],vals[1],vals[2],vals[3]])
        summary.append(['합계',sum(v[0] for v in agg.values()),sum(v[1] for v in agg.values()),sum(v[2] for v in agg.values()),sum(v[3] for v in agg.values())])
        try:
            save_xlsx(dest, details, summary)
            messagebox.showinfo('엑셀 저장 완료',f'정산자료가 저장되었습니다.\n{dest}\n\n※ 내부 지급대장으로 활용하되 세무 증빙 인정 여부는 세무대리인에게 확인해 주세요.')
        except Exception as ex:messagebox.showerror('엑셀 저장 실패',str(ex))

    def manual_backup(self):
        folder=filedialog.askdirectory(title='백업 저장 폴더 선택')
        if not folder:return
        dest=Path(folder)/('주차비정산_백업_'+now().strftime('%Y%m%d_%H%M%S')+'.sqlite3')
        try:
            with sqlite3.connect(str(dest)) as dbcopy:self.db.backup(dbcopy)
            messagebox.showinfo('백업 완료',f'백업 파일이 저장되었습니다.\n{dest}')
        except Exception as ex:messagebox.showerror('백업 오류',str(ex))

    def close(self):
        try: backup_db(self.db,force=True)
        except Exception as ex:messagebox.showwarning('백업 오류',f'종료 전 백업에 실패했습니다.\n{ex}')
        self.db.close(); self.root.destroy()


if __name__ == '__main__':
    root=Tk()
    try: ParkingApp(root); root.mainloop()
    except Exception as ex:
        messagebox.showerror('프로그램 오류',str(ex))
        raise
