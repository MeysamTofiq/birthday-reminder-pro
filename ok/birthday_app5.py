import os
import sys
import json
import jdatetime
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from threading import Thread
import time

# برای ساخت شورتکات
try:
    from win32com.client import Dispatch
except ImportError:
    Dispatch = None

# برای نوتیفیکیشن ویندوز در EXE
try:
    from winotify import Notification, audio
    use_winotify = True
except ImportError:
    use_winotify = False

# ====== تنظیمات ======
DATA_FILE = "birthdays.json"
NOTIFY_BEFORE_DAYS = 4
PERSIAN_MONTHS = [
    "فروردین", "اردیبهشت", "خرداد", "تیر",
    "مرداد", "شهریور", "مهر", "آبان",
    "آذر", "دی", "بهمن", "اسفند"
]

# ====== توابع کمکی ======
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_today_jalali():
    greg = datetime.date.today()
    return jdatetime.date.fromgregorian(date=greg)

def normalize(text):
    return text.replace('ي','ی').lower()

def compute_days_left(birth_str):
    today = get_today_jalali()
    try:
        bd = jdatetime.datetime.strptime(birth_str, "%Y-%m-%d").date()
    except:
        return None
    upcoming = jdatetime.date(today.year, bd.month, bd.day)
    delta = (upcoming.togregorian() - today.togregorian()).days
    if delta < 0:
        upcoming = jdatetime.date(today.year + 1, bd.month, bd.day)
        delta = (upcoming.togregorian() - today.togregorian()).days
    return delta

def create_startup_shortcut():
    if Dispatch is None:
        return
    startup = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs\Startup")
    path = os.path.join(startup, "BirthdayApp.lnk")
    if os.path.exists(path):
        return
    try:
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = sys.executable
        shortcut.Arguments = f'"{os.path.abspath(__file__)}" --silent'
        shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
        icon = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.ico')
        if os.path.exists(icon): shortcut.IconLocation = icon
        shortcut.save()
    except:
        pass

def show_notifications(enriched):
    for p in enriched:
        title, msg = None, None
        if p['days_left'] == 0:
            title = "تولد مبارک!"
            msg = f"امروز تولد {p['name']} است."
        elif 0 < p['days_left'] <= NOTIFY_BEFORE_DAYS:
            title = "تولد در راه است"
            msg = f"{p['name']} در {p['days_left']} روز تولدش است."
        if title and msg:
            try:
                if use_winotify:
                    toast = Notification(app_id="BirthdayApp", title=title, msg=msg, duration="short")
                    toast.set_audio(audio.Mail, loop=False)
                    toast.show()
            except:
                pass

class BirthdayApp(tk.Tk):
    def __init__(self, silent=False):
        super().__init__()
        self.silent_mode = silent
        self.title("مدیریت تولدها")
        self.geometry("840x500")
        self.withdraw() if silent else self.deiconify()
        icon = os.path.join(os.path.dirname(__file__), "app.ico")
        if os.path.exists(icon): self.iconbitmap(icon)

        create_startup_shortcut()
        self.enriched_data = []
        self.create_widgets()
        self.load_data()

        # در حالت silent فقط اعلان بده و بعد از چند ثانیه برنامه رو ببند
        if self.silent_mode:
            Thread(target=self.auto_notify_and_quit).start()
        else:
            self.refresh_table()

    def auto_notify_and_quit(self):
        show_notifications(self.enriched_data)
        time.sleep(5)
        self.quit()

    def create_widgets(self):
        if self.silent_mode:
            return
        top = ttk.Frame(self); top.pack(fill="x", padx=10, pady=5)
        ttk.Button(top, text="➕ افزودن", command=self.open_add).pack(side="left", padx=5)
        ttk.Button(top, text="🗑 حذف", command=self.delete_selected).pack(side="left")

        search = ttk.Frame(self); search.pack(fill="x", padx=10)
        ttk.Label(search, text="جستجو:").pack(side="left")
        self.search_var = tk.StringVar()
        e = ttk.Entry(search, textvariable=self.search_var)
        e.pack(side="left", fill="x", expand=True, padx=5)
        e.bind('<KeyRelease>', lambda evt: self.refresh_table())

        frame = ttk.Frame(self); frame.pack(fill="both", expand=True, padx=10, pady=5)
        cols = ("row","name","date","note","days_left")
        heads = ["ردیف","نام","تاریخ تولد","توضیحات","روز مانده"]
        self.tree = ttk.Treeview(frame, columns=cols, show="headings")
        for c,h in zip(cols,heads):
            self.tree.heading(c, text=h)
            w = 50 if c=="row" else 150
            self.tree.column(c, width=w, anchor="center", stretch=True)
        self.tree.tag_configure('upcoming', foreground='red')

        vs = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1); frame.columnconfigure(0, weight=1)

    def load_data(self):
        raw = load_data(); enriched=[]
        for p in raw:
            d=compute_days_left(p.get("date",""))
            enriched.append({**p, "days_left":d if d is not None else -1})
        enriched.sort(key=lambda x: x['days_left'])
        self.enriched_data = enriched

    def refresh_table(self):
        q = normalize(self.search_var.get().strip())
        for i in self.tree.get_children(): self.tree.delete(i)
        for idx,p in enumerate(self.enriched_data,1):
            txt = f"{p['name']} {p['date']} {p.get('note','')} {p['days_left']}"
            if q and q not in normalize(txt): continue
            tag = 'upcoming' if 0 <= p['days_left'] <= NOTIFY_BEFORE_DAYS else ''
            self.tree.insert("",tk.END, values=(idx,p['name'],p['date'],p.get('note',''),p['days_left']), tags=(tag,))

    def open_add(self): AddDialog(self,self.on_add)
    def on_add(self, person):
        d=load_data(); d.append(person); save_data(d)
        self.load_data(); show_notifications(self.enriched_data); self.refresh_table()
    def delete_selected(self):
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("هشدار","ابتدا یک ردیف انتخاب کنید"); return
        vals=self.tree.item(sel[0],"values")
        d=load_data(); d=[p for p in d if not(p['name']==vals[1] and p['date']==vals[2])]
        save_data(d); self.load_data(); self.refresh_table()

class AddDialog(tk.Toplevel):
    def __init__(self,parent,cb):
        super().__init__(parent); self.callback=cb
        self.title("افزودن تولد"); self.resizable(False,False)
        ttk.Label(self,text="نام:").grid(row=0,column=0,padx=5,pady=5,sticky="e")
        self.entry=ttk.Entry(self); self.entry.grid(row=0,column=1,pady=5)
        t=get_today_jalali();
        ttk.Label(self,text="سال:").grid(row=1,column=0,sticky="e")
        self.cb_year=ttk.Combobox(self,width=6,values=list(range(1300,t.year+2)))
        self.cb_year.set(str(t.year)); self.cb_year.grid(row=1,column=1,sticky="w")
        ttk.Label(self,text="ماه:").grid(row=2,column=0,sticky="e")
        self.cb_month=ttk.Combobox(self,values=PERSIAN_MONTHS,state="readonly"); self.cb_month.current(0); self.cb_month.grid(row=2,column=1,sticky="w")
        ttk.Label(self,text="روز:").grid(row=3,column=0,sticky="e")
        self.cb_day=ttk.Combobox(self,width=4,state="readonly"); self.cb_day.grid(row=3,column=1,sticky="w")
        self.cb_year.bind("<<ComboboxSelected>>",self.update_days); self.cb_month.bind("<<ComboboxSelected>>",self.update_days)
        self.update_days();
        ttk.Label(self,text="توضیحات:").grid(row=4,column=0,sticky="ne",pady=5)
        self.txt_note=tk.Text(self,width=30,height=4); self.txt_note.grid(row=4,column=1,pady=5)
        ttk.Button(self,text="ثبت",command=self.on_submit).grid(row=5,column=0,columnspan=2,pady=10)
    def update_days(self,event=None):
        try: y=int(self.cb_year.get()); m=self.cb_month.current()+1
        except: return
        if m<=6: dim=31
        elif m<=11: dim=30
        else: dim=30 if jdatetime.date.isleap(jdatetime.date(y,1,1)) else 29
        self.cb_day['values']=list(range(1,dim+1)); self.cb_day.current(0)
    def on_submit(self):
        name=self.entry.get().strip()
        if not name: messagebox.showerror("خطا","نام را وارد کنید"); return
        y=int(self.cb_year.get()); m=self.cb_month.current()+1; d=int(self.cb_day.get())
        date_str=f"{y:04d}-{m:02d}-{d:02d}"
        note=self.txt_note.get("1.0","end").strip(); self.callback({"name":name,"date":date_str,"note":note}); self.destroy()

if __name__ == "__main__":
    silent = "--silent" in sys.argv
    app = BirthdayApp(silent=silent)
    app.mainloop()
