import os
import sys
import json
import jdatetime
import datetime
import time
import warnings
import tkinter as tk
from tkinter import ttk, messagebox, font
from winotify import Notification, audio
import winshell
from win32com.client import Dispatch

# ====== تنظیمات ======
DATA_FILE = "birthdays.json"
NOTIFY_BEFORE_DAYS = 4
PERSIAN_MONTHS = [
    "فروردین", "اردیبهشت", "خرداد", "تیر",
    "مرداد", "شهریور", "مهر", "آبان",
    "آذر", "دی", "بهمن", "اسفند"
]
ICON_FILE = "app.ico"
APP_ID = "BirthdayNotifier"

# ====== توابع کمکی ======
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def create_shortcut():
    try:
        startup_folder = winshell.startup()
        shortcut_path = os.path.join(startup_folder, f"{APP_ID}.lnk")
        if not os.path.exists(shortcut_path):
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortcut(shortcut_path)
            if getattr(sys, 'frozen', False):
                shortcut.TargetPath = sys.executable
                shortcut.WorkingDirectory = os.path.dirname(sys.executable)
            else:
                script_path = os.path.abspath(sys.argv[0])
                shortcut.TargetPath = sys.executable
                shortcut.Arguments = f'"{script_path}"'
                shortcut.WorkingDirectory = os.path.dirname(script_path)
            icon_path = resource_path(ICON_FILE)
            if os.path.exists(icon_path):
                shortcut.IconLocation = icon_path
            shortcut.save()
    except Exception as e:
        messagebox.showwarning("هشدار", f"ایجاد شورت‌کات ناموفق:\n{str(e)}")

def load_data():
    return json.load(open(DATA_FILE, "r", encoding="utf-8")) if os.path.exists(DATA_FILE) else []

def save_data(data):
    json.dump(data, open(DATA_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

def days_in_month(year, month):
    return 31 if month <= 6 else 30 if month <= 11 else 30 if jdatetime.date.isleap(jdatetime.date(year,1,1)) else 29

def get_today_jalali():
    return jdatetime.date.fromgregorian(date=datetime.date.today())

def notify_upcoming():
    today = get_today_jalali()
    birthdays = load_data()
    icon_path = resource_path(ICON_FILE)
    
    for person in birthdays:
        try:
            bd = jdatetime.datetime.strptime(person["date"], "%Y-%m-%d").date()
        except ValueError:
            continue
        
        upcoming = jdatetime.date(today.year, bd.month, bd.day)
        delta = (upcoming.togregorian() - today.togregorian()).days
        if delta < 0:
            upcoming = jdatetime.date(today.year + 1, bd.month, bd.day)
            delta = (upcoming.togregorian() - today.togregorian()).days

        if 0 <= delta <= NOTIFY_BEFORE_DAYS:
            toast = Notification(
                app_id=APP_ID,
                title=f"🎂 تولد نزدیک: {person['name']}",
                msg=f"{delta} روز مانده به تولد\n{person.get('note', '')}",
                icon=icon_path if os.path.exists(icon_path) else None
            )
            toast.set_audio(audio.Default, loop=False)
            toast.show()
            time.sleep(0.5)

# ====== کلاس اصلی برنامه ======
class BirthdayApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("مدیریت تولدها")
        icon_path = resource_path(ICON_FILE)
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        self.geometry("700x500")
        self.font = self.load_custom_font()
        self.create_widgets()
        self.load_table()
        self.after(1000, notify_upcoming)

    def load_custom_font(self):
        font_path = resource_path("Vazirmatn-Bold.ttf")
        try:
            return font.Font(family="Vazirmatn-Bold", size=10, filename=font_path)
        except:
            return font.nametofont("TkDefaultFont")

    def create_widgets(self):
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", pady=5)
    
    # ✅ استفاده از فونت در دکمه‌ها
        ttk.Button(btn_frame, text="➕ افزودن", command=self.open_add, style="Rounded.TButton").pack(side="left", padx=5)
        ttk.Button(btn_frame, text="🗑 حذف", command=self.delete_selected, style="Rounded.TButton").pack(side="left")

        cols = ("name", "date", "note")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
    for col, txt in zip(cols, ["نام", "تاریخ تولد", "توضیحات"]):
        self.tree.heading(col, text=txt, font=self.font)  # ✅ استفاده از فونت
        self.tree.column(col, anchor="center", width=150)
    self.tree.pack(expand=True, fill="both", padx=10, pady=5)
    
    def on_click_link(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            if col == "#4":  # ستون شماره تماس
                item = self.tree.identify_row(event.y)
                values = self.tree.item(item, "values")
                phone = values[3].replace("Tell:", "").strip()
                if phone:
                    webbrowser.open(f"tel:{phone}")

    def load_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for p in load_data():
            self.tree.insert("", "end", values=(
                p["name"],
                p["date"],
                p.get("note", ""),
                f"Tell:{p.get('phone', '')}" if p.get('phone') else ""
            ))

    def open_add(self):
        AddDialog(self, self.on_add)

    def on_add(self, person):
        data = load_data()
        data.append(person)
        save_data(data)
        self.load_table()

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("هشدار", "ابتدا یک ردیف را انتخاب کنید")
            return
        data = load_data()
        vals = self.tree.item(sel[0], "values")
        data = [p for p in data if not (p["name"]==vals[0] and p["date"]==vals[1])]
        save_data(data)
        self.load_table()

# ====== دیالوگ افزودن ======
class AddDialog(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        self.title("افزودن تولد")
        self.resizable(False, False)
        self.font = parent.font
        
        PADX = (10, 20)
        PADY = (5, 5)
        
        # نام
        ttk.Label(self, text="نام:", font=self.font).grid(row=0, column=0, pady=PADY, padx=PADX, sticky="e")
        self.entry_name = ttk.Entry(self, font=self.font)
        self.entry_name.grid(row=0, column=1, pady=PADY, sticky="w")
        
        # سال
        today = get_today_jalali()
        ttk.Label(self, text="سال:", font=self.font).grid(row=1, column=0, pady=PADY, padx=PADX, sticky="e")
        self.cb_year = ttk.Combobox(self, width=6, values=list(range(1300, today.year+1)), font=self.font)
        self.cb_year.set(1330)
        self.cb_year.grid(row=1, column=1, pady=PADY, sticky="w")
        
        # ماه
        ttk.Label(self, text="ماه:", font=self.font).grid(row=2, column=0, pady=PADY, padx=PADX, sticky="e")
        self.cb_month = ttk.Combobox(self, values=PERSIAN_MONTHS, state="readonly", font=self.font)
        self.cb_month.current(0)
        self.cb_month.grid(row=2, column=1, pady=PADY, sticky="w")
        
        # روز
        ttk.Label(self, text="روز:", font=self.font).grid(row=3, column=0, pady=PADY, padx=PADX, sticky="e")
        self.cb_day = ttk.Combobox(self, width=4, state="readonly", font=self.font)
        self.cb_day.grid(row=3, column=1, pady=PADY, sticky="w")
        
        # توضیحات
        ttk.Label(self, text="توضیحات:", font=self.font).grid(row=4, column=0, pady=(PADY[0], 0), padx=PADX, sticky="e")
        self.txt_note = tk.Text(self, width=30, height=4, font=self.font)
        self.txt_note.grid(row=4, column=1, rowspan=2, pady=(0, PADY[1]), sticky="w")
        
        # شماره تماس
        ttk.Label(self, text="شماره تماس:", font=self.font).grid(row=6, column=0, pady=PADY, padx=PADX, sticky="e")
        self.entry_phone = ttk.Entry(self, font=self.font)
        self.entry_phone.grid(row=6, column=1, pady=PADY, sticky="w")
        
        # ذخیره
        ttk.Button(self, text="ثبت", command=self.on_submit, style="Rounded.TButton").grid(row=7, column=0, columnspan=2, pady=15, sticky="ew")
        
        # بایندینگ
        self.cb_year.bind("<<ComboboxSelected>>", self.update_days)
        self.cb_month.bind("<<ComboboxSelected>>", self.update_days)
        self.update_days()
        self.entry_name.focus()

    def update_days(self, event=None):
        try:
            y = int(self.cb_year.get())
            m = self.cb_month.current()+1
        except:
            return
        dim = days_in_month(y, m)
        current_day = int(self.cb_day.get()) if self.cb_day.get() else 1
        if 1 <= current_day <= dim:
            self.cb_day['values'] = list(range(1, dim+1))
            self.cb_day.set(current_day)
        else:
            self.cb_day['values'] = list(range(1, dim+1))
            self.cb_day.set(1)

    def on_submit(self):
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showerror("خطا","نام را وارد کنید")
            return
        y = int(self.cb_year.get())
        m = self.cb_month.current()+1
        d = int(self.cb_day.get())
        date_str = f"{y:04d}-{m:02d}-{d:02d}"
        note = self.txt_note.get("1.0","end").strip()
        phone = self.entry_phone.get().strip()
        self.callback({
            "name": name,
            "date": date_str,
            "note": note,
            "phone": phone
        })
        self.destroy()

# ====== اجرای برنامه ======
if __name__ == "__main__":
    try:
        create_shortcut()
        app = BirthdayApp()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در اجرا: {e}")
