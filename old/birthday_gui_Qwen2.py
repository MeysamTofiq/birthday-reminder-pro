import os
import sys
import json
import jdatetime
import datetime
import time
import tkinter as tk
from tkinter import ttk, messagebox
from winotify import Notification, audio
import winshell  # برای دسترسی به پوشه استارت‌آپ
from win32com.client import Dispatch

# ====== تنظیمات ======
DATA_FILE = "birthdays.json"
NOTIFY_BEFORE_DAYS = 4
PERSIAN_MONTHS = [
    "فروردین", "اردیبهشت", "خرداد", "تیر",
    "مرداد", "شهریور", "مهر", "آبان",
    "آذر", "دی", "بهمن", "اسفند"
]
ICON_FILE = "app.ico"  # باید در کنار EXE باشد
APP_ID = "BirthdayNotifier"

# ====== توابع کمکی ======
def resource_path(relative_path):
    """پیدا کردن مسیر فایل در حالت EXE"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def create_shortcut():
    try:
        # استفاده از winshell برای دسترسی قطعی به پوشه استارت‌آپ
        startup_folder = winshell.startup()
        shortcut_path = os.path.join(startup_folder, f"{APP_ID}.lnk")

        if not os.path.exists(shortcut_path):
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortcut(shortcut_path)

            if getattr(sys, 'frozen', False):
                # حالت EXE: مسیر اجرایی
                shortcut.TargetPath = sys.executable
                shortcut.WorkingDirectory = os.path.dirname(sys.executable)
            else:
                # حالت Script: مسیر اسکریپت
                script_path = os.path.abspath(sys.argv[0])
                shortcut.TargetPath = sys.executable
                shortcut.Arguments = f'"{script_path}"'
                shortcut.WorkingDirectory = os.path.dirname(script_path)

            # تنظیم آیکون
            icon_path = resource_path(ICON_FILE)
            if os.path.exists(icon_path):
                shortcut.IconLocation = icon_path
            else:
                shortcut.IconLocation = ""  # بدون آیکون

            shortcut.save()
    except Exception as e:
        messagebox.showwarning("هشدار", f"ایجاد شورت‌کات ناموفق:\n{str(e)}")

# ====== نوتیفیکیشن تولدهای پیش‌رو ======
def notify_upcoming():
    today = jdatetime.date.fromgregorian(date=datetime.date.today())
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

# ====== توابع ذخیره و بارگیری داده ======
def load_data():
    return json.load(open(DATA_FILE, "r", encoding="utf-8")) if os.path.exists(DATA_FILE) else []

def save_data(data):
    json.dump(data, open(DATA_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

# ====== کلاس اصلی برنامه ======
class BirthdayApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("مدیریت تولدها")
        icon_path = resource_path(ICON_FILE)
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        self.geometry("600x400")
        self.create_widgets()
        self.load_table()
        self.after(1000, notify_upcoming)

    def create_widgets(self):
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="➕ افزودن", command=self.open_add).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="🗑 حذف", command=self.delete_selected).pack(side="left")

        cols = ("name", "date", "note")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for col, txt in zip(cols, ["نام", "تاریخ تولد", "توضیحات"]):
            self.tree.heading(col, text=txt)
        self.tree.pack(expand=True, fill="both")

    def load_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for p in load_data():
            self.tree.insert("", "end", values=(p["name"], p["date"], p.get("note", "")))

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
        data = [p for p in data if not (p["name"] == vals[0] and p["date"] == vals[1])]
        save_data(data)
        self.load_table()

# ====== دیالوگ افزودن ======
class AddDialog(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        self.title("افزودن تولد")
        self.resizable(False, False)
        
        ttk.Label(self, text="نام:").grid(row=0, column=0, pady=5, padx=5, sticky="e")
        self.entry_name = ttk.Entry(self)
        self.entry_name.grid(row=0, column=1, pady=5)

        today = jdatetime.date.fromgregorian(date=datetime.date.today())
        ttk.Label(self, text="سال:").grid(row=1, column=0, sticky="e")
        self.cb_year = ttk.Combobox(self, width=6, values=list(range(1300, today.year + 1)))
        self.cb_year.set(1330)
        self.cb_year.grid(row=1, column=1, sticky="w")

        ttk.Label(self, text="ماه:").grid(row=2, column=0, sticky="e")
        self.cb_month = ttk.Combobox(self, values=PERSIAN_MONTHS, state="readonly")
        self.cb_month.current(0)
        self.cb_month.grid(row=2, column=1, sticky="w")

        ttk.Label(self, text="روز:").grid(row=3, column=0, sticky="e")
        self.cb_day = ttk.Combobox(self, width=4, state="readonly")
        self.cb_day.grid(row=3, column=1, sticky="w")

        self.cb_year.bind("<<ComboboxSelected>>", self.update_days)
        self.cb_month.bind("<<ComboboxSelected>>", self.update_days)
        self.update_days()

        ttk.Label(self, text="توضیحات:").grid(row=4, column=0, sticky="ne", pady=5)
        self.txt_note = tk.Text(self, width=30, height=4)
        self.txt_note.grid(row=4, column=1, pady=5)

        ttk.Button(self, text="ثبت", command=self.on_submit).grid(row=5, column=0, columnspan=2, pady=10)

    def update_days(self, event=None):
        try:
            y = int(self.cb_year.get())
            m = self.cb_month.current() + 1
        except:
            return
        dim = 31 if m <= 6 else 30 if m <= 11 else 30 if jdatetime.date.isleap(jdatetime.date(y,1,1)) else 29
        self.cb_day['values'] = list(range(1, dim + 1))
        self.cb_day.current(0)

    def on_submit(self):
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showerror("خطا", "نام را وارد کنید")
            return
        y = int(self.cb_year.get())
        m = self.cb_month.current() + 1
        d = int(self.cb_day.get())
        date_str = f"{y:04d}-{m:02d}-{d:02d}"
        note = self.txt_note.get("1.0", "end").strip()
        self.callback({"name": name, "date": date_str, "note": note})
        self.destroy()

# ====== اجرای برنامه ======
if __name__ == "__main__":
    create_shortcut()  # قبل از نمایش GUI اجرا می‌شود
    try:
        app = BirthdayApp()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در اجرا: {e}")