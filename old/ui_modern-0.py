import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk

# منطق اصلی
import birthday_app6 as core


ctk.set_appearance_mode("dark")  # "dark" or "light"
ctk.set_default_color_theme("blue")


class AddDialog(ctk.CTkToplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback = callback
        self.title("افزودن تولد")
        self.resizable(False, False)
        self.geometry("360x360")

        ctk.CTkLabel(self, text="نام:").grid(row=0, column=0, padx=10, pady=8, sticky="e")
        self.entry = ctk.CTkEntry(self, width=220)
        self.entry.grid(row=0, column=1, pady=8, sticky="w")

        t = core.get_today_jalali()
        ctk.CTkLabel(self, text="سال:").grid(row=1, column=0, padx=10, pady=8, sticky="e")
        self.cb_year = ctk.CTkComboBox(self, width=100, values=[str(y) for y in range(1300, t.year + 2)])
        self.cb_year.set(str(t.year))
        self.cb_year.grid(row=1, column=1, pady=8, sticky="w")

        ctk.CTkLabel(self, text="ماه:").grid(row=2, column=0, padx=10, pady=8, sticky="e")
        self.cb_month = ctk.CTkComboBox(self, values=core.PERSIAN_MONTHS, width=150, state="readonly")
        self.cb_month.set(core.PERSIAN_MONTHS[0])
        self.cb_month.grid(row=2, column=1, pady=8, sticky="w")

        ctk.CTkLabel(self, text="روز:").grid(row=3, column=0, padx=10, pady=8, sticky="e")
        self.cb_day = ctk.CTkComboBox(self, width=80, values=[])
        self.cb_day.grid(row=3, column=1, pady=8, sticky="w")

        self.cb_year.configure(command=lambda _: self.update_days())
        self.cb_month.configure(command=lambda _: self.update_days())
        self.update_days()

        ctk.CTkLabel(self, text="توضیحات:").grid(row=4, column=0, padx=10, pady=8, sticky="ne")
        self.txt_note = ctk.CTkTextbox(self, width=220, height=80)
        self.txt_note.grid(row=4, column=1, pady=8, sticky="w")

        ctk.CTkButton(self, text="ثبت", command=self.on_submit).grid(
            row=5, column=0, columnspan=2, pady=12
        )

    def update_days(self, event=None):
        try:
            y = int(self.cb_year.get())
            m = core.PERSIAN_MONTHS.index(self.cb_month.get()) + 1
        except:
            return
        if m <= 6:
            dim = 31
        elif m <= 11:
            dim = 30
        else:
            dim = 30 if core.jdatetime.date.isleap(core.jdatetime.date(y, 1, 1)) else 29

        self.cb_day.configure(values=[str(d) for d in range(1, dim + 1)])
        self.cb_day.set("1")

    def on_submit(self):
        name = self.entry.get().strip()
        if not name:
            messagebox.showerror("خطا", "نام را وارد کنید")
            return
        y = int(self.cb_year.get())
        m = core.PERSIAN_MONTHS.index(self.cb_month.get()) + 1
        d = int(self.cb_day.get())
        date_str = f"{y:04d}-{m:02d}-{d:02d}"
        note = self.txt_note.get("1.0", "end").strip()
        self.callback({"name": name, "date": date_str, "note": note})
        self.destroy()


class ModernBirthdayApp(ctk.CTk):
    def __init__(self, silent=False):
        super().__init__()
        self.silent_mode = silent
        self.title("مدیریت تولدها")
        self.geometry("980x560")
        self.minsize(900, 520)

        if os.path.exists(core.ICON_PATH):
            try:
                self.iconbitmap(core.ICON_PATH)
            except:
                pass

        # startup shortcut
        core.create_startup_shortcut()

        self.enriched_data = []
        self.load_data()

        # silent mode (tray / notify)
        if self.silent_mode and core.TRAY_AVAILABLE and os.path.exists(core.ICON_PATH):
            image = core.Image.open(core.ICON_PATH)
            menu = core.pystray.Menu(core.pystray.MenuItem("خروج", lambda _: icon.stop()))
            global icon
            icon = core.pystray.Icon("BirthdayApp", image, "BirthdayApp", menu)
            core.Thread(target=lambda: core.start_tray_and_notify(self.enriched_data), daemon=True).start()
            icon.run()
            sys.exit(0)
        elif self.silent_mode:
            core.show_notifications(self.enriched_data)
            sys.exit(0)

        self.build_layout()
        core.show_notifications(self.enriched_data)
        self.refresh_table()

    def build_layout(self):
        # Sidebar
        sidebar = ctk.CTkFrame(self, width=200, corner_radius=10)
        sidebar.pack(side="left", fill="y", padx=12, pady=12)

        ctk.CTkLabel(sidebar, text="🎂 Birthday Manager", font=("Segoe UI", 16, "bold")).pack(pady=(20, 10))

        ctk.CTkButton(sidebar, text="➕ افزودن تولد", command=self.open_add).pack(pady=8, padx=15, fill="x")
        ctk.CTkButton(sidebar, text="🗑 حذف انتخاب", command=self.delete_selected).pack(pady=8, padx=15, fill="x")
        ctk.CTkButton(sidebar, text="🔄 بروزرسانی", command=self.refresh_table).pack(pady=8, padx=15, fill="x")

        # Main area
        main = ctk.CTkFrame(self, corner_radius=10)
        main.pack(side="right", fill="both", expand=True, padx=12, pady=12)

        # Top search bar
        search_frame = ctk.CTkFrame(main)
        search_frame.pack(fill="x", padx=12, pady=(12, 6))

        ctk.CTkLabel(search_frame, text="جستجو:").pack(side="left", padx=(8, 6))
        self.search_var = tk.StringVar()
        search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_var)
        search_entry.pack(side="left", fill="x", expand=True, padx=6)
        search_entry.bind('<KeyRelease>', lambda evt: self.refresh_table())

        # Table
        table_frame = ctk.CTkFrame(main)
        table_frame.pack(fill="both", expand=True, padx=12, pady=12)

        cols = ("row", "name", "date", "note", "days_left")
        heads = ["ردیف", "نام", "تاریخ تولد", "توضیحات", "روز مانده"]

        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        for c, h in zip(cols, heads):
            self.tree.heading(c, text=h)
            w = 60 if c == "row" else 180
            self.tree.column(c, width=w, anchor="center", stretch=True)

        self.tree.tag_configure('upcoming', foreground='red')

        vs = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    def load_data(self):
        raw = core.load_data()
        enriched = []
        for p in raw:
            d = core.compute_days_left(p.get("date", ""))
            enriched.append({**p, "days_left": d if d is not None else -1})
        enriched.sort(key=lambda x: x['days_left'])
        self.enriched_data = enriched

    def refresh_table(self):
        q = core.normalize(self.search_var.get().strip())
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, p in enumerate(self.enriched_data, start=1):
            txt = f"{p['name']} {p['date']} {p.get('note','')} {p['days_left']}"
            if q and q not in core.normalize(txt):
                continue
            tag = 'upcoming' if 0 <= p['days_left'] <= core.NOTIFY_BEFORE_DAYS else ''
            self.tree.insert("", tk.END,
                             values=(idx, p['name'], p['date'], p.get('note',''), p['days_left']),
                             tags=(tag,))

    def open_add(self):
        AddDialog(self, self.on_add)

    def on_add(self, person):
        data = core.load_data()
        data.append(person)
        core.save_data(data)
        self.load_data()
        core.show_notifications(self.enriched_data)
        self.refresh_table()

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("هشدار", "ابتدا یک ردیف انتخاب کنید")
            return
        vals = self.tree.item(sel[0], "values")
        data = core.load_data()
        data = [p for p in data if not (p['name'] == vals[1] and p['date'] == vals[2])]
        core.save_data(data)
        self.load_data()
        self.refresh_table()


if __name__ == "__main__":
    silent = "--silent" in sys.argv
    app = ModernBirthdayApp(silent=silent)
    app.mainloop()
