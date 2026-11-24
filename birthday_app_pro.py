"""
Birthday Reminder Pro - PyQt5 + Fluent Design
"""
import os, sys, json, datetime, jdatetime
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from qfluentwidgets import (
    FluentWindow, MSFluentWindow, PrimaryPushButton, PushButton,
    TableWidget, SearchLineEdit, ComboBox, SpinBox, TitleLabel,
    SubtitleLabel, CaptionLabel, CardWidget, HeaderCardWidget,
    FluentIcon as FIF, FluentTheme, setTheme, Theme
)
from threading import Thread
import time

# ====== وابستگی‌ها ======
try:
    import pystray
    from PIL import Image
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False
try:
    from win32com.client import Dispatch
except ImportError:
    Dispatch = None
try:
    from winotify import Notification, audio as win_audio
    WINOTOAST = True
except ImportError:
    WINOTOAST = False
try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    OPENPYXL = True
except ImportError:
    OPENPYXL = False

# ====== تنظیمات ======
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "birthdays.json")
NOTIFY_BEFORE_DAYS = 4
PERSIAN_MONTHS = [
    "فروردین", "اردیبهشت", "خرداد", "تیر",
    "مرداد", "شهریور", "مهر", "آبان",
    "آذر", "دی", "بهمن", "اسفند"
]
CATEGORIES = ["همه", "پرسنل", "مشتری", "سایر"]
ICON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
FALLBACK_ICON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")

# ====== رنگ‌ها ======
class Theme:
    BG = "#1B1B2F"
    BG_CARD = "#1F1F3A"
    BG_INPUT = "#2A2A4A"
    ACCENT = "#E94560"
    ACCENT_HOVER = "#C81E45"
    GREEN = "#4CAF50"
    BLUE = "#2196F3"
    ORANGE = "#FF9800"
    GOLD = "#FFD700"
    WHITE = "#FFFFFF"
    TEXT_DIM = "#A0A0B0"
    BORDER = "#3A3A5A"
    RED = "#FF6B6B"
    ROW_EVEN = "#1E1E38"
    ROW_ODD = "#222244"
    SELECT = "#3A1A2A"


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
    return jdatetime.date.fromgregorian(date=datetime.date.today())

def normalize(text):
    return text.replace('ي', 'ی').replace('ك', 'ک').lower()

def days_in_jalali_month(year, month):
    if month <= 6:
        return 31
    elif month <= 11:
        return 30
    else:
        return 30 if jdatetime.date.isleap(jdatetime.date(year, 1, 1)) else 29

def compute_days_left(birth_str):
    today = get_today_jalali()
    try:
        bd = jdatetime.datetime.strptime(birth_str, "%Y-%m-%d").date()
    except Exception:
        return None
    upcoming = jdatetime.date(today.year, bd.month, bd.day)
    delta = (upcoming.togregorian() - today.togregorian()).days
    if delta < 0:
        upcoming = jdatetime.date(today.year + 1, bd.month, bd.day)
        delta = (upcoming.togregorian() - today.togregorian()).days
    return delta

def get_category(note):
    n = normalize(note)
    for cat in ["پرسنل", "مشتری"]:
        if normalize(cat) in n:
            return cat
    return "سایر"

def clean_note(note):
    cat = get_category(note)
    if cat != "سایر":
        for m in [f"({cat})", f" ({cat})", f"({cat}) "]:
            note = note.replace(m, "")
    return note.strip()

def create_startup_shortcut():
    if Dispatch is None:
        return
    startup = os.path.join(os.environ.get("APPDATA", ""),
                           r"Microsoft\Windows\Start Menu\Programs\Startup")
    path = os.path.join(startup, "BirthdayAppPro.lnk")
    if os.path.exists(path):
        return
    try:
        shell = Dispatch('WScript.Shell')
        sc = shell.CreateShortCut(path)
        sc.TargetPath = sys.executable
        sc.Arguments = f'"{os.path.abspath(__file__)}" --silent'
        sc.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
        if os.path.exists(ICON_PATH):
            sc.IconLocation = ICON_PATH
        sc.save()
    except Exception:
        pass

def show_notifications(enriched):
    for p in enriched:
        if p['days_left'] == 0:
            title = "تولد مبارک!"
            msg = f"امروز تولد {p['name']} است."
        elif 0 < p['days_left'] <= NOTIFY_BEFORE_DAYS:
            title = "تولد در راه است"
            msg = f"{p['name']} در {p['days_left']} روز تولدش است."
        else:
            continue
        try:
            if WINOTOAST:
                toast = Notification(
                    app_id="BirthdayAppPro", title=title, msg=msg,
                    duration="short",
                    icon=ICON_PATH if os.path.exists(ICON_PATH) else None)
                toast.set_audio(win_audio.Mail, loop=False)
                toast.show()
            else:
                QMessageBox.information(None, title, msg)
        except Exception:
            pass

def export_excel(data, path):
    if not OPENPYXL:
        return False
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "تولدها"
    hfont = Font(name="B Nazanin", size=12, bold=True, color="FFFFFF")
    hfill = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
    halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    headers = ["ردیف", "نام", "تاریخ تولد", "دسته‌بندی", "توضیحات", "روز مانده"]
    widths = [8, 25, 18, 14, 35, 12]
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = hfont; c.fill = hfill; c.alignment = halign; c.border = border
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w
    df = Font(name="B Nazanin", size=11)
    da = Alignment(horizontal="center", vertical="center")
    fills = {"پرسنل": "E3F2FD", "مشتری": "E8F5E9", "سایر": "FFF3E0"}
    for ri, p in enumerate(data, 2):
        cat = get_category(p.get("note", ""))
        row = [ri-1, p['name'], p['date'], cat, clean_note(p.get("note","")), p['days_left']]
        fill = PatternFill(start_color=fills.get(cat, fills["سایر"]),
                           end_color=fills.get(cat, fills["سایر"]), fill_type="solid")
        for ci, v in enumerate(row, 1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.font = df; c.alignment = da; c.border = border; c.fill = fill
        if p['days_left'] <= NOTIFY_BEFORE_DAYS:
            ws.cell(row=ri, column=6).font = Font(name="B Nazanin", size=11, bold=True, color="FF0000")
    ws.auto_filter.ref = f"A1:F{len(data)+1}"
    ws.freeze_panes = "A2"
    wb.save(path)
    return True


# ====== ویجت انتخاب تاریخ جلالی ======
class JalaliDatePicker(QWidget):
    dateChanged = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        today = get_today_jalali()
        self._date_str = f"{today.year:04d}-{today.month:02d}-{today.day:02d}"

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        lbl_style = "color: #A0A0B0; font-size: 11px;"

        # روز
        lbl_r = QLabel("روز")
        lbl_r.setStyleSheet(lbl_style)
        layout.addWidget(lbl_r)
        self.day_spin = SpinBox()
        self.day_spin.setRange(1, 31)
        self.day_spin.setValue(today.day)
        self.day_spin.setFixedWidth(60)
        self.day_spin.valueChanged.connect(self._emit)
        layout.addWidget(self.day_spin)

        # ماه
        lbl_m = QLabel("ماه")
        lbl_m.setStyleSheet(lbl_style)
        layout.addWidget(lbl_m)
        self.month_combo = ComboBox()
        for m in PERSIAN_MONTHS:
            self.month_combo.addItem(m)
        self.month_combo.setCurrentIndex(today.month - 1)
        self.month_combo.setFixedWidth(110)
        self.month_combo.currentIndexChanged.connect(self._on_month_change)
        layout.addWidget(self.month_combo)

        # سال
        lbl_y = QLabel("سال")
        lbl_y.setStyleSheet(lbl_style)
        layout.addWidget(lbl_y)
        self.year_spin = SpinBox()
        self.year_spin.setRange(1300, today.year + 1)
        self.year_spin.setValue(today.year)
        self.year_spin.setFixedWidth(70)
        self.year_spin.valueChanged.connect(self._on_year_change)
        layout.addWidget(self.year_spin)

        layout.addStretch()

    def _on_month_change(self):
        self._update_day_range()
        self._emit()

    def _on_year_change(self):
        self._update_day_range()
        self._emit()

    def _update_day_range(self):
        y = self.year_spin.value()
        m = self.month_combo.currentIndex() + 1
        dim = days_in_jalali_month(y, m)
        cur = min(self.day_spin.value(), dim)
        self.day_spin.setRange(1, dim)
        self.day_spin.setValue(cur)

    def _emit(self):
        y = self.year_spin.value()
        m = self.month_combo.currentIndex() + 1
        d = self.day_spin.value()
        self._date_str = f"{y:04d}-{m:02d}-{d:02d}"
        self.dateChanged.emit(self._date_str)

    def get_date(self):
        return self._date_str

    def set_date(self, date_str):
        try:
            parts = date_str.split("-")
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            self.year_spin.setValue(y)
            self.month_combo.setCurrentIndex(m - 1)
            self._update_day_range()
            self.day_spin.setValue(d)
        except Exception:
            pass


# ====== دیالوگ افزودن/ویرایش ======
class PersonDialog(QDialog):
    def __init__(self, parent=None, person=None):
        super().__init__(parent)
        self.person = person
        self.result_data = None
        is_edit = person is not None
        self.setWindowTitle("ویرایش تولد" if is_edit else "افزودن تولد جدید")
        self.setFixedSize(420, 480)
        self.setStyleSheet(f"""
            QDialog {{ background-color: {Theme.BG}; }}
            QLabel {{ color: {Theme.WHITE}; font-size: 13px; }}
            QLineEdit, QPlainTextEdit {{
                background-color: {Theme.BG_INPUT}; color: {Theme.WHITE};
                border: 1px solid {Theme.BORDER}; border-radius: 8px;
                padding: 8px; font-size: 13px;
            }}
            QLineEdit:focus, QPlainTextEdit:focus {{
                border: 1px solid {Theme.ACCENT};
            }}
            QRadioButton {{
                color: {Theme.WHITE}; font-size: 12px; spacing: 8px;
            }}
            QRadioButton::indicator {{
                width: 16px; height: 16px; border-radius: 8px;
                border: 2px solid {Theme.BORDER};
                background-color: transparent;
            }}
            QRadioButton::indicator:checked {{
                background-color: {Theme.ACCENT}; border-color: {Theme.ACCENT};
            }}
            QSpinBox, QComboBox {{
                background-color: {Theme.BG_INPUT}; color: {Theme.WHITE};
                border: 1px solid {Theme.BORDER}; border-radius: 6px;
                padding: 4px 8px; font-size: 12px;
            }}
            QSpinBox:focus, QComboBox:focus {{
                border: 1px solid {Theme.ACCENT};
            }}
            QComboBox QAbstractItemView {{
                background-color: {Theme.BG_INPUT}; color: {Theme.WHITE};
                selection-background-color: {Theme.ACCENT};
            }}
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # عنوان
        title = TitleLabel("ویرایش اطلاعات" if is_edit else "افزودن تولد جدید")
        title.setStyleSheet(f"color: {Theme.ACCENT}; font-size: 18px; font-weight: bold;")
        main_layout.addWidget(title)

        # نام
        lbl_name = QLabel("نام:")
        lbl_name.setStyleSheet(f"color: {Theme.TEXT_DIM}; margin-top: 8px;")
        main_layout.addWidget(lbl_name)
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("نام و نام خانوادگی")
        main_layout.addWidget(self.name_edit)

        # دسته‌بندی
        lbl_cat = QLabel("دسته‌بندی:")
        lbl_cat.setStyleSheet(f"color: {Theme.TEXT_DIM}; margin-top: 4px;")
        main_layout.addWidget(lbl_cat)
        cat_layout = QHBoxLayout()
        self.cat_group = QButtonGroup(self)
        for i, cat in enumerate(["پرسنل", "مشتری"]):
            rb = QRadioButton(cat)
            if i == 0:
                rb.setChecked(True)
            self.cat_group.addButton(rb, i)
            cat_layout.addWidget(rb)
        cat_layout.addStretch()
        main_layout.addLayout(cat_layout)

        # تاریخ تولد
        lbl_date = QLabel("تاریخ تولد:")
        lbl_date.setStyleSheet(f"color: {Theme.TEXT_DIM}; margin-top: 4px;")
        main_layout.addWidget(lbl_date)
        self.date_picker = JalaliDatePicker()
        main_layout.addWidget(self.date_picker)

        # توضیحات
        lbl_note = QLabel("توضیحات:")
        lbl_note.setStyleSheet(f"color: {Theme.TEXT_DIM}; margin-top: 4px;")
        main_layout.addWidget(lbl_note)
        self.note_edit = QPlainTextEdit()
        self.note_edit.setPlaceholderText("توضیحات اختیاری...")
        self.note_edit.setFixedHeight(70)
        main_layout.addWidget(self.note_edit)

        main_layout.addStretch()

        # دکمه‌ها
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.btn_save = PrimaryPushButton("ذخیره")
        self.btn_save.setFixedSize(120, 38)
        self.btn_save.clicked.connect(self._on_submit)
        btn_layout.addWidget(self.btn_save)

        self.btn_cancel = PushButton("لغو")
        self.btn_cancel.setFixedSize(120, 38)
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        btn_layout.addStretch()
        main_layout.addLayout(btn_layout)

        # پر کردن در حالت ویرایش
        if is_edit:
            self.name_edit.setText(person.get("name", ""))
            cat = get_category(person.get("note", ""))
            if cat == "پرسنل":
                self.cat_group.button(0).setChecked(True)
            elif cat == "مشتری":
                self.cat_group.button(1).setChecked(True)
            self.note_edit.setPlainText(clean_note(person.get("note", "")))
            self.date_picker.set_date(person.get("date", ""))

        self.name_edit.setFocus()

    def _on_submit(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.critical(self, "خطا", "نام را وارد کنید")
            return
        date_str = self.date_picker.get_date()
        note = self.note_edit.toPlainText().strip()
        cat = "پرسنل" if self.cat_group.checkedId() == 0 else "مشتری"
        if cat:
            note = f"{note} ({cat})" if note else f"({cat})"
        self.result_data = {"name": name, "date": date_str, "note": note}
        self.accept()


# ====== ویجت تقویم جلالی ======
class JalaliCalendarWidget(CardWidget):
    dateSelected = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.today = get_today_jalali()
        self.current_year = self.today.year
        self.current_month = self.today.month
        self.selected_day = None

        self.setStyleSheet(f"""
            CardWidget {{
                background-color: {Theme.BG_CARD};
                border: 1px solid {Theme.BORDER};
                border-radius: 12px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        # هدر
        header = QHBoxLayout()
        self.btn_prev = PushButton("◀")
        self.btn_prev.setFixedSize(32, 28)
        self.btn_prev.clicked.connect(self._prev_month)
        header.addWidget(self.btn_prev)

        self.lbl_month = TitleLabel("")
        self.lbl_month.setStyleSheet(f"color: {Theme.ACCENT}; font-size: 14px;")
        self.lbl_month.setAlignment(Qt.AlignCenter)
        header.addWidget(self.lbl_month, 1)

        self.btn_next = PushButton("▶")
        self.btn_next.setFixedSize(32, 28)
        self.btn_next.clicked.connect(self._next_month)
        header.addWidget(self.btn_next)

        layout.addLayout(header)

        # هدر روزها
        days_header = QHBoxLayout()
        for d in ["ش", "ی", "د", "س", "چ", "پ", "ج"]:
            lbl = QLabel(d)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet(f"color: {Theme.ACCENT}; font-size: 11px; font-weight: bold; padding: 2px;")
            days_header.addWidget(lbl)
        layout.addLayout(days_header)

        # شبکه روزها
        self.grid = QGridLayout()
        self.grid.setSpacing(2)
        self.day_buttons = []
        for r in range(6):
            row_btns = []
            for c in range(7):
                btn = QPushButton("")
                btn.setFixedSize(32, 28)
                btn.setCursor(Qt.PointingHandCursor)
                btn.clicked.connect(lambda _, row=r, col=c: self._on_day_click(row, col))
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: transparent; color: {Theme.WHITE};
                        border: none; border-radius: 6px; font-size: 11px;
                    }}
                    QPushButton:hover {{ background-color: {Theme.BG_INPUT}; }}
                """)
                self.grid.addWidget(btn, r, c)
                row_btns.append(btn)
            self.day_buttons.append(row_btns)
        layout.addLayout(self.grid)

        self._refresh()

    def _prev_month(self):
        self.current_month -= 1
        if self.current_month < 1:
            self.current_month = 12
            self.current_year -= 1
        self._refresh()

    def _next_month(self):
        self.current_month += 1
        if self.current_month > 12:
            self.current_month = 1
            self.current_year += 1
        self._refresh()

    def _refresh(self):
        self.lbl_month.setText(f"{PERSIAN_MONTHS[self.current_month-1]} {self.current_year}")
        dim = days_in_jalali_month(self.current_year, self.current_month)
        first_greg = jdatetime.date(self.current_year, self.current_month, 1).togregorian()
        start_col = (first_greg.weekday() + 2) % 7

        for r in range(6):
            for c in range(7):
                btn = self.day_buttons[r][c]
                day_num = r * 7 + c - start_col + 1
                if 1 <= day_num <= dim:
                    btn.setText(str(day_num))
                    btn.setProperty("day", day_num)
                    # استایل پیش‌فرض
                    bg = "transparent"
                    fg = Theme.WHITE
                    font_w = "normal"
                    if (day_num == self.today.day and
                        self.current_month == self.today.month and
                        self.current_year == self.today.year):
                        bg = Theme.GOLD
                        fg = Theme.BG
                        font_w = "bold"
                    elif self.selected_day == day_num:
                        bg = Theme.ACCENT
                        fg = Theme.WHITE
                        font_w = "bold"
                    btn.setStyleSheet(f"""
                        QPushButton {{
                            background-color: {bg}; color: {fg};
                            border: none; border-radius: 6px;
                            font-size: 11px; font-weight: {font_w};
                        }}
                        QPushButton:hover {{ background-color: {Theme.BG_INPUT}; }}
                    """)
                else:
                    btn.setText("")
                    btn.setProperty("day", None)
                    btn.setStyleSheet("QPushButton { background: transparent; border: none; }")

    def _on_day_click(self, row, col):
        btn = self.day_buttons[row][col]
        day = btn.property("day")
        if day is not None:
            self.selected_day = day
            self._refresh()
            ds = f"{self.current_year:04d}-{self.current_month:02d}-{day:02d}"
            self.dateSelected.emit(ds)

    def set_date(self, date_str):
        try:
            parts = date_str.split("-")
            self.current_year = int(parts[0])
            self.current_month = int(parts[1])
            self.selected_day = int(parts[2])
            self._refresh()
        except Exception:
            pass


# ====== صفحه اصلی ======
class MainWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.enriched_data = []
        self._build_ui()
        self._refresh_all()

    def _build_ui(self):
        self.setStyleSheet(f"background-color: {Theme.BG};")

        root = QHBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(12)

        # ====== پنل راست (جدول + دکمه‌ها) ======
        right = QVBoxLayout()
        right.setSpacing(10)

        # نوار ابزار
        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        self.btn_add = PrimaryPushButton("  افزودن")
        self.btn_add.setIcon(FIF.ADD)
        self.btn_add.setFixedHeight(36)
        self.btn_add.clicked.connect(self._add_person)
        toolbar.addWidget(self.btn_add)

        self.btn_edit = PushButton("  ویرایش")
        self.btn_edit.setIcon(FIF.EDIT)
        self.btn_edit.setFixedHeight(36)
        self.btn_edit.clicked.connect(self._edit_person)
        toolbar.addWidget(self.btn_edit)

        self.btn_delete = PushButton("  حذف")
        self.btn_delete.setIcon(FIF.DELETE)
        self.btn_delete.setFixedHeight(36)
        self.btn_delete.clicked.connect(self._delete_person)
        toolbar.addWidget(self.btn_delete)

        toolbar.addStretch()

        self.btn_excel = PrimaryPushButton("  خروجی اکسل")
        self.btn_excel.setIcon(FIF.SAVE)
        self.btn_excel.setFixedHeight(36)
        self.btn_excel.clicked.connect(self._export_excel)
        toolbar.addWidget(self.btn_excel)

        right.addLayout(toolbar)

        # نوار جستجو و فیلتر
        filter_bar = QHBoxLayout()
        filter_bar.setSpacing(10)

        self.search = SearchLineEdit()
        self.search.setPlaceholderText("جستجو بر اساس نام...")
        self.search.setFixedWidth(280)
        self.search.textChanged.connect(self._refresh_table)
        filter_bar.addWidget(self.search)

        filter_bar.addStretch()

        lbl_f = QLabel("فیلتر:")
        lbl_f.setStyleSheet(f"color: {Theme.TEXT_DIM}; font-size: 12px;")
        filter_bar.addWidget(lbl_f)

        self.filter_combo = ComboBox()
        for c in CATEGORIES:
            self.filter_combo.addItem(c)
        self.filter_combo.setFixedWidth(100)
        self.filter_combo.currentIndexChanged.connect(self._refresh_table)
        filter_bar.addWidget(self.filter_combo)

        right.addLayout(filter_bar)

        # جدول
        self.table = TableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ["ردیف", "نام", "دسته‌بندی", "تاریخ تولد", "توضیحات", "روز مانده"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 45)
        self.table.setColumnWidth(2, 80)
        self.table.setColumnWidth(3, 100)
        self.table.setColumnWidth(5, 80)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(f"""
            TableWidget {{
                background-color: {Theme.BG_CARD};
                alternate-background-color: {Theme.ROW_ODD};
                color: {Theme.WHITE};
                border: 1px solid {Theme.BORDER};
                border-radius: 10px;
                gridline-color: {Theme.BORDER};
                font-size: 12px;
            }}
            QHeaderView::section {{
                background-color: {Theme.BG_CARD};
                color: {Theme.ACCENT};
                border: none;
                border-bottom: 2px solid {Theme.ACCENT};
                padding: 6px;
                font-size: 12px;
                font-weight: bold;
            }}
            QTableWidget::item {{
                padding: 4px;
                border: none;
            }}
            QTableWidget::item:selected {{
                background-color: {Theme.SELECT};
                color: {Theme.WHITE};
            }}
        """)
        self.table.doubleClicked.connect(self._edit_person)
        right.addWidget(self.table, 1)

        # نوار وضعیت
        self.status_bar = QLabel()
        self.status_bar.setStyleSheet(f"color: {Theme.TEXT_DIM}; font-size: 11px; padding: 4px;")
        right.addWidget(self.status_bar)

        root.addLayout(right, 2)

        # ====== پنل چپ (تقویم + نزدیک‌ترین تولدها) ======
        left = QVBoxLayout()
        left.setSpacing(10)

        # تقویم
        self.calendar = JalaliCalendarWidget()
        self.calendar.dateSelected.connect(self._on_cal_date)
        left.addWidget(self.calendar)

        # نزدیک‌ترین تولدها
        upcoming_card = HeaderCardWidget("نزدیک‌ترین تولدها")
        self.upcoming_layout = QVBoxLayout()
        self.upcoming_layout.setSpacing(4)
        upcoming_card.setLayout(self.upcoming_layout)
        left.addWidget(upcoming_card, 1)

        left.addStretch()
        root.addLayout(left, 1)

    def _compute_data(self):
        raw = load_data()
        enriched = []
        for p in raw:
            d = compute_days_left(p.get("date", ""))
            enriched.append({**p, "days_left": d if d is not None else -1})
        enriched.sort(key=lambda x: x['days_left'])
        return enriched

    def _refresh_all(self):
        self.enriched_data = self._compute_data()
        self._refresh_table()
        self._refresh_upcoming()
        self._update_status()
        show_notifications(self.enriched_data)

    def _refresh_table(self):
        q = normalize(self.search.text().strip())
        cat_filter = self.filter_combo.currentText()

        self.table.setRowCount(0)
        row_idx = 0
        for p in self.enriched_data:
            txt = f"{p['name']} {p['date']} {p.get('note','')} {p['days_left']}"
            if q and q not in normalize(txt):
                continue
            cat = get_category(p.get("note", ""))
            if cat_filter != "همه" and cat != cat_filter:
                continue

            self.table.insertRow(row_idx)

            items = [
                str(row_idx + 1),
                p['name'],
                cat if cat != "سایر" else "-",
                p['date'],
                clean_note(p.get("note", "")),
                str(p['days_left'])
            ]

            for col, val in enumerate(items):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                if col == 1:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

                # رنگ روز مانده
                if col == 5:
                    if p['days_left'] == 0:
                        item.setForeground(QColor(Theme.GOLD))
                        font = item.font()
                        font.setBold(True)
                        item.setFont(font)
                    elif p['days_left'] <= NOTIFY_BEFORE_DAYS:
                        item.setForeground(QColor(Theme.RED))
                        font = item.font()
                        font.setBold(True)
                        item.setFont(font)

                # رنگ دسته‌بندی
                if col == 2:
                    colors = {"پرسنل": Theme.BLUE, "مشتری": Theme.GREEN}
                    if cat in colors:
                        item.setForeground(QColor(colors[cat]))
                        font = item.font()
                        font.setBold(True)
                        item.setFont(font)

                self.table.setItem(row_idx, col, item)

            self.table.setRowHeight(row_idx, 32)
            row_idx += 1

    def _refresh_upcoming(self):
        # پاک کردن محتوای قبلی
        while self.upcoming_layout.count():
            child = self.upcoming_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        upcoming = [p for p in self.enriched_data if 0 <= p['days_left'] <= 14]
        if not upcoming:
            lbl = QLabel("تولدی در ۱۴ روز آینده نیست")
            lbl.setStyleSheet(f"color: {Theme.TEXT_DIM}; font-size: 11px; padding: 8px;")
            self.upcoming_layout.addWidget(lbl)
            return

        for p in upcoming[:8]:
            item = QHBoxLayout()
            item.setSpacing(6)

            days_text = "امروز!" if p['days_left'] == 0 else f"{p['days_left']} روز"
            color = Theme.GOLD if p['days_left'] == 0 else Theme.RED

            lbl_days = QLabel(days_text)
            lbl_days.setStyleSheet(f"color: {color}; font-size: 11px; font-weight: bold;")
            lbl_days.setFixedWidth(55)
            lbl_days.setAlignment(Qt.AlignCenter)
            item.addWidget(lbl_days)

            lbl_name = QLabel(p['name'])
            lbl_name.setStyleSheet(f"color: {Theme.WHITE}; font-size: 11px;")
            item.addWidget(lbl_name, 1)

            container = QWidget()
            container.setLayout(item)
            self.upcoming_layout.addWidget(container)

    def _update_status(self):
        total = len(self.enriched_data)
        near = len([p for p in self.enriched_data if 0 <= p['days_left'] <= NOTIFY_BEFORE_DAYS])
        today = get_today_jalali()
        self.status_bar.setText(
            f"مجموع: {total}  |  نزدیک: {near}  |  امروز: {today.strftime('%Y/%m/%d')}")

    def _on_cal_date(self, date_str):
        self.search.setText("")
        self.filter_combo.setCurrentIndex(0)
        # فیلتر جدول روی تاریخ انتخاب شده
        self.table.setRowCount(0)
        row_idx = 0
        for p in self.enriched_data:
            if p['date'] == date_str:
                self.table.insertRow(row_idx)
                cat = get_category(p.get("note", ""))
                items = [str(row_idx+1), p['name'], cat if cat != "سایر" else "-",
                         p['date'], clean_note(p.get("note","")), str(p['days_left'])]
                for col, val in enumerate(items):
                    item = QTableWidgetItem(val)
                    item.setTextAlignment(Qt.AlignCenter)
                    if col == 1:
                        item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.table.setItem(row_idx, col, item)
                self.table.setRowHeight(row_idx, 32)
                row_idx += 1

    def _add_person(self):
        dlg = PersonDialog(self)
        if dlg.exec_() == QDialog.Accepted and dlg.result_data:
            data = load_data()
            data.append(dlg.result_data)
            save_data(data)
            self._refresh_all()

    def _edit_person(self):
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            QMessageBox.warning(self, "هشدار", "ابتدا یک ردیف انتخاب کنید")
            return
        row = sel[0].row()
        name = self.table.item(row, 1).text()
        date = self.table.item(row, 3).text()

        for p in self.enriched_data:
            if p['name'] == name and p['date'] == date:
                dlg = PersonDialog(self, person=p)
                if dlg.exec_() == QDialog.Accepted and dlg.result_data:
                    data = load_data()
                    for i, orig in enumerate(data):
                        if orig['name'] == name and orig['date'] == date:
                            data[i] = dlg.result_data
                            break
                    save_data(data)
                    self._refresh_all()
                break

    def _delete_person(self):
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            QMessageBox.warning(self, "هشدار", "ابتدا یک ردیف انتخاب کنید")
            return
        row = sel[0].row()
        name = self.table.item(row, 1).text()
        date = self.table.item(row, 3).text()

        reply = QMessageBox.question(self, "تأیید حذف",
            f"آیا از حذف '{name}' اطمینان دارید؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        data = load_data()
        data = [p for p in data if not (p['name'] == name and p['date'] == date)]
        save_data(data)
        self._refresh_all()

    def _export_excel(self):
        if not OPENPYXL:
            QMessageBox.critical(self, "خطا",
                "کتابخانه openpyxl نصب نیست.\n\npip install openpyxl")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "ذخیره خروجی اکسل",
            f"birthdays_{get_today_jalali().strftime('%Y%m%d')}.xlsx",
            "Excel Files (*.xlsx)")
        if path:
            if export_excel(self.enriched_data, path):
                QMessageBox.information(self, "موفق", f"فایل ذخیره شد:\n{path}")


# ====== پنجره اصلی ======
class BirthdayMainWindow(MSFluentWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Birthday Reminder Pro")
        self.setWindowIcon(QIcon(ICON_PATH) if os.path.exists(ICON_PATH) else QIcon())
        self.resize(1000, 650)
        self.setMinimumSize(850, 550)

        setTheme(Theme.DARK)
        create_startup_shortcut()

        self.main_widget = MainWidget(self)
        self.hBoxLayout = QVBoxLayout(self)
        self.hBoxLayout.setContentsMargins(0, 0, 0, 0)
        self.hBoxLayout.addWidget(self.main_widget)

        # اعمال استایل کلی
        self.setStyleSheet(f"""
            MSFluentWindow {{
                background-color: {Theme.BG};
            }}
        """)


# ====== سیستم تری ======
class TrayManager:
    def __init__(self, enriched):
        self.enriched = enriched
        self.icon = None

    def run(self):
        show_notifications(self.enriched)
        time.sleep(5)

    def start(self):
        if not TRAY_AVAILABLE or not os.path.exists(ICON_PATH):
            return
        image = Image.open(ICON_PATH)
        menu = pystray.Menu(pystray.MenuItem("خروج", lambda _: self.icon.stop()))
        self.icon = pystray.Icon("BirthdayAppPro", image, "Birthday Reminder Pro", menu)
        Thread(target=self._run_thread, daemon=True).start()

    def _run_thread(self):
        self.run()
        if self.icon:
            self.icon.stop()


# ====== اجرا ======
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("BirthdayAppPro")
    app.setStyle("Fusion")

    # فونت
    font = app.font()
    font.setFamily("Segoe UI")
    font.setPointSize(10)
    app.setFont(font)

    if "--silent" in sys.argv:
        enriched = []
        raw = load_data()
        for p in raw:
            d = compute_days_left(p.get("date", ""))
            enriched.append({**p, "days_left": d if d is not None else -1})
        enriched.sort(key=lambda x: x['days_left'])

        if TRAY_AVAILABLE and os.path.exists(ICON_PATH):
            tray = TrayManager(enriched)
            tray.start()
            # نگه داشتن برنامه
            timer = QTimer()
            timer.timeout.connect(lambda: None)
            timer.start(1000)
        else:
            show_notifications(enriched)
            sys.exit(0)
    else:
        window = BirthdayMainWindow()
        window.show()

    sys.exit(app.exec_())
