"""
Birthday Reminder Pro v2 - PyQt5 + Fluent Design
"""
import os, sys, json, datetime, jdatetime
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPlainTextEdit, QPushButton, QRadioButton,
    QButtonGroup, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QMessageBox, QFileDialog, QDialog,
    QApplication, QSpinBox, QFrame, QSplitter, QSizePolicy, QComboBox
)
from PyQt5.QtCore import Qt, pyqtSignal, QSize, QTimer
from PyQt5.QtGui import QIcon, QColor, QFont, QPalette
from qfluentwidgets import (
    PrimaryPushButton, PushButton, ComboBox, SearchLineEdit,
    TitleLabel, SubtitleLabel, CardWidget, HeaderCardWidget,
    setTheme, Theme, FluentIcon as FIF
)
from threading import Thread
import time

# ====== وابستگی‌ها ======
try:
    import pystray
    from PIL import Image as PILImage
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
    from openpyxl.styles import Font as XLFont, Alignment, PatternFill, Border, Side
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


# ====== رنگ‌ها ======
class C:
    BG = "#1B1B2F"
    BG2 = "#1F1F3A"
    INPUT = "#2A2A4A"
    ACCENT = "#E94560"
    ACCENT2 = "#C81E45"
    GREEN = "#4CAF50"
    BLUE = "#2196F3"
    GOLD = "#FFD700"
    WHITE = "#FFFFFF"
    DIM = "#A0A0B0"
    BORDER = "#3A3A5A"
    RED = "#FF6B6B"
    ORANGE = "#FF9800"
    ROW1 = "#1E1E38"
    ROW2 = "#222244"
    SEL = "#3A1A2A"


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
            title, msg = "تولد مبارک!", f"امروز تولد {p['name']} است."
        elif 0 < p['days_left'] <= NOTIFY_BEFORE_DAYS:
            title, msg = "تولد در راه است", f"{p['name']} در {p['days_left']} روز تولدش است."
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
    hfont = XLFont(name="B Nazanin", size=12, bold=True, color="FFFFFF")
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
    df = XLFont(name="B Nazanin", size=11)
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
            ws.cell(row=ri, column=6).font = XLFont(name="B Nazanin", size=11, bold=True, color="FF0000")
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

        lbl_s = f"color: {C.DIM}; font-size: 11px;"

        lbl_r = QLabel("روز"); lbl_r.setStyleSheet(lbl_s)
        layout.addWidget(lbl_r)
        self.day_spin = QSpinBox(); self.day_spin.setRange(1, 31)
        self.day_spin.setValue(today.day); self.day_spin.setFixedWidth(60)
        self.day_spin.setStyleSheet(f"""
            QSpinBox {{ background-color: {C.INPUT}; color: {C.WHITE};
                border: 1px solid {C.BORDER}; border-radius: 6px;
                padding: 4px 8px; font-size: 12px; }}
            QSpinBox:focus {{ border: 1px solid {C.ACCENT}; }}
            QSpinBox::up-button, QSpinBox::down-button {{
                width: 20px; border: none; background: {C.BG2};
                border-radius: 3px; }}
        """)
        self.day_spin.valueChanged.connect(self._emit)
        layout.addWidget(self.day_spin)

        lbl_m = QLabel("ماه"); lbl_m.setStyleSheet(lbl_s)
        layout.addWidget(lbl_m)
        self.month_combo = QComboBox()
        self.month_combo.addItems(PERSIAN_MONTHS)
        self.month_combo.setCurrentIndex(today.month - 1)
        self.month_combo.setFixedWidth(110)
        self.month_combo.setStyleSheet(f"""
            QComboBox {{ background-color: {C.INPUT}; color: {C.WHITE};
                border: 1px solid {C.BORDER}; border-radius: 6px;
                padding: 4px 8px; font-size: 12px; }}
            QComboBox:focus {{ border: 1px solid {C.ACCENT}; }}
            QComboBox::drop-down {{ border: none; width: 20px; }}
            QComboBox QAbstractItemView {{
                background-color: {C.INPUT}; color: {C.WHITE};
                selection-background-color: {C.ACCENT}; border: 1px solid {C.BORDER};
            }}
        """)
        self.month_combo.currentIndexChanged.connect(self._on_month_change)
        layout.addWidget(self.month_combo)

        lbl_y = QLabel("سال"); lbl_y.setStyleSheet(lbl_s)
        layout.addWidget(lbl_y)
        self.year_spin = QSpinBox(); self.year_spin.setRange(1300, today.year + 1)
        self.year_spin.setValue(today.year); self.year_spin.setFixedWidth(70)
        self.year_spin.setStyleSheet(f"""
            QSpinBox {{ background-color: {C.INPUT}; color: {C.WHITE};
                border: 1px solid {C.BORDER}; border-radius: 6px;
                padding: 4px 8px; font-size: 12px; }}
            QSpinBox:focus {{ border: 1px solid {C.ACCENT}; }}
            QSpinBox::up-button, QSpinBox::down-button {{
                width: 20px; border: none; background: {C.BG2};
                border-radius: 3px; }}
        """)
        self.year_spin.valueChanged.connect(self._on_year_change)
        layout.addWidget(self.year_spin)

        layout.addStretch()

    def _on_month_change(self):
        self._update_day_range(); self._emit()

    def _on_year_change(self):
        self._update_day_range(); self._emit()

    def _update_day_range(self):
        y = self.year_spin.value()
        m = self.month_combo.currentIndex() + 1
        dim = days_in_jalali_month(y, m)
        cur = min(self.day_spin.value(), dim)
        self.day_spin.setRange(1, dim); self.day_spin.setValue(cur)

    def _emit(self):
        y, m, d = self.year_spin.value(), self.month_combo.currentIndex() + 1, self.day_spin.value()
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
        self.setFixedSize(440, 500)
        self.setStyleSheet(f"""
            QDialog {{ background-color: {C.BG}; }}
            QLabel {{ color: {C.WHITE}; font-size: 13px; }}
            QLineEdit, QPlainTextEdit {{
                background-color: {C.INPUT}; color: {C.WHITE};
                border: 1px solid {C.BORDER}; border-radius: 8px;
                padding: 8px; font-size: 13px;
            }}
            QLineEdit:focus, QPlainTextEdit:focus {{ border: 1px solid {C.ACCENT}; }}
            QRadioButton {{
                color: {C.WHITE}; font-size: 12px; spacing: 8px;
            }}
            QRadioButton::indicator {{
                width: 16px; height: 16px; border-radius: 8px;
                border: 2px solid {C.BORDER}; background-color: transparent;
            }}
            QRadioButton::indicator:checked {{
                background-color: {C.ACCENT}; border-color: {C.ACCENT};
            }}
        """)

        ml = QVBoxLayout(self); ml.setSpacing(12); ml.setContentsMargins(20, 20, 20, 20)

        title = QLabel("ویرایش اطلاعات" if is_edit else "افزودن تولد جدید")
        title.setStyleSheet(f"color: {C.ACCENT}; font-size: 18px; font-weight: bold;")
        ml.addWidget(title)

        lbl = QLabel("نام:"); lbl.setStyleSheet(f"color: {C.DIM}; margin-top: 8px;")
        ml.addWidget(lbl)
        self.name_edit = QLineEdit(); self.name_edit.setPlaceholderText("نام و نام خانوادگی")
        ml.addWidget(self.name_edit)

        lbl = QLabel("دسته‌بندی:"); lbl.setStyleSheet(f"color: {C.DIM}; margin-top: 4px;")
        ml.addWidget(lbl)
        cat_lay = QHBoxLayout()
        self.cat_group = QButtonGroup(self)
        for i, cat in enumerate(["پرسنل", "مشتری"]):
            rb = QRadioButton(cat)
            if i == 0: rb.setChecked(True)
            self.cat_group.addButton(rb, i)
            cat_lay.addWidget(rb)
        cat_lay.addStretch()
        ml.addLayout(cat_lay)

        lbl = QLabel("تاریخ تولد:"); lbl.setStyleSheet(f"color: {C.DIM}; margin-top: 4px;")
        ml.addWidget(lbl)
        self.date_picker = JalaliDatePicker()
        ml.addWidget(self.date_picker)

        lbl = QLabel("توضیحات:"); lbl.setStyleSheet(f"color: {C.DIM}; margin-top: 4px;")
        ml.addWidget(lbl)
        self.note_edit = QPlainTextEdit(); self.note_edit.setPlaceholderText("توضیحات اختیاری...")
        self.note_edit.setFixedHeight(70)
        ml.addWidget(self.note_edit)

        ml.addStretch()

        btn_lay = QHBoxLayout(); btn_lay.setSpacing(10)
        self.btn_save = PrimaryPushButton("ذخیره")
        self.btn_save.setFixedSize(120, 38)
        self.btn_save.clicked.connect(self._on_submit)
        btn_lay.addWidget(self.btn_save)

        self.btn_cancel = PushButton("لغو")
        self.btn_cancel.setFixedSize(120, 38)
        self.btn_cancel.clicked.connect(self.reject)
        btn_lay.addWidget(self.btn_cancel)
        btn_lay.addStretch()
        ml.addLayout(btn_lay)

        if is_edit:
            self.name_edit.setText(person.get("name", ""))
            cat = get_category(person.get("note", ""))
            if cat == "پرسنل": self.cat_group.button(0).setChecked(True)
            elif cat == "مشتری": self.cat_group.button(1).setChecked(True)
            self.note_edit.setPlainText(clean_note(person.get("note", "")))
            self.date_picker.set_date(person.get("date", ""))

        self.name_edit.setFocus()

    def _on_submit(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.critical(self, "خطا", "نام را وارد کنید"); return
        date_str = self.date_picker.get_date()
        note = self.note_edit.toPlainText().strip()
        cat = "پرسنل" if self.cat_group.checkedId() == 0 else "مشتری"
        note = f"{note} ({cat})" if note else f"({cat})"
        self.result_data = {"name": name, "date": date_str, "note": note}
        self.accept()


# ====== تقویم جلالی ======
class DayCell(QWidget):
    """یک سلول در تقویم: عدد روز + نشانگر رنگی"""
    clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.day_num = 0
        self.setFixedSize(36, 38)
        self.setCursor(Qt.PointingHandCursor)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 1, 0, 0)
        lay.setSpacing(0)

        self.lbl_day = QLabel("")
        self.lbl_day.setAlignment(Qt.AlignCenter)
        self.lbl_day.setStyleSheet(f"color: {C.WHITE}; font-size: 11px; background: transparent; border: none;")
        lay.addWidget(self.lbl_day)

        self.lbl_dot = QLabel("")
        self.lbl_dot.setAlignment(Qt.AlignCenter)
        self.lbl_dot.setFixedHeight(6)
        self.lbl_dot.setStyleSheet("background: transparent; border: none;")
        lay.addWidget(self.lbl_dot)

        self._bg = "transparent"
        self._fg = C.WHITE

    def set_day(self, num):
        self.day_num = num
        self.lbl_day.setText(str(num) if num > 0 else "")

    def set_style(self, bg, fg, bold=False):
        self._bg = bg; self._fg = fg
        fw = "bold" if bold else "normal"
        self.lbl_day.setStyleSheet(f"color: {fg}; font-size: 11px; font-weight: {fw}; background: {bg}; border: none; border-radius: 4px;")

    def set_dot(self, color):
        if color:
            self.lbl_dot.setFixedHeight(6)
            self.lbl_dot.setStyleSheet(f"background-color: {color}; border-radius: 3px; margin: 0 8px;")
        else:
            self.lbl_dot.setFixedHeight(0)
            self.lbl_dot.setStyleSheet("background: transparent; border: none;")

    def paintEvent(self, event):
        super().paintEvent(event)

    def mousePressEvent(self, event):
        if self.day_num > 0:
            self.clicked.emit()


class JalaliCalendarWidget(QFrame):
    dateSelected = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.today = get_today_jalali()
        self.cy = self.today.year
        self.cm = self.today.month
        self.sd = None
        self.birthdays_map = {}
        self.setFrameShape(QFrame.NoFrame)
        self.setFixedWidth(280)
        self.setStyleSheet(f"""
            JalaliCalendarWidget {{
                background-color: {C.BG2};
                border: 1px solid {C.BORDER};
                border-radius: 12px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(4)

        # هدر
        hdr = QHBoxLayout()
        self.btn_p = QPushButton("◀")
        self.btn_p.setFixedSize(30, 26)
        self.btn_p.setStyleSheet(f"""
            QPushButton {{ background-color: {C.ACCENT}; color: {C.WHITE};
                border: none; border-radius: 6px; font-weight: bold; font-size: 13px; }}
            QPushButton:hover {{ background-color: {C.ACCENT2}; }}
        """)
        self.btn_p.setCursor(Qt.PointingHandCursor)
        self.btn_p.clicked.connect(self._prev)
        hdr.addWidget(self.btn_p)

        self.lbl_m = QLabel("")
        self.lbl_m.setAlignment(Qt.AlignCenter)
        self.lbl_m.setStyleSheet(f"color: {C.ACCENT}; font-size: 14px; font-weight: bold;")
        hdr.addWidget(self.lbl_m, 1)

        self.btn_n = QPushButton("▶")
        self.btn_n.setFixedSize(30, 26)
        self.btn_n.setStyleSheet(self.btn_p.styleSheet())
        self.btn_n.setCursor(Qt.PointingHandCursor)
        self.btn_n.clicked.connect(self._next)
        hdr.addWidget(self.btn_n)
        layout.addLayout(hdr)

        # هدر روزها
        dh = QHBoxLayout()
        for d in ["ش", "ی", "د", "س", "چ", "پ", "ج"]:
            lbl = QLabel(d)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet(f"color: {C.ACCENT}; font-size: 11px; font-weight: bold; padding: 2px;")
            dh.addWidget(lbl)
        layout.addLayout(dh)

        # شبکه سلول‌ها
        self.grid = QGridLayout()
        self.grid.setSpacing(1)
        self.cells = []
        for r in range(6):
            row = []
            for co in range(7):
                cell = DayCell()
                cell.clicked.connect(lambda rr=r, cc=co: self._click(rr, cc))
                self.grid.addWidget(cell, r, co)
                row.append(cell)
            self.cells.append(row)
        layout.addLayout(self.grid)

        # راهنمای رنگ‌ها
        legend = QHBoxLayout()
        legend.setSpacing(10)
        for color, label in [(C.BLUE, "پرسنل"), (C.GREEN, "مشتری"), (C.ORANGE, "هر دو")]:
            dot = QLabel("●")
            dot.setStyleSheet(f"color: {color}; font-size: 10px; background: transparent; border: none;")
            lbl = QLabel(label)
            lbl.setStyleSheet(f"color: {C.DIM}; font-size: 9px; background: transparent; border: none;")
            legend.addWidget(dot)
            legend.addWidget(lbl)
        legend.addStretch()
        layout.addLayout(legend)

        self._refresh()

    def set_birthdays(self, enriched_data):
        self.birthdays_map = {}
        for p in enriched_data:
            ds = p.get("date", "")
            cat = get_category(p.get("note", ""))
            name = p.get("name", "")
            if ds not in self.birthdays_map:
                self.birthdays_map[ds] = []
            self.birthdays_map[ds].append({"name": name, "cat": cat})
        self._refresh()

    def _prev(self):
        self.cm -= 1
        if self.cm < 1:
            self.cm = 12; self.cy -= 1
        self._refresh()

    def _next(self):
        self.cm += 1
        if self.cm > 12:
            self.cm = 1; self.cy += 1
        self._refresh()

    def _build_tooltip(self, day):
        ds = f"{self.cy:04d}-{self.cm:02d}-{day:02d}"
        people = self.birthdays_map.get(ds, [])
        if not people:
            return None
        lines = []
        for p in people:
            color = C.BLUE if p["cat"] == "پرسنل" else C.GREEN
            lines.append(f"<span style='color:{color};'>●</span> {p['name']} <span style='color:{C.DIM};'>({p['cat']})</span>")
        return "<br>".join(lines)

    def _refresh(self):
        self.lbl_m.setText(f"{PERSIAN_MONTHS[self.cm-1]} {self.cy}")
        dim = days_in_jalali_month(self.cy, self.cm)
        fg = jdatetime.date(self.cy, self.cm, 1).togregorian()
        sc = (fg.weekday() + 2) % 7

        for r in range(6):
            for co in range(7):
                cell = self.cells[r][co]
                dn = r * 7 + co - sc + 1

                if 1 <= dn <= dim:
                    cell.set_day(dn)
                    cell.show()

                    ds = f"{self.cy:04d}-{self.cm:02d}-{dn:02d}"
                    people = self.birthdays_map.get(ds, [])

                    # رنگ پس‌زمینه
                    bg, fg_c, bold = "transparent", C.WHITE, False
                    if dn == self.today.day and self.cm == self.today.month and self.cy == self.today.year:
                        bg, fg_c, bold = C.GOLD, C.BG, True
                    elif self.sd == dn:
                        bg, fg_c, bold = C.ACCENT, C.WHITE, True

                    cell.set_style(bg, fg_c, bold)

                    # نشانگر رنگی
                    if people:
                        cats = set(p["cat"] for p in people)
                        if "پرسنل" in cats and "مشتری" in cats:
                            dot_color = C.ORANGE
                        elif "پرسنل" in cats:
                            dot_color = C.BLUE
                        else:
                            dot_color = C.GREEN
                        cell.set_dot(dot_color)
                    else:
                        cell.set_dot(None)

                    # tooltip
                    tip = self._build_tooltip(dn)
                    cell.setToolTip(tip if tip else "")
                else:
                    cell.set_day(0)
                    cell.set_style("transparent", "transparent")
                    cell.set_dot(None)
                    cell.setToolTip("")
                    cell.hide()

    def _click(self, r, c):
        cell = self.cells[r][c]
        d = cell.day_num
        if d > 0:
            self.sd = d
            self._refresh()
            self.dateSelected.emit(f"{self.cy:04d}-{self.cm:02d}-{d:02d}")

    def set_date(self, ds):
        try:
            p = ds.split("-")
            self.cy, self.cm, self.sd = int(p[0]), int(p[1]), int(p[2])
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
        self.setStyleSheet(f"background-color: {C.BG};")
        root = QHBoxLayout(self); root.setContentsMargins(12, 12, 12, 12); root.setSpacing(12)

        # === پنل راست ===
        right = QVBoxLayout(); right.setSpacing(8)

        # تولبار
        tb = QHBoxLayout(); tb.setSpacing(6)
        self.btn_add = PrimaryPushButton("افزودن"); self.btn_add.setIcon(FIF.ADD)
        self.btn_add.setFixedHeight(34); self.btn_add.clicked.connect(self._add)
        tb.addWidget(self.btn_add)

        self.btn_edit = PushButton("ویرایش"); self.btn_edit.setIcon(FIF.EDIT)
        self.btn_edit.setFixedHeight(34); self.btn_edit.clicked.connect(self._edit)
        tb.addWidget(self.btn_edit)

        self.btn_del = PushButton("حذف"); self.btn_del.setIcon(FIF.DELETE)
        self.btn_del.setFixedHeight(34); self.btn_del.clicked.connect(self._delete)
        tb.addWidget(self.btn_del)

        tb.addStretch()
        self.btn_xl = PrimaryPushButton("خروجی اکسل"); self.btn_xl.setIcon(FIF.SAVE)
        self.btn_xl.setFixedHeight(34); self.btn_xl.clicked.connect(self._export)
        tb.addWidget(self.btn_xl)
        right.addLayout(tb)

        # جستجو و فیلتر
        fb = QHBoxLayout(); fb.setSpacing(8)
        self.search = SearchLineEdit(); self.search.setPlaceholderText("جستجو...")
        self.search.setFixedWidth(240); self.search.textChanged.connect(self._refresh_table)
        fb.addWidget(self.search)
        fb.addStretch()
        lbl = QLabel("فیلتر:"); lbl.setStyleSheet(f"color: {C.DIM}; font-size: 12px;")
        fb.addWidget(lbl)
        self.fcombo = QComboBox(); self.fcombo.addItems(CATEGORIES)
        self.fcombo.setFixedWidth(90)
        self.fcombo.setStyleSheet(f"""
            QComboBox {{ background-color: {C.INPUT}; color: {C.WHITE};
                border: 1px solid {C.BORDER}; border-radius: 6px; padding: 4px 8px; font-size: 12px; }}
            QComboBox::drop-down {{ border: none; width: 20px; }}
            QComboBox QAbstractItemView {{ background-color: {C.INPUT}; color: {C.WHITE};
                selection-background-color: {C.ACCENT}; border: 1px solid {C.BORDER}; }}
        """)
        self.fcombo.currentIndexChanged.connect(self._refresh_table)
        fb.addWidget(self.fcombo)
        right.addLayout(fb)

        # جدول
        self.table = QTableWidget(); self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(
            ["ردیف", "نام", "دسته‌بندی", "تاریخ تولد", "توضیحات", "روز مانده"])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 40); self.table.setColumnWidth(2, 75)
        self.table.setColumnWidth(3, 95); self.table.setColumnWidth(5, 70)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(f"""
            QTableWidget {{
                background-color: {C.BG2}; alternate-background-color: {C.ROW2};
                color: {C.WHITE}; border: 1px solid {C.BORDER};
                border-radius: 10px; gridline-color: {C.BORDER}; font-size: 12px;
            }}
            QHeaderView::section {{
                background-color: {C.BG2}; color: {C.ACCENT}; border: none;
                border-bottom: 2px solid {C.ACCENT}; padding: 6px;
                font-size: 12px; font-weight: bold;
            }}
            QTableWidget::item {{ padding: 4px; border: none; }}
            QTableWidget::item:selected {{ background-color: {C.SEL}; color: {C.WHITE}; }}
        """)
        self.table.doubleClicked.connect(self._edit)
        right.addWidget(self.table, 1)

        self.status = QLabel()
        self.status.setStyleSheet(f"color: {C.DIM}; font-size: 11px; padding: 4px;")
        right.addWidget(self.status)

        root.addLayout(right, 2)

        # === پنل چپ ===
        left = QVBoxLayout(); left.setSpacing(10)

        self.cal = JalaliCalendarWidget()
        self.cal.dateSelected.connect(self._cal_date)
        left.addWidget(self.cal)

        lbl2 = QLabel("نزدیک‌ترین تولدها")
        lbl2.setStyleSheet(f"color: {C.ACCENT}; font-size: 13px; font-weight: bold; margin-top: 6px;")
        left.addWidget(lbl2)

        self.up_frame = QFrame()
        self.up_frame.setFrameShape(QFrame.NoFrame)
        self.up_frame.setStyleSheet(f"""
            QFrame {{ background-color: {C.BG2}; border: 1px solid {C.BORDER};
                border-radius: 10px; }}
        """)
        self.up_layout = QVBoxLayout(self.up_frame); self.up_layout.setContentsMargins(8, 8, 8, 8)
        self.up_layout.setSpacing(3)
        left.addWidget(self.up_frame, 1)
        root.addLayout(left, 1)

    def _compute_data(self):
        raw = load_data(); enriched = []
        for p in raw:
            d = compute_days_left(p.get("date", ""))
            enriched.append({**p, "days_left": d if d is not None else -1})
        enriched.sort(key=lambda x: x['days_left'])
        return enriched

    def _refresh_all(self):
        self.enriched_data = self._compute_data()
        self.cal.set_birthdays(self.enriched_data)
        self._refresh_table(); self._refresh_upcoming(); self._update_status()
        show_notifications(self.enriched_data)

    def _refresh_table(self):
        q = normalize(self.search.text().strip())
        cf = self.fcombo.currentText()
        self.table.setRowCount(0); ri = 0
        for p in self.enriched_data:
            txt = f"{p['name']} {p['date']} {p.get('note','')} {p['days_left']}"
            if q and q not in normalize(txt): continue
            cat = get_category(p.get("note", ""))
            if cf != "همه" and cat != cf: continue
            self.table.insertRow(ri)
            vals = [str(ri+1), p['name'], cat if cat != "سایر" else "-",
                    p['date'], clean_note(p.get("note","")), str(p['days_left'])]
            for ci, v in enumerate(vals):
                it = QTableWidgetItem(v); it.setTextAlignment(Qt.AlignCenter)
                if ci == 1: it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if ci == 5:
                    if p['days_left'] == 0:
                        it.setForeground(QColor(C.GOLD)); f = it.font(); f.setBold(True); it.setFont(f)
                    elif p['days_left'] <= NOTIFY_BEFORE_DAYS:
                        it.setForeground(QColor(C.RED)); f = it.font(); f.setBold(True); it.setFont(f)
                if ci == 2 and cat in ("پرسنل", "مشتری"):
                    clr = C.BLUE if cat == "پرسنل" else C.GREEN
                    it.setForeground(QColor(clr)); f = it.font(); f.setBold(True); it.setFont(f)
                self.table.setItem(ri, ci, it)
            self.table.setRowHeight(ri, 30); ri += 1

    def _refresh_upcoming(self):
        while self.up_layout.count():
            ch = self.up_layout.takeAt(0)
            if ch.widget(): ch.widget().deleteLater()
        up = [p for p in self.enriched_data if 0 <= p['days_left'] <= 14]
        if not up:
            lbl = QLabel("تولدی در ۱۴ روز آینده نیست")
            lbl.setStyleSheet(f"color: {C.DIM}; font-size: 11px; padding: 8px;")
            self.up_layout.addWidget(lbl); return
        for p in up[:8]:
            h = QHBoxLayout(); h.setSpacing(6)
            dt = "امروز!" if p['days_left'] == 0 else f"{p['days_left']} روز"
            clr = C.GOLD if p['days_left'] == 0 else C.RED
            ld = QLabel(dt); ld.setStyleSheet(f"color: {clr}; font-size: 11px; font-weight: bold;")
            ld.setFixedWidth(55); ld.setAlignment(Qt.AlignCenter)
            h.addWidget(ld)
            ln = QLabel(p['name']); ln.setStyleSheet(f"color: {C.WHITE}; font-size: 11px;")
            h.addWidget(ln, 1)
            cw = QWidget(); cw.setLayout(h)
            self.up_layout.addWidget(cw)

    def _update_status(self):
        t = len(self.enriched_data)
        n = len([p for p in self.enriched_data if 0 <= p['days_left'] <= NOTIFY_BEFORE_DAYS])
        self.status.setText(f"مجموع: {t}  |  نزدیک: {n}  |  امروز: {get_today_jalali().strftime('%Y/%m/%d')}")

    def _cal_date(self, ds):
        self.search.setText(""); self.fcombo.setCurrentIndex(0)
        self.table.setRowCount(0); ri = 0
        for p in self.enriched_data:
            if p['date'] == ds:
                self.table.insertRow(ri)
                cat = get_category(p.get("note", ""))
                vals = [str(ri+1), p['name'], cat if cat != "سایر" else "-",
                        p['date'], clean_note(p.get("note","")), str(p['days_left'])]
                for ci, v in enumerate(vals):
                    it = QTableWidgetItem(v); it.setTextAlignment(Qt.AlignCenter)
                    if ci == 1: it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.table.setItem(ri, ci, it)
                self.table.setRowHeight(ri, 30); ri += 1

    def _add(self):
        dlg = PersonDialog(self)
        if dlg.exec_() == QDialog.Accepted and dlg.result_data:
            data = load_data(); data.append(dlg.result_data); save_data(data)
            self._refresh_all()

    def _edit(self):
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            QMessageBox.warning(self, "هشدار", "ابتدا یک ردیف انتخاب کنید"); return
        r = sel[0].row()
        name, date = self.table.item(r, 1).text(), self.table.item(r, 3).text()
        for p in self.enriched_data:
            if p['name'] == name and p['date'] == date:
                dlg = PersonDialog(self, person=p)
                if dlg.exec_() == QDialog.Accepted and dlg.result_data:
                    data = load_data()
                    for i, o in enumerate(data):
                        if o['name'] == name and o['date'] == date:
                            data[i] = dlg.result_data; break
                    save_data(data); self._refresh_all()
                break

    def _delete(self):
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            QMessageBox.warning(self, "هشدار", "ابتدا یک ردیف انتخاب کنید"); return
        r = sel[0].row()
        name, date = self.table.item(r, 1).text(), self.table.item(r, 3).text()
        if QMessageBox.question(self, "تأیید", f"حذف '{name}'؟",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes: return
        data = load_data()
        data = [p for p in data if not (p['name'] == name and p['date'] == date)]
        save_data(data); self._refresh_all()

    def _export(self):
        if not OPENPYXL:
            QMessageBox.critical(self, "خطا", "pip install openpyxl"); return
        path, _ = QFileDialog.getSaveFileName(
            self, "ذخیره", f"birthdays_{get_today_jalali().strftime('%Y%m%d')}.xlsx",
            "Excel (*.xlsx)")
        if path and export_excel(self.enriched_data, path):
            QMessageBox.information(self, "موفق", f"ذخیره شد:\n{path}")


# ====== پنجره اصلی ======
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Birthday Reminder Pro")
        self.setWindowIcon(QIcon(ICON_PATH) if os.path.exists(ICON_PATH) else QIcon())
        self.resize(1000, 650)
        self.setMinimumSize(850, 550)
        self.setStyleSheet(f"QMainWindow {{ background-color: {C.BG}; }}")

        central = QWidget()
        self.setCentralWidget(central)
        self.main_widget = MainWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.main_widget)


# ====== تری ======
class TrayManager:
    def __init__(self, enriched):
        self.enriched = enriched
        self.icon = None

    def start(self):
        if not TRAY_AVAILABLE or not os.path.exists(ICON_PATH): return
        image = PILImage.open(ICON_PATH)
        menu = pystray.Menu(pystray.MenuItem("خروج", lambda _: self.icon.stop()))
        self.icon = pystray.Icon("BirthdayAppPro", image, "Birthday Reminder Pro", menu)
        Thread(target=self._run, daemon=True).start()

    def _run(self):
        show_notifications(self.enriched)
        time.sleep(5)
        if self.icon: self.icon.stop()


# ====== اجرا ======
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("BirthdayAppPro")
    app.setStyle("Fusion")
    font = app.font(); font.setFamily("Segoe UI"); font.setPointSize(10)
    app.setFont(font)

    if "--silent" in sys.argv:
        enriched = []
        for p in load_data():
            d = compute_days_left(p.get("date", ""))
            enriched.append({**p, "days_left": d if d is not None else -1})
        enriched.sort(key=lambda x: x['days_left'])
        if TRAY_AVAILABLE and os.path.exists(ICON_PATH):
            tray = TrayManager(enriched); tray.start()
            timer = QTimer(); timer.timeout.connect(lambda: None); timer.start(1000)
        else:
            show_notifications(enriched); sys.exit(0)
    else:
        create_startup_shortcut()
        win = MainWindow(); win.show()

    sys.exit(app.exec_())
