# Birthday Reminder Pro - راهنمای پروژه

## فایل‌های نهایی
| فایل | توضیح |
|------|-------|
| `birthday_app_pro3.py` | **نسخه نهایی** - PyQt5 + Fluent Design |
| `birthday_app_pro3.spec` | فایل PyInstaller برای ساخت EXE |
| `dist/BirthdayAppPro.exe` | فایل اجرایی ساخته شده (~64MB) |
| `app.ico` | آیکون برنامه |
| `birthdays.json` | داده‌های تولد (76+ رکورد) |
| `Vazirmatn-Bold.ttf` | فونت فارسی |

## فایل‌های قدیمی (حذف نشوند)
| فایل | توضیح |
|------|-------|
| `birthday_app4.py` | نسخه اولیه tkinter |
| `birthday_app5.py` | نسخه با silent mode |
| `birthday_app6.py` | نسخه با system tray |
| `birthday_gui_Qwen4.py` | نسخه با فونت سفارشی (دارای خطای indentation) |
| `birthday_app_pro.py` | نسخه PyQt5 اولیه (دارای باگ FluentTheme) |
| `birthday_app_pro2.py` | نسخه PyQt5 اصلاح شده |

## وابستگی‌ها (pip install)
```
PyQt5
PyQt-Fluent-Widgets
jdatetime
openpyxl
winotify
pystray
Pillow
pywin32 (win32com)
```

## ساخت EXE
```bash
pyinstaller birthday_app_pro3.spec --noconfirm
# خروجی: dist/BirthdayAppPro.exe
```

## امکانات نسخه نهایی (pro3)
- رابط گرافیکی مدرن PyQt5 با تم تاریک
- تقویم جلالی با نشانگرهای رنگی تولد (آبی=پرسنل، سبز=مشتری، نارنجی=هر دو)
- Tooltip با هاور روی روزهای تقویم
- SpinBox سال/ماه/روز برای انتخاب تاریخ (بدون نیاز به 40 بار کلیک)
- جستجو و فیلتر دسته‌بندی (پرسنل/مشتری/همه)
- ویرایش و حذف با دابل‌کلیک
- خروجی اکسل با استایل حرفه‌ای
- نوتیفیکیشن ویندوز 4 روز قبل از تولد
- آیکون در system tray با pystray
- ایجاد شورتکات در استارتاپ ویندوز

## فرمت داده JSON
```json
{
  "name": "نام",
  "date": "1369-1-5",
  "note": "توضیحات (پرسنل)"
}
```
- دسته‌بندی در فیلد `note` با فرمت `(پرسنل)` یا `(مشتری)` ذخیره می‌شود
- تاریخ به صورت شمسی YYYY-MM-DD

## باگ‌های حل شده
1. **`FluentTheme` ImportError**: این کلاس در نسخه نصب شده qfluentwidgets وجود ندارد
2. **`QComboBox` و `QTimer` missing import**: فراموش شده بود اضافه شود
3. **`QPushButton` HTML rendering**: دایره رنگی تولد با setText HTML نمایش داده نمی‌شد → حل شد با DayCell ویجت جداگانه
4. **`MSFluentWindow` layout**: برای سادگی از QMainWindow استفاده شد
5. **`C.ORANGE` missing**: رنگ نارنجی فراموش شده بود اضافه شود

## نکات مهم
- برنامه با `--silent` در استارتاپ اجرا می‌شود (فقط نوتیفیکیشن بدون پنجره)
- فایل `birthdays.json` باید کنار EXE باشد
- آیکون `app.ico` باید کنار EXE باشد
