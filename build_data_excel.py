# -*- coding: utf-8 -*-
"""
بناء ملف Excel لإدخال بيانات الفحص في أعمدة.
كل صف = مرحلة فحص واحدة لقطعة. الصفوف التي تحمل نفس "رقم القطعة"
تُجمَّع معاً لتوليد استمارة Word واحدة لتلك القطعة.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

NAVY, LIGHT, WHITE = "1F3864", "D9E1F2", "FFFFFF"
thin = Side(style="thin", color="808080")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def font(sz=11, b=False, c="000000"):
    return Font(name="Arial", size=sz, bold=b, color=c)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center", wrap_text=True)

wb = Workbook()
ws = wb.active
ws.title = "بيانات الفحص"
ws.sheet_view.rightToLeft = True

# الأعمدة
COLS = [
    ("رقم القطعة", 14),
    ("اسم القطعة", 22),
    ("اسم المشروع", 20),
    ("التاريخ", 14),
    ("المجموعة", 14),
    ("م", 6),
    ("مرحلة العمل", 24),
    ("عدد القطع", 10),
    ("الفني", 16),
    ("مقبول/مرفوض", 13),
    ("السماحية", 12),
    ("السبب", 24),
]

# صف تعليمات علوي
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
note = ws.cell(row=1, column=1,
    value="أدخل كل مرحلة فحص في صف. الصفوف التي تحمل نفس (رقم القطعة) تُنشئ استمارة Word واحدة. "
          "كرّر بيانات الترويسة (رقم/اسم القطعة/المشروع/التاريخ/المجموعة) في كل صفوف نفس القطعة.")
note.font = font(10, b=True, c="C00000")
note.alignment = right
ws.row_dimensions[1].height = 34

# رأس الأعمدة (الصف 2)
for c, (h, w) in enumerate(COLS, start=1):
    cell = ws.cell(row=2, column=c, value=h)
    cell.font = font(12, b=True, c=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = center
    cell.border = border
    ws.column_dimensions[cell.column_letter].width = w
ws.row_dimensions[2].height = 26

# بيانات عيّنة (قطعتان بمراحل مختلفة)
SAMPLE = [
    # رقم، اسم، مشروع، تاريخ، مجموعة، م، مرحلة العمل، عدد، فني، قرار، سماحية، سبب
    ("P-1001", "عمود إدارة رئيسي", "مشروع الخطوط الناقلة", "01/07/2026", "المجموعة أ", 1, "خراطة أولية", 5, "أحمد علي", "مقبول", "±0.05", ""),
    ("P-1001", "عمود إدارة رئيسي", "مشروع الخطوط الناقلة", "01/07/2026", "المجموعة أ", 2, "تفريز مجاري", 5, "أحمد علي", "مقبول", "±0.02", ""),
    ("P-1001", "عمود إدارة رئيسي", "مشروع الخطوط الناقلة", "01/07/2026", "المجموعة أ", 3, "تجليخ نهائي", 5, "سالم محمد", "مرفوض", "±0.01", "خدش سطحي خارج الحد"),
    ("P-1002", "غطاء علبة التروس", "مشروع الصيانة الدورية", "02/07/2026", "المجموعة ب", 1, "قص وتشكيل", 8, "خالد يحيى", "مقبول", "±0.10", ""),
    ("P-1002", "غطاء علبة التروس", "مشروع الصيانة الدورية", "02/07/2026", "المجموعة ب", 2, "ثقب المسامير", 8, "خالد يحيى", "مقبول", "±0.05", ""),
    ("P-1002", "غطاء علبة التروس", "مشروع الصيانة الدورية", "02/07/2026", "المجموعة ب", 3, "تسوية الأسطح", 8, "ماجد سعيد", "مقبول", "±0.03", ""),
    ("P-1002", "غطاء علبة التروس", "مشروع الصيانة الدورية", "02/07/2026", "المجموعة ب", 4, "طلاء واقٍ", 8, "ماجد سعيد", "مرفوض", "-", "تفاوت في سماكة الطلاء"),
]
for i, row in enumerate(SAMPLE, start=3):
    for c, val in enumerate(row, start=1):
        cell = ws.cell(row=i, column=c, value=val)
        cell.font = font(11)
        cell.alignment = center if c in (1, 4, 5, 6, 8, 10, 11) else right
        cell.border = border
    ws.row_dimensions[i].height = 20

# قائمة منسدلة لعمود مقبول/مرفوض (العمود J = 10)، والصفوف من 3 إلى 500
dv = DataValidation(type="list", formula1='"مقبول,مرفوض"', allow_blank=True)
ws.add_data_validation(dv)
dv.add("J3:J500")

# قائمة لعدد القطع (أعداد صحيحة موجبة) العمود H=8
dvq = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
ws.add_data_validation(dvq)
dvq.add("H3:H500")

# تجميد الرأس
ws.freeze_panes = "A3"

OUT = "/projects/sandbox/بيانات_الفحص.xlsx"
wb.save(OUT)
print("تم إنشاء:", OUT, "| أعمدة:", len(COLS), "| صفوف عينة:", len(SAMPLE))
