#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_labels.py - เพิ่มชีต "ป้ายชั้นวาง" 3 ชีต ลงในไฟล์ LAY_OUT_02 (พิมพ์ A4 ได้เลย)

ชีตที่สร้าง (วางไว้ก่อนชีต "สินค้า"):
    กรอกป้าย   หน้ากรอกข้อมูล 60 แถว: รหัสสินค้า / ตำแหน่ง (ไม่บังคับ) / จำนวน (ไม่บังคับ)
               ชื่อสินค้า + หน่วย ดึงจากชีต "สินค้า" เอง
               ตำแหน่งจากชีต "รวม" ดึงให้เอง (Location แรกที่รหัสตรงกัน) ถ้าไม่พิมพ์ตำแหน่งเอง
    ป้ายแบบ1   12 ป้าย/หน้า A4 (2 คอลัมน์ x 6 แถว ป้ายละ ~9.5 x 4.5 ซม.)  รหัส / ชื่อ / จำนวน
    ป้ายแบบ2   8 ป้าย/หน้า A4 (1 x 8 ป้ายละ ~19 x 3.4 ซม.)  หัวน้ำเงิน ตำแหน่ง|รหัส|จำนวน + ชื่อใหญ่ | จำนวน

วิธีใช้ (รันจาก root โปรเจค):
    python3 tools/build_labels.py                       # ทำกับ LAY_OUT_02.xlsx และ LAY_OUT_02_mockup.xlsx (ถ้ามี)
    python3 tools/build_labels.py --file path/to/a.xlsx # ระบุไฟล์เอง (ใส่ซ้ำได้หลายไฟล์)
    python3 tools/build_labels.py --file a.xlsx --mock  # ใส่ข้อมูลตัวอย่างใน กรอกป้าย (ถ้ายังว่าง)
    (ไฟล์ที่ชื่อมีคำว่า mockup จะใส่ข้อมูลตัวอย่างให้อัตโนมัติ)

รันซ้ำได้ (idempotent):
    - สร้างชีตทั้ง 3 ใหม่ทุกครั้ง ไม่ซ้ำซ้อน ไม่แตะชีตอื่น
    - ข้อมูลที่พิมพ์ไว้ใน กรอกป้าย (รหัส / ตำแหน่ง / จำนวน) ถูกอ่านเก็บไว้ก่อน แล้วเติมกลับให้
    - ก่อนเขียนทับ สำรองไฟล์เดิมไว้ที่ backup/ ข้างไฟล์

หลักการที่ใช้ (อ่านก่อนแก้):
    - ป้ายแต่ละใบดึงข้อมูลจากแถวที่ k ของ กรอกป้าย ด้วยสูตรธรรมดา (ไม่ใช้ FILTER/XLOOKUP/LET ที่ Excel เก่าไม่มี)
    - แถวว่าง = ป้ายว่าง: สูตรคืนค่า "" และสีพื้น/เส้นกรอบทั้งหมดเป็น Conditional Formatting
      ที่ทำงานเมื่อแถวนั้นมีข้อมูลเท่านั้น (ใช้คอลัมน์ช่วยที่ซ่อนอยู่ท้ายชีต เป็นตัวบอกว่าแถวไหนมีข้อมูล)
    - ขนาดป้ายคุมด้วยความกว้างคอลัมน์ (พิกเซล) + ความสูงแถว (pt) + ระยะขอบหน้ากระดาษ + scale 100%
      (ห้ามเปลี่ยนเป็น "พอดีหน้า" เพราะป้ายจะเล็กลงไม่คงที่)
    - จำนวน: พิมพ์ตัวเลข = พิมพ์ลงป้าย / เว้นว่าง = ป้ายมีช่องว่างไว้เขียนมือ
"""
import argparse
import datetime
import os
import re
import shutil
import sys
import zipfile

from openpyxl import load_workbook
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins

# ---------------------------------------------------------------- ค่าคงที่
SHEET_INPUT = "กรอกป้าย"
SHEET_S1 = "ป้ายแบบ1"
SHEET_S2 = "ป้ายแบบ2"
SHEET_PRODUCT = "สินค้า"
SHEET_RECORD = "รวม"

N = 60            # จำนวนแถว/ป้าย
FR = 6            # แถวแรกของข้อมูลใน กรอกป้าย
HDR = 5           # แถวหัวตาราง
LR = FR + N - 1   # แถวสุดท้าย (65)
FONT = "Tahoma"   # ฟอนต์ไทยที่มีใน Excel ทุกเครื่อง

# ช่วงอ้างอิงชีต รวม (แถวข้อมูล 5-2004) และ สินค้า
REC_LOC = SHEET_RECORD + "!$B$5:$B$2004"
REC_COD = SHEET_RECORD + "!$C$5:$C$2004"
PROD = SHEET_PRODUCT + "!$A:$C"

# สี
C_BROWN = "5A2A12"    # ตัวหนังสือน้ำตาลเข้ม (แบบ 1 - จากรูปที่ 1)
C_SKY = "C9E2F5"      # ฟ้าอ่อนพื้นป้าย (แบบ 1)
C_NAVY = "1E3FAE"     # น้ำเงินเข้ม หัวป้าย + กรอบ (แบบ 2 - จากรูปที่ 2)
C_GUIDE = "7F7F7F"    # เส้นตัดสีเทา


def px(cm):
    """ซม. -> พิกเซล (96 dpi)"""
    return round(cm / 2.54 * 96)


def set_col_px(ws, col_idx, pixels):
    # Excel เก็บความกว้างเป็นหน่วยตัวอักษร (ฟอนต์ Calibri 11 = 7 พิกเซล/ตัว)
    ws.column_dimensions[CL(col_idx)].width = pixels / 7.0


def q(sheet_cell):
    """อ้างอิงเซลล์ในชีต กรอกป้าย"""
    return "'%s'!%s" % (SHEET_INPUT, sheet_cell)


# ---------------------------------------------------------------- Conditional Formatting
def solid(color):
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def cf_add(ws, ranges, flag_col, fill=None, font_color=None, border=None):
    """เพิ่มกฎ CF 1 ข้อให้หลายช่วง (ranges = list ของ 'A1:B2')
    สูตรเงื่อนไข =$<flag_col><แถวบนสุดของช่วงแรก>=1  (อ้างอิงสัมพัทธ์แถว: ช่วงแรกต้องเป็นช่วงที่อยู่บนสุดเสมอ)"""
    if not ranges:
        return
    top_row = int(re.search(r"\d+", ranges[0]).group())
    flag_formula = "$%s%d=1" % (flag_col, top_row)
    dxf = DifferentialStyle(
        fill=solid(fill) if fill else None,
        font=Font(color=font_color) if font_color else None,
        border=border,
    )
    rule = Rule(type="expression", dxf=dxf, formula=[flag_formula])
    ws.conditional_formatting.add(" ".join(ranges), rule)


def side(color=C_GUIDE):
    return Side(style="thin", color=color)


# ---------------------------------------------------------------- เตรียมหน้ากระดาษ
def setup_page(ws, left_right_cm, top_bottom_cm, last_row, last_col, rows_per_page):
    ws.page_setup.paperSize = 9               # A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.scale = 100                 # ขนาดจริง ห้าม fit-to-page
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=False)
    ws.page_margins = PageMargins(
        left=left_right_cm / 2.54, right=left_right_cm / 2.54,
        top=top_bottom_cm / 2.54, bottom=top_bottom_cm / 2.54,
        header=0.2, footer=0.2)
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True
    ws.print_options.gridLines = False
    ws.sheet_view.showGridLines = False
    ws.print_area = "A1:%s%d" % (CL(last_col), last_row)
    # ตัดหน้าเอง: ป้ายจะไม่ถูกผ่าครึ่งข้ามหน้า
    for r in range(rows_per_page, last_row, rows_per_page):
        ws.row_breaks.append(Break(id=r))


# ---------------------------------------------------------------- ชีต กรอกป้าย
def build_input(wb, idx, old, has_record):
    ws = wb.create_sheet(SHEET_INPUT, idx)
    ws.sheet_properties.tabColor = "FFC000"
    thin = Side(style="thin", color="A6A6A6")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = "กรอกข้อมูลป้าย"
    ws["A1"].font = Font(name=FONT, size=18, bold=True, color="1F3864")
    notes = [
        "1) พิมพ์ \"รหัสสินค้า\" ในช่องสีเหลือง ชื่อสินค้าและหน่วยขึ้นให้เอง   "
        "2) ตำแหน่ง: เว้นว่างได้ ถ้ารหัสอยู่ในชีต \"รวม\" จะดึง Location มาให้ (ช่องเทา) ถ้าพิมพ์เอง จะใช้ที่พิมพ์",
        "3) จำนวน: พิมพ์ตัวเลข = พิมพ์ลงป้าย / เว้นว่าง = ป้ายเว้นช่องไว้เขียนมือ   "
        "4) แถวที่ว่างทั้งแถว = ไม่ออกป้าย   แถวที่ k ในตารางนี้ = ป้ายใบที่ k",
        "พิมพ์ที่ชีต \"ป้ายแบบ1\" (12 ป้าย/A4) หรือ \"ป้ายแบบ2\" (8 ป้าย/A4)   "
        "ตั้งพิมพ์: A4 แนวตั้ง ขนาด 100% (ห้ามเลือก \"ย่อให้พอดีหน้า\")",
    ]
    for i, t in enumerate(notes):
        r = 2 + i
        ws.merge_cells("A%d:G%d" % (r, r))
        ws["A%d" % r] = t
        ws["A%d" % r].font = Font(name=FONT, size=10, color="404040")
        ws["A%d" % r].alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 27

    heads = ["ลำดับ", "รหัสสินค้า", "รายการ", "ตำแหน่ง", "ตำแหน่งจากชีตรวม", "จำนวน", "หน่วย"]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(row=HDR, column=i, value=h)
        c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
        c.fill = solid("1F3864" if i not in (2, 4, 6) else "BF8F00")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = box
    ws.row_dimensions[HDR].height = 30

    for w, i in zip([7, 17, 46, 18, 20, 12, 9], range(1, 8)):
        ws.column_dimensions[CL(i)].width = w

    yellow = solid("FFF2CC")
    grey = solid("EDEDED")
    for k in range(N):
        r = FR + k
        ws.row_dimensions[r].height = 21
        ws.cell(row=r, column=1, value=k + 1)
        # รายการ: ดึงชื่อจากชีตสินค้า
        ws.cell(row=r, column=3).value = (
            '=IF($B{r}="","",IFERROR(VLOOKUP(TRIM($B{r})&"",{p},2,FALSE)&"","ไม่พบรหัส"))'
            .format(r=r, p=PROD))
        # ตำแหน่งจากชีตรวม: Location แรกที่รหัสตรง (ลองแบบข้อความ แล้วลองแบบตัวเลข)
        if has_record:
            ws.cell(row=r, column=5).value = (
                '=IF($B{r}="","",IFERROR(INDEX({loc},MATCH(TRIM($B{r})&"",{cod},0))&"",'
                'IFERROR(INDEX({loc},MATCH(VALUE(TRIM($B{r})),{cod},0))&"","")))'
                .format(r=r, loc=REC_LOC, cod=REC_COD))
        else:
            ws.cell(row=r, column=5).value = '=""'
        ws.cell(row=r, column=7).value = (
            '=IF($B{r}="","",IFERROR(VLOOKUP(TRIM($B{r})&"",{p},3,FALSE)&"",""))'
            .format(r=r, p=PROD))
        for col in range(1, 8):
            c = ws.cell(row=r, column=col)
            c.border = box
            c.font = Font(name=FONT, size=11, color="7F7F7F" if col == 5 else "000000")
            c.alignment = Alignment(vertical="center",
                                    horizontal="center" if col in (1, 2, 4, 5, 6, 7) else "left")
            if col in (2, 4, 6):
                c.fill = yellow
            elif col != 1:
                c.fill = grey
        ws.cell(row=r, column=1).font = Font(name=FONT, size=10, color="7F7F7F")
        ws.cell(row=r, column=2).number_format = "@"   # รหัสเป็นข้อความ (กันเลขศูนย์หาย)
        ws.cell(row=r, column=4).number_format = "@"
        ws.cell(row=r, column=6).number_format = "General"

    # ใส่ข้อมูลที่เคยพิมพ์ไว้กลับคืน
    for k, (code, loc, qty) in enumerate(old):
        r = FR + k
        if code not in (None, ""):
            ws.cell(row=r, column=2).value = str(code)
        if loc not in (None, ""):
            ws.cell(row=r, column=4).value = loc
        if qty not in (None, ""):
            ws.cell(row=r, column=6).value = qty

    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                        allow_blank=True, showErrorMessage=True,
                        errorTitle="จำนวนไม่ถูกต้อง", error="กรอกเป็นตัวเลข 0 ขึ้นไป หรือเว้นว่างไว้")
    ws.add_data_validation(dv)
    dv.add("F%d:F%d" % (FR, LR))

    ws.freeze_panes = "A%d" % FR
    ws.sheet_view.showGridLines = False
    # พิมพ์หน้าตารางนี้: A4 แนวตั้ง พอดีความกว้าง
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_area = "A1:G%d" % LR
    ws.print_title_rows = "%d:%d" % (HDR, HDR)
    return ws


# ---------------------------------------------------------------- ตัวช่วยสร้างป้าย
def src_formulas(k):
    """สูตรดึงข้อมูลของป้ายใบที่ k (0-59) จากแถว FR+k ของ กรอกป้าย"""
    r = FR + k
    B, C, D, E, F, G = (q("$%s%d" % (c, r)) for c in "BCDEFG")
    return {
        # flag = แถวนี้มีข้อมูลไหม (แบบ 2 แสดงตำแหน่งด้วย จึงนับตำแหน่งที่พิมพ์เอง)
        "flag": '=IF(OR({B}<>"",{D}<>"",{F}<>""),1,"")'.format(B=B, D=D, F=F),
        # แบบ 1 ไม่แสดงตำแหน่ง -> นับเฉพาะรหัส/จำนวน
        "flag1": '=IF(OR({B}<>"",{F}<>""),1,"")'.format(B=B, F=F),
        "code": '=IF({B}="","",TRIM({B}))'.format(B=B),
        "name": '=IF({B}="","",{C})'.format(B=B, C=C),
        "loc": '=IF({D}<>"",{D},IF({B}="","",{E}))'.format(B=B, D=D, E=E),
        "qty": '=IF({F}="","",{F})'.format(F=F),
        "unit": '=IF({B}="","",{G})'.format(B=B, G=G),
    }


def style_cell(c, size, bold=False, color="000000", halign="center", wrap=False, shrink=False):
    c.font = Font(name=FONT, size=size, bold=bold, color=color)
    c.alignment = Alignment(horizontal=halign, vertical="center", wrap_text=wrap,
                            shrink_to_fit=shrink)


# ---------------------------------------------------------------- ป้ายแบบ 1  (2 x 6, ~9.5 x 4.5 ซม.)
# ในป้าย 1 ใบมี 5 คอลัมน์: ขอบซ้าย | ซ้าย(รหัส/ชื่อ) | คั่น | ขวา(จำนวน) | ขอบขวา
# และ 6 แถว: ขอบบน | หัว(รหัส, "จำนวน") | คั่น | ตัว(ชื่อ, ช่องจำนวน) | (ชื่อต่อ, หน่วย) | ขอบล่าง
S1_COLS = [6, 220, 6, 121, 6]                 # พิกเซล รวม 359 px = 9.5 ซม.
S1_ROWS = [4.5, 30, 3, 52.5, 33, 4.5]         # pt รวม 127.5 pt = 4.50 ซม.
S1_PER_PAGE = 12


def build_style1(wb, idx):
    ws = wb.create_sheet(SHEET_S1, idx)
    ws.sheet_properties.tabColor = "9DC3E6"
    ncols = 5 * 2
    for lab_col in range(2):
        for j, p in enumerate(S1_COLS):
            set_col_px(ws, lab_col * 5 + j + 1, p)
    flag_cols = [ncols + 1, ncols + 2]        # K, L = คอลัมน์ช่วย (ซ่อน) ของป้ายซ้าย/ขวา
    for fc in flag_cols:
        ws.column_dimensions[CL(fc)].hidden = True

    n_pages = N // S1_PER_PAGE
    rows_page = 6 * len(S1_ROWS)
    last_row = n_pages * rows_page

    rng = {"fill": [[], []], "L": [[], []], "R": [[], []], "T": [[], []], "B": [[], []],
           "TL": [[], []], "TR": [[], []], "BL": [[], []], "BR": [[], []]}

    for k in range(N):
        page, i = divmod(k, S1_PER_PAGE)
        rr, cc = divmod(i, 2)
        r0 = page * rows_page + rr * len(S1_ROWS) + 1     # แถวแรก (ขอบบน) ของป้าย
        c0 = cc * 5 + 1                                    # คอลัมน์แรก (ขอบซ้าย)
        f = src_formulas(k)
        R = [r0 + j for j in range(6)]
        Cc = [c0 + j for j in range(5)]
        fcol = CL(flag_cols[cc])
        for j, h in enumerate(S1_ROWS):
            ws.row_dimensions[R[j]].height = h
        # คอลัมน์ช่วย: ใส่ในทุกแถวของป้าย (ให้ CF อ้างแถวเดียวกันได้)
        for rj in R:
            ws["%s%d" % (fcol, rj)] = f["flag1"]
        # เนื้อหา
        a = lambda col, row: "%s%d" % (CL(col), row)
        ws[a(Cc[1], R[1])] = f["code"]
        style_cell(ws[a(Cc[1], R[1])], 20, True, C_BROWN, shrink=True)
        ws[a(Cc[3], R[1])] = '=IF(%s%d=1,"จำนวน","")' % (fcol, R[1])
        style_cell(ws[a(Cc[3], R[1])], 15, True, C_BROWN, shrink=True)
        ws.merge_cells("%s:%s" % (a(Cc[1], R[3]), a(Cc[1], R[4])))
        ws[a(Cc[1], R[3])] = f["name"]
        style_cell(ws[a(Cc[1], R[3])], 16, False, C_BROWN, wrap=True)
        ws[a(Cc[3], R[3])] = f["qty"]
        style_cell(ws[a(Cc[3], R[3])], 26, True, "000000", shrink=True)
        ws[a(Cc[3], R[4])] = f["unit"]
        style_cell(ws[a(Cc[3], R[4])], 13, False, C_BROWN, shrink=True)
        # ช่วงสำหรับ CF
        g = rng
        # ช่วง CF แยกกันไม่ซ้อนกัน (กันปัญหา LibreOffice ใช้กฎแรกข้อเดียว) แต่ละกฎ = สีพื้น + เส้นขอบของตำแหน่งนั้น
        g["fill"][cc] += ["%s:%s" % (a(Cc[2], R[1]), a(Cc[2], R[4])),           # คอลัมน์คั่นกลาง (ต้องมาก่อน = บนสุด)
                          a(Cc[1], R[2]), a(Cc[3], R[2])]                       # แถวคั่นกลาง (ช่องซ้าย/ขวา)
        g["L"][cc].append("%s:%s" % (a(Cc[0], R[1]), a(Cc[0], R[4])))
        g["R"][cc].append("%s:%s" % (a(Cc[4], R[1]), a(Cc[4], R[4])))
        g["T"][cc].append("%s:%s" % (a(Cc[1], R[0]), a(Cc[3], R[0])))
        g["B"][cc].append("%s:%s" % (a(Cc[1], R[5]), a(Cc[3], R[5])))
        g["TL"][cc].append(a(Cc[0], R[0]))
        g["TR"][cc].append(a(Cc[4], R[0]))
        g["BL"][cc].append(a(Cc[0], R[5]))
        g["BR"][cc].append(a(Cc[4], R[5]))

    # กฎ CF (ทำงานเมื่อแถวนั้นมีข้อมูลเท่านั้น -> แถวว่าง = ป้ายว่างเปล่า)
    S = side()
    borders = {
        "L": Border(left=S), "R": Border(right=S), "T": Border(top=S), "B": Border(bottom=S),
        "TL": Border(left=S, top=S), "TR": Border(right=S, top=S),
        "BL": Border(left=S, bottom=S), "BR": Border(right=S, bottom=S),
    }
    for cc in range(2):
        flag = CL(flag_cols[cc])
        cf_add(ws, rng["fill"][cc], flag, fill=C_SKY)
        for key, b in borders.items():
            cf_add(ws, rng[key][cc], flag, fill=C_SKY, border=b)

    setup_page(ws, 0.6, 1.2, last_row, ncols, rows_page)
    return ws


# ---------------------------------------------------------------- ป้ายแบบ 2  (1 x 8, ~19 x 3.4 ซม.)
# คอลัมน์: ขอบซ้าย | ตำแหน่ง | รหัส | คั่น | จำนวน | หน่วย | ขอบขวา      (px รวม 718 = 19.0 ซม.)
# แถว    : ช่องว่าง(เส้นตัด) | หัวน้ำเงิน | ตัว | ขอบล่าง | ช่องว่าง          (pt รวม 96 = 3.39 ซม.)
S2_COLS = [6, 185, 340, 6, 126, 49, 6]
S2_ROWS = [3, 28.5, 57, 4.5, 3]
S2_PER_PAGE = 8


def build_style2(wb, idx):
    ws = wb.create_sheet(SHEET_S2, idx)
    ws.sheet_properties.tabColor = "1E3FAE"
    ncols = len(S2_COLS)
    for j, p in enumerate(S2_COLS):
        set_col_px(ws, j + 1, p)
    flag_col = ncols + 1                      # H = คอลัมน์ช่วย (ซ่อน)
    ws.column_dimensions[CL(flag_col)].hidden = True
    fl = CL(flag_col)

    n_pages = -(-N // S2_PER_PAGE)            # 8 หน้า (หน้าสุดท้ายว่างบางส่วน)
    rows_page = S2_PER_PAGE * len(S2_ROWS)
    last_row = n_pages * rows_page

    navy, guide, band = [], [], []
    sep = []
    for slot in range(n_pages * S2_PER_PAGE):
        r0 = slot * len(S2_ROWS) + 1
        R = [r0 + j for j in range(len(S2_ROWS))]
        for j, h in enumerate(S2_ROWS):
            ws.row_dimensions[R[j]].height = h
        if slot >= N:
            continue                           # ช่องเกิน 60 = ว่างเสมอ (ไม่มีสูตร)
        k = slot
        f = src_formulas(k)
        for rj in R:
            ws["%s%d" % (fl, rj)] = f["flag"]
        a = lambda col, row: "%s%d" % (CL(col), row)
        # หัว
        ws[a(2, R[1])] = f["loc"]
        style_cell(ws[a(2, R[1])], 18, True, "000000", shrink=True)
        ws[a(3, R[1])] = f["code"]
        style_cell(ws[a(3, R[1])], 18, True, "000000", shrink=True)
        ws.merge_cells("%s:%s" % (a(5, R[1]), a(6, R[1])))
        ws[a(5, R[1])] = '=IF(%s%d=1,"จำนวน","")' % (fl, R[1])
        style_cell(ws[a(5, R[1])], 16, True, "000000", shrink=True)
        # ตัว
        ws.merge_cells("%s:%s" % (a(2, R[2]), a(3, R[2])))
        ws[a(2, R[2])] = f["name"]
        style_cell(ws[a(2, R[2])], 20, True, "000000", wrap=True)
        ws[a(5, R[2])] = f["qty"]
        style_cell(ws[a(5, R[2])], 30, True, "000000", shrink=True)
        ws[a(6, R[2])] = f["unit"]
        style_cell(ws[a(6, R[2])], 12, False, "000000", shrink=True)
        # ช่วง CF: หัว + ขอบน้ำเงิน
        navy += ["%s:%s" % (a(1, R[1]), a(2, R[1])),
                 "%s:%s" % (a(4, R[1]), a(ncols, R[1])),
                 "%s:%s" % (a(1, R[2]), a(1, R[2])),
                 "%s:%s" % (a(4, R[2]), a(4, R[2])),
                 "%s:%s" % (a(ncols, R[2]), a(ncols, R[2])),
                 "%s:%s" % (a(1, R[3]), a(ncols, R[3]))]
        sep.append(a(3, R[1]))
        guide.append("%s:%s" % (a(1, R[0]), a(ncols, R[0])))

    flag = fl
    cf_add(ws, navy, flag, fill=C_NAVY, font_color="FFFFFF")
    cf_add(ws, sep, flag, fill=C_NAVY, font_color="FFFFFF", border=Border(left=side("BFBFBF")))
    cf_add(ws, guide, flag, border=Border(top=side()))

    setup_page(ws, 0.6, 1.0, last_row, ncols, rows_page)
    return ws


# ---------------------------------------------------------------- ข้อมูลตัวอย่าง (ไฟล์ mockup)
def mock_rows(wb):
    """~16 แถวตัวอย่าง: รหัสที่มีใน รวม ก่อน (เพื่อโชว์ตำแหน่งอัตโนมัติ) แล้วเติมจาก สินค้า"""
    codes = []
    if SHEET_RECORD in wb.sheetnames:
        for (c,) in wb[SHEET_RECORD].iter_rows(min_row=5, max_row=2004, min_col=3, max_col=3,
                                               values_only=True):
            c = (str(c).strip() if c is not None else "")
            if c and c not in codes:
                codes.append(c)
            if len(codes) >= 10:
                break
    prod = wb[SHEET_PRODUCT]
    pcodes = [str(r[0]) for r in prod.iter_rows(min_row=2, max_row=prod.max_row, max_col=2,
                                                values_only=True)
              if r[0] and r[1] and r[0] != "Grand Total"]
    stride = max(1, len(pcodes) // 20)
    for c in pcodes[7::stride]:
        if len(codes) >= 16:
            break
        if c not in codes:
            codes.append(c)
    codes = codes[:16]
    typed_loc = {2: "ชั้นวาง 3-2", 7: "Belt 05", 12: "A-12"}   # พิมพ์ตำแหน่งเองบางใบ
    qtys = {0: 221, 1: 48, 3: 1500, 4: 0, 6: 36, 8: 12, 9: 305, 11: 80, 13: 7}  # ที่เหลือเว้นว่าง
    return [(c, typed_loc.get(i), qtys.get(i)) for i, c in enumerate(codes)]


# ---------------------------------------------------------------- อ่านข้อมูลเก่า
def read_old(wb):
    if SHEET_INPUT not in wb.sheetnames:
        return []
    ws = wb[SHEET_INPUT]
    if ws.cell(row=HDR, column=2).value != "รหัสสินค้า" or ws.cell(row=HDR, column=6).value != "จำนวน":
        print("  คำเตือน: หัวตาราง กรอกป้าย เดิมไม่ตรงรูปแบบ - ข้ามการเก็บข้อมูลเดิม")
        return []
    out = []
    for r in range(FR, FR + N):
        out.append((ws.cell(row=r, column=2).value, ws.cell(row=r, column=4).value,
                    ws.cell(row=r, column=6).value))
    return out


def has_data(rows):
    return any(any(v not in (None, "") for v in row) for row in rows)


# ---------------------------------------------------------------- main
def check_src_safe(src):
    """openpyxl เปิด-บันทึกแล้วรูป/ชาร์ต/shape หาย -> ถ้าไฟล์ต้นทางมีให้หยุดก่อน (ไม่ทำลายงานเดย์) — เหมือน build_lookup.py"""
    with zipfile.ZipFile(src) as z:
        bad = [n for n in z.namelist() if n.startswith(("xl/media/", "xl/drawings/", "xl/charts/"))]
    if bad:
        raise SystemExit(f"ไฟล์ {src} มีรูป/ชาร์ต/shape ({bad[0]} ...) openpyxl จะทำให้หาย -> หยุดไว้ก่อน ไม่ได้เขียนอะไร")


def process(path, force_mock):
    print("== %s" % path)
    if not os.path.exists(path):
        print("  ไม่พบไฟล์ - ข้าม")
        return False
    check_src_safe(path)
    wb = load_workbook(path)
    if SHEET_PRODUCT not in wb.sheetnames:
        print("  ผิดพลาด: ไม่พบชีต \"สินค้า\" - ข้าม")
        return False
    has_record = SHEET_RECORD in wb.sheetnames
    if not has_record:
        print("  คำเตือน: ไม่พบชีต \"รวม\" (รัน build_lookup.py ก่อน) - ตำแหน่งอัตโนมัติจะว่าง")

    old = read_old(wb)
    is_mock = force_mock or "mockup" in os.path.basename(path).lower()
    if is_mock and not has_data(old):
        old = mock_rows(wb)
        print("  ใส่ข้อมูลตัวอย่าง %d แถว" % len(old))
    elif has_data(old):
        print("  เก็บข้อมูลที่พิมพ์ไว้เดิม %d แถว" % sum(1 for o in old if any(v not in (None, "") for v in o)))
    old = (list(old) + [(None, None, None)] * N)[:N]

    active = wb.active
    for name in (SHEET_INPUT, SHEET_S1, SHEET_S2):
        if name in wb.sheetnames:
            wb.remove(wb[name])
    idx = wb.sheetnames.index(SHEET_PRODUCT)
    build_input(wb, idx, old, has_record)
    build_style1(wb, idx + 1)
    build_style2(wb, idx + 2)
    if active.title in wb.sheetnames:
        wb.active = wb.sheetnames.index(active.title)
    wb.calculation.fullCalcOnLoad = True      # ให้ Excel คำนวณสูตรใหม่ทั้งหมดตอนเปิด

    # สำรองไฟล์เดิม ก่อนเขียนทับ
    folder = os.path.dirname(os.path.abspath(path))
    bdir = os.path.join(folder, "backup")
    os.makedirs(bdir, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    base = os.path.splitext(os.path.basename(path))[0]
    shutil.copy2(path, os.path.join(bdir, "%s_labels_%s.xlsx" % (base, stamp)))
    tmp = path + ".tmp"
    wb.save(tmp)
    os.replace(tmp, path)
    print("  เสร็จ: ชีต %s / %s / %s (ก่อนชีต สินค้า)" % (SHEET_INPUT, SHEET_S1, SHEET_S2))
    return True


def main():
    ap = argparse.ArgumentParser(description="สร้างชีตป้ายชั้นวาง (กรอกป้าย / ป้ายแบบ1 / ป้ายแบบ2)")
    ap.add_argument("--file", "-f", action="append", nargs="+", default=None,
                    help="ไฟล์ xlsx (ใส่ได้หลายไฟล์) ค่าเริ่มต้น: LAY_OUT_02.xlsx และ LAY_OUT_02_mockup.xlsx")
    ap.add_argument("--mock", action="store_true",
                    help="ใส่ข้อมูลตัวอย่างใน กรอกป้าย (เฉพาะเมื่อยังว่าง)")
    args = ap.parse_args()
    if args.file:
        files = [p for grp in args.file for p in grp]
    else:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        files = [os.path.join(root, n) for n in ("LAY_OUT_02.xlsx", "LAY_OUT_02_mockup.xlsx")
                 if os.path.exists(os.path.join(root, n))]
        if not files:
            sys.exit("ไม่พบ LAY_OUT_02.xlsx - ระบุด้วย --file")
    ok = [process(p, args.mock) for p in files]
    sys.exit(0 if all(ok) else 1)


if __name__ == "__main__":
    main()
