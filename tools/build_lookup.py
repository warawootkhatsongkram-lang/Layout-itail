#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_lookup.py — ทำสมุดงาน "ของอยู่ที่ไหน" จากไฟล์ที่เดย์แก้เอง (ชีต ผังคลัง / รวม / สินค้า)

วิธีใช้ (จากโฟลเดอร์โปรเจค):
    python3 tools/build_lookup.py --src ไฟล์ที่เดย์แก้.xlsx
    (ถ้าไม่ใส่ --src จะใช้ LAY_OUT_02.xlsx ในโฟลเดอร์โปรเจค -> รันซ้ำกับไฟล์ที่ได้จากสคริปต์นี้ได้ ผลเหมือนเดิม)

ได้ไฟล์ 2 ไฟล์ (เขียนทับในโฟลเดอร์โปรเจค):
    LAY_OUT_02.xlsx          = ไฟล์จริง (ชีตจดว่างเปล่า) ให้ลูกน้องกรอก
    LAY_OUT_02_mockup.xlsx   = ไฟล์ตัวอย่าง ชีต "รวม" ถูกเติมข้อมูลสมมติเต็ม เพื่อดูว่าตอนกรอกครบหน้าตาเป็นยังไง

สิ่งที่สคริปต์ทำกับไฟล์จริง:
    1) ลบชีตรายชั้น (ชื่อตาม layout_config.csv คอลัมน์ชื่อชีต) และ "กล่องพลาสติก (2)"
    2) ชีต "ผังคลัง": ลบเฉพาะลิงก์ที่ชี้ไปชีตที่ถูกลบ (หน้าตา/ค่า/สี/ขนาดเหมือนเดิมทุกอย่าง)
    3) ชีต "รวม": ขยายแถวข้อมูลเป็น 2,000 แถว (แถว 5-2004) ย้ายหมายเหตุท้ายตารางลงไปไว้ใต้แถวสุดท้าย
       ใส่ dropdown เลือก Location (เป็นแบบ "เตือน" -> พิมพ์ค่าอื่นได้ ไม่ถูกบล็อก)
       ** ข้อมูลที่กรอกไว้ในคอลัมน์ B/C/E จะไม่ถูกแตะ **
    4) สร้างชีต "ค้นหา" : พิมพ์รหัส หรือคำในชื่อสินค้า -> ขึ้นทุกแถวที่ตรง พร้อม Location (ไม่ใช้สูตรใหม่ๆ
       เช่น FILTER/XLOOKUP ใช้ได้ตั้งแต่ Excel 2010 และ Google Sheets/LibreOffice)
    5) เรียงชีต: ผังคลัง, ค้นหา, รวม, สินค้า

หลักการทำงานของชีต "ค้นหา" (อ่านแล้วแก้เองได้):
    - คอลัมน์ H (ซ่อน) = "ตัวนับสะสม": ไล่ดูแถวของชีต รวม ทีละแถว ถ้าแถวนั้นมีรหัสและ (รหัส+ชื่อ) มีคำที่พิมพ์ ก็บวก 1
    - ผลลัพธ์ลำดับที่ k = แถวแรกที่ตัวนับ = k  (MATCH(k, ตัวนับ, 0)) แล้วดึงข้อมูลแถวนั้นมาด้วย INDEX
    - ค้นแบบไม่สนตัวพิมพ์เล็ก/ใหญ่ (ใช้ FIND + LOWER ไม่ใช้ SEARCH เพราะ SEARCH ตีความ * ? เป็นอักขระพิเศษ)

หมายเหตุ: ไฟล์ผลลัพธ์ไม่มีค่าที่คำนวณไว้ในไฟล์ (openpyxl เขียนแต่สูตร) Excel จะคำนวณเองตอนเปิด
          ถ้าดูผ่านตัวดูตัวอย่างบนมือถือ/ไดรฟ์แล้วเห็นช่องว่าง ให้เปิดด้วย Excel/Google Sheets
สคริปต์นี้ไม่แตะ LAY_OUT_01.xlsx / layout_config.csv / finder (อ่านอย่างเดียว)
"""
import argparse
import csv
import random
import re
import shutil
import sys
import tempfile
import zipfile
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = Path(__file__).resolve().parent.parent
sys.dont_write_bytecode = True   # ไม่สร้างโฟลเดอร์ __pycache__ ตอน import build_finder
sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_finder  # ใช้ MOCK_KEYWORDS + กติกาสุ่ม seed 42 ชุดเดียวกับเว็บ (อ่านอย่างเดียว ไม่แก้ไฟล์นั้น)

CSV_PATH = ROOT / "layout_config.csv"
OUT_REAL = ROOT / "LAY_OUT_02.xlsx"
OUT_MOCK = ROOT / "LAY_OUT_02_mockup.xlsx"

# ---- ชื่อชีต ----
SHEET_MAP = "ผังคลัง"
SHEET_FIND = "ค้นหา"
SHEET_LOG = "รวม"
SHEET_PROD = "สินค้า"
SHEET_ORDER = [SHEET_MAP, SHEET_FIND, SHEET_LOG, SHEET_PROD]
EXTRA_DELETE = ["กล่องพลาสติก (2)", "รายการ_Location"]   # ชีตอื่นที่ต้องลบนอกจากชีตรายชั้น
MAP_LINK = f"'{SHEET_MAP}'!D11"    # ปลายทางลิงก์ "กลับผัง" (ตำแหน่งเดียวกับที่เดย์ตั้งไว้ในชีต รวม)

# ---- ขนาดตารางในชีต รวม ----
FIRST_ROW = 5          # แถวข้อมูลแรก
N_ROWS = 2000          # จำนวนแถวข้อมูล
LAST_ROW = FIRST_ROW + N_ROWS - 1      # 2004
FOOTER_ROW = LAST_ROW + 2              # 2006 หมายเหตุท้ายตาราง
D_FORMULA = '=IF(C{r}="","",IFERROR(VLOOKUP(C{r}&"",สินค้า!$A:$C,2,FALSE),"ไม่พบรหัส"))'

# ---- ขนาดชีต ค้นหา ----
MAX_SHOW = 200         # แสดงผลสูงสุดกี่รายการ
R_TITLE, R_LABEL, R_INPUT, R_STATUS, R_HEAD = 1, 2, 3, 4, 5
R_RES1 = 6             # แถวผลลัพธ์แรก
R_RESN = R_RES1 + MAX_SHOW - 1         # 205
H_FIRST = 6            # ตัวนับสะสม (คอลัมน์ H) แถว 6 <-> แถว 5 ของชีต รวม (เลื่อน 1)
H_LAST = H_FIRST + N_ROWS - 1          # 2005
LIST_COL = "L"         # รายชื่อ Location สำหรับ dropdown (คอลัมน์ซ่อน)

SAMPLE_QUERY = "เชือก"
MOCK_NOTES = ["เหลือน้อย รอสั่งซื้อ", "กล่องฉีกขาด", "ย้ายมาจากชั้นอื่น", "รอตรวจนับ",
              "ของเสีย รอส่งคืน", "ใช้แล้วบางส่วน", "แยกเก็บ ห้ามปน", "ป้ายหลุด ติดใหม่แล้ว"]


# ===========================================================================
# ส่วนช่วยอ่านข้อมูล
# ===========================================================================
def read_csv_frames():
    """อ่าน layout_config.csv -> (ชื่อกรอบที่เก็บของ ตามลำดับ CSV, ชื่อชีตรายชั้นทั้งหมด)"""
    names, sheets = [], []
    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if (row["ทำชีต"] or "").strip().upper() != "Y":
                continue
            n = (row["ชื่อในผัง"] or "").strip()
            if n and n not in names:
                names.append(n)
            s = (row["ชื่อชีต"] or "").strip()
            if s:
                sheets.append(s)
    return names, sheets


def read_products(wb):
    """อ่านชีตสินค้า -> [[code, name, unit, total], ...] (ข้าม header และ Grand Total)"""
    ws = wb[SHEET_PROD]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[0]
        if code is None or str(code).strip() in ("", "Grand Total"):
            continue
        name = "" if row[1] is None else str(row[1]).strip()
        unit = "" if row[2] is None else str(row[2]).strip()
        total = row[3] if isinstance(row[3], (int, float)) else 0
        out.append([str(code).strip(), name, unit, total])
    return out


def check_src_safe(src):
    """openpyxl เปิด-บันทึกแล้วรูป/ชาร์ตหาย -> ถ้าไฟล์ต้นทางมีรูปให้หยุดก่อน (ไม่ทำลายงานเดย์)"""
    with zipfile.ZipFile(src) as z:
        bad = [n for n in z.namelist() if n.startswith(("xl/media/", "xl/drawings/", "xl/charts/"))]
    if bad:
        raise SystemExit(f"ไฟล์ {src} มีรูป/ชาร์ต/shape ({bad[0]} ...) openpyxl จะทำให้หาย -> หยุดไว้ก่อน ไม่ได้เขียนอะไร")


# ===========================================================================
# 1) ลบชีตรายชั้น + ล้างลิงก์ที่ชี้ไปชีตที่ไม่มีแล้ว
# ===========================================================================
def delete_old_sheets(wb, shelf_sheets):
    """ลบชีตรายชั้น/กล่องพลาสติก (2)/ค้นหาเดิม (สร้างใหม่ทุกครั้ง) -> คืนชื่อชีตที่ลบ"""
    victims = set(shelf_sheets) | set(EXTRA_DELETE) | {SHEET_FIND}
    victims -= {SHEET_MAP, SHEET_LOG, SHEET_PROD}      # กันพลาด ชีตหลัก 3 ตัวห้ามลบ
    deleted = []
    for name in list(wb.sheetnames):
        if name in victims:
            del wb[name]
            deleted.append(name)
    return deleted


def link_sheet_name(location):
    """'ชื่อชีต'!A1 -> ชื่อชีต (ไม่มีชื่อชีต = None)"""
    m = re.match(r"^'((?:[^']|'')+)'!|^([^'!]+)!", location or "")
    if not m:
        return None
    return (m.group(1) or m.group(2)).replace("''", "'")


def remove_dead_links(wb, ws):
    """ลบ hyperlink ที่ชี้ไปชีตที่ไม่มีอยู่แล้ว (ไม่แตะค่า/สี/ฟอนต์) -> (ลบกี่อัน, เหลือกี่อัน)"""
    removed = kept = 0
    for row in ws.iter_rows():
        for c in row:
            h = c.hyperlink
            if h is None:
                continue
            target = link_sheet_name(h.location) if h.location else None
            if target is not None and target not in wb.sheetnames:
                c.hyperlink = None
                removed += 1
            else:
                kept += 1
    return removed, kept


# ===========================================================================
# 2) ชีต รวม : ขยายเป็น 2,000 แถว
# ===========================================================================
def extend_log_sheet(ws):
    """ขยายแถวข้อมูลของชีต รวม เป็นแถว 5-2004 + ย้ายหมายเหตุท้ายตารางไปแถว 2006"""
    # 2.1 หาหมายเหตุท้ายตาราง = แถวที่ merge ครบ A:E อยู่ใต้ตาราง (เดิมคือแถว 36)
    footer = None
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= FIRST_ROW and rng.min_col == 1 and rng.max_col == 5 and rng.min_row == rng.max_row:
            footer = rng
            break
    foot_text = foot_height = None
    foot_styles = []
    if footer is not None:
        r0 = footer.min_row
        foot_text = ws.cell(r0, 1).value
        foot_height = ws.row_dimensions[r0].height
        foot_styles = [copy(ws.cell(r0, c)._style) for c in range(1, 6)]
        ws.unmerge_cells(str(footer))
        for c in range(1, 6):                       # ล้างแถวเดิม (ถ้าอยู่ในช่วงข้อมูลจะถูกแต่งใหม่ด้านล่าง)
            ws.cell(r0, c).value = None
        if r0 <= LAST_ROW:
            ws.row_dimensions[r0].height = None
            for c in range(1, 6):
                ws.cell(r0, c)._style = copy(ws.cell(FIRST_ROW, c)._style)

    # 2.2 แถวต้นแบบ = แถว 5 (สไตล์/ความสูง/รูปแบบตัวเลข)
    tpl_style = [copy(ws.cell(FIRST_ROW, c)._style) for c in range(1, 6)]
    tpl_height = ws.row_dimensions[FIRST_ROW].height

    for r in range(FIRST_ROW, LAST_ROW + 1):
        ws.cell(r, 1).value = r - FIRST_ROW + 1               # A = ลำดับ
        ws.cell(r, 4).value = D_FORMULA.format(r=r)           # D = สูตรหาชื่อสินค้า
        for c in range(1, 6):
            cell = ws.cell(r, c)
            if cell.border.left.style is None:                # แถวใหม่/แถวที่ยังไม่มีเส้น -> ใช้สไตล์แถว 5
                cell._style = copy(tpl_style[c - 1])
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = tpl_height
        # B/C/E ไม่แตะค่าเลย (ข้อมูลที่กรอกไว้อยู่ครบ)

    # 2.3 วางหมายเหตุท้ายตารางใต้แถวสุดท้าย
    if foot_text is not None:
        for c in range(1, 6):
            ws.cell(FOOTER_ROW, c)._style = copy(foot_styles[c - 1])
        ws.cell(FOOTER_ROW, 1).value = foot_text
        ws.merge_cells(start_row=FOOTER_ROW, start_column=1, end_row=FOOTER_ROW, end_column=5)
        if foot_height is not None:
            ws.row_dimensions[FOOTER_ROW].height = foot_height

    # 2.4 พิมพ์: ขยายพื้นที่พิมพ์ + ให้ขึ้นหลายหน้า (fitToHeight=0) หัวตารางซ้ำทุกหน้า (print title เดิมคงไว้)
    ws.print_area = f"A1:E{FOOTER_ROW if foot_text is not None else LAST_ROW}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    # freeze panes (A5) ไม่ถูกแตะ


def add_location_dropdown(ws, list_ref):
    """dropdown เลือก Location คอลัมน์ B (แบบเตือน: พิมพ์ค่าอื่นได้)"""
    keep = []
    for dv in ws.data_validations.dataValidation:      # ล้างของเดิมที่เราเคยใส่ (กันซ้ำตอนรันซ้ำ)
        if not (dv.type == "list" and "B5" in str(dv.sqref)):
            keep.append(dv)
    ws.data_validations.dataValidation = keep
    dv = DataValidation(type="list", formula1=list_ref, allow_blank=True,
                        showErrorMessage=True, errorStyle="warning",
                        errorTitle="Location ไม่ตรงกับรายการ",
                        error="ชื่อนี้ไม่อยู่ในรายการ Location ของผัง\nกด 'ใช่' เพื่อใช้ชื่อที่พิมพ์ต่อ หรือ 'ไม่' เพื่อแก้")
    dv.add(f"B{FIRST_ROW}:B{LAST_ROW}")
    ws.add_data_validation(dv)


# ===========================================================================
# 3) ชีต ค้นหา
# ===========================================================================
def build_find_sheet(wb, loc_names, font_name):
    ws = wb.create_sheet(SHEET_FIND)
    ws.sheet_properties.tabColor = "FF70AD47"
    ws.sheet_view.showGridLines = False
    thin = Side(style="thin", color="FF808080")
    thick = Side(style="medium", color="FF7F6000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    def F(**k):
        return Font(name=font_name, **k)

    for col, w in zip("ABCDE", (7, 20, 17, 40, 22)):     # กว้างพอดีจอมือถือ เลื่อนดูได้
        ws.column_dimensions[col].width = w

    # --- หัวเรื่อง + ลิงก์กลับผัง ---
    ws.merge_cells("A1:D1")
    ws["A1"] = "ค้นหาว่าของอยู่ที่ไหน"
    ws["A1"].font = F(sz=18, b=True)
    ws["A1"].alignment = Alignment(vertical="center")
    ws["E1"] = "← ผังคลัง"
    ws["E1"].hyperlink = Hyperlink(ref="E1", location=MAP_LINK, display="← ผังคลัง")
    ws["E1"].font = F(sz=11, b=True, u="single", color="FF1F3864")
    ws["E1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[R_TITLE].height = 34

    # --- ช่องพิมพ์ (สีเหลืองใหญ่) ---
    ws.merge_cells("A2:E2")
    ws["A2"] = "พิมพ์รหัส หรือ คำในชื่อสินค้า"
    ws["A2"].font = F(sz=13, b=True)
    ws["A2"].alignment = Alignment(vertical="bottom")
    ws.row_dimensions[R_LABEL].height = 24
    ws.merge_cells("A3:E3")
    for c in range(1, 6):
        cell = ws.cell(R_INPUT, c)
        cell.fill = PatternFill("solid", fgColor="FFFFEB84")
        cell.border = Border(left=thick if c == 1 else None, right=thick if c == 5 else None, top=thick, bottom=thick)
    ws["A3"].font = F(sz=20, b=True)
    ws["A3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws["A3"].number_format = "@"          # พิมพ์รหัสตัวเลขแล้วไม่ถูกแปลงเป็นตัวเลข
    ws.row_dimensions[R_INPUT].height = 42

    # --- บรรทัดสถานะ "พบ N รายการ" ---
    ws.merge_cells("A4:E4")
    ws["A4"] = ('=IF($J$1="","พิมพ์รหัส หรือคำในชื่อสินค้าในช่องสีเหลือง เช่น เชือก",'
                'IF($J$2=0,"ไม่พบรายการ ลองพิมพ์คำอื่น หรือพิมพ์ให้สั้นลง",'
                f'"พบ "&$J$2&" รายการ"&IF($J$2>{MAX_SHOW},"  (แสดง {MAX_SHOW} รายการแรก พิมพ์เพิ่มเพื่อกรอง)","")))')
    ws["A4"].font = F(sz=13, b=True, color="FF375623")
    ws["A4"].alignment = Alignment(vertical="center")
    ws.row_dimensions[R_STATUS].height = 26

    # --- หัวตาราง (สไตล์เดียวกับชีต รวม) ---
    for c, t in enumerate(["ลำดับ", "Location", "รหัสสินค้า", "รายการ", "หมายเหตุ"], start=1):
        cell = ws.cell(R_HEAD, c, t)
        cell.font = F(sz=10, b=True)
        cell.fill = PatternFill("solid", fgColor="FFD9D9D9")
        cell.border = box
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[R_HEAD].height = 26

    # --- ช่วยคำนวณ (คอลัมน์ซ่อน H, I, J, L) ---
    ws["J1"] = f"=LOWER(TRIM($A${R_INPUT}))"          # คำค้นที่ทำให้เรียบร้อยแล้ว
    ws["J2"] = f"=$H${H_LAST}"                         # จำนวนที่ตรงทั้งหมด = ค่าตัวนับตัวสุดท้าย
    for i in range(N_ROWS):
        hr = H_FIRST + i               # แถวของตัวนับ
        lr = FIRST_ROW + i             # แถวที่ตรงกันในชีต รวม
        test = (f'IF(AND($J$1<>"",{SHEET_LOG}!C{lr}<>"",'
                f'ISNUMBER(FIND($J$1,TRIM(LOWER({SHEET_LOG}!C{lr}&" "&{SHEET_LOG}!D{lr}))))),1,0)')
        ws.cell(hr, 8).value = f"={test}" if i == 0 else f"=H{hr - 1}+{test}"
    for k in range(MAX_SHOW):
        r = R_RES1 + k
        ws.cell(r, 1).value = f'=IF(ROWS(A${R_RES1}:A{r})<=MIN($J$2,{MAX_SHOW}),ROWS(A${R_RES1}:A{r}),"")'
        # I = แถวใน H ที่ตัวนับเท่ากับ k ครั้งแรก = แถวของรายการที่ k
        ws.cell(r, 9).value = f'=IF($A{r}="","",MATCH($A{r},$H${H_FIRST}:$H${H_LAST},0))'
        for c, col in zip((2, 3, 4, 5), "BCDE"):
            ws.cell(r, c).value = f'=IF($I{r}="","",INDEX({SHEET_LOG}!{col}${FIRST_ROW}:{col}${LAST_ROW},$I{r})&"")'
        for c in range(1, 6):
            cell = ws.cell(r, c)
            cell.font = F(sz=11, b=(c == 2), color=("FF595959" if c == 1 else None))
            cell.alignment = Alignment(horizontal="center" if c in (1, 2) else None, vertical="center",
                                       wrap_text=(c in (2, 4, 5)))
        ws.row_dimensions[r].height = 24

    # เส้นขอบตารางผลลัพธ์: ขึ้นเฉพาะแถวที่มีผล (แถวว่างจะไม่มีกรอบรก/ไม่พิมพ์กรอบเปล่า)
    ws.conditional_formatting.add(
        f"A{R_RES1}:E{R_RESN}",
        FormulaRule(formula=[f'$A{R_RES1}<>""'], border=box))

    # รายชื่อ Location สำหรับ dropdown ของชีต รวม
    for i, n in enumerate(loc_names, start=1):
        ws[f"{LIST_COL}{i}"] = n
    for col in ("H", "I", "J", LIST_COL):
        ws.column_dimensions[col].hidden = True

    # --- ตั้งค่าพิมพ์ + แช่แข็งหัวตาราง ---
    ws.freeze_panes = f"A{R_RES1}"
    ws.print_area = f"A1:E{R_RESN}"
    ws.print_title_rows = f"{R_HEAD}:{R_HEAD}"
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    return ws


# ===========================================================================
# 4) ข้อมูลสมมติสำหรับไฟล์ mockup
# ===========================================================================
def make_mock_rows(products):
    """สินค้าทุกตัว 1 ที่ + ~10% มีที่ที่สอง -> [(location, code, note)] เรียงตามลำดับกรอบในผัง
    ที่หลักใช้กติกาเดียวกับเว็บ finder (MOCK_KEYWORDS + Random(42)) ที่รองก็สุ่มจากกรอบชั้นวาง/ตัวอักษร"""
    frames, _decor = build_finder.read_frames()
    id2name = {f["id"]: f["name"] for f in frames}
    order = {f["id"]: i for i, f in enumerate(frames)}
    locs, _matched, _rand = build_finder.make_mock_locations(frames, products)   # ที่หลัก (ตรงกับเว็บ)
    pool = [f for f in frames if f["store"] and f["type"] in build_finder.RANDOM_TYPES]

    share = 0.10
    while True:
        rng = random.Random(43)
        rows = [(loc["frame"], loc["code"]) for loc in locs]
        for loc in locs:
            if rng.random() < share:
                other = rng.choice([f for f in pool if f["id"] != loc["frame"]])
                rows.append((other["id"], loc["code"]))
        if len(rows) <= N_ROWS or share <= 0:
            break
        share = max(0.0, share - 0.02)          # เกิน 2,000 แถว -> ลดสัดส่วนที่ที่สอง
    rows.sort(key=lambda t: (order[t[0]], t[1]))
    rng_note = random.Random(7)
    out = []
    for fid, code in rows:
        note = rng_note.choice(MOCK_NOTES) if rng_note.random() < 0.04 else None
        out.append((id2name[fid], code, note))
    return out


def fill_mock(wb, products):
    """เติมชีต รวม ด้วยข้อมูลสมมติ + ใส่คำค้นตัวอย่าง + ป้ายแดง 'ตัวอย่างข้อมูลสมมติ'"""
    rows = make_mock_rows(products)
    if len(rows) > N_ROWS:
        raise SystemExit(f"ข้อมูลสมมติ {len(rows)} แถว เกิน {N_ROWS}")
    ws = wb[SHEET_LOG]
    for i in range(N_ROWS):                      # ล้าง B/C/E ก่อน (ไฟล์ตัวอย่างไม่เอาข้อมูลจริงปนเข้ามา)
        r = FIRST_ROW + i
        for c in (2, 3, 5):
            ws.cell(r, c).value = None
    for i, (loc, code, note) in enumerate(rows):
        r = FIRST_ROW + i
        ws.cell(r, 2).value = loc
        ws.cell(r, 3).value = str(code)
        ws.cell(r, 3).number_format = "@"
        if note:
            ws.cell(r, 5).value = note

    fn = ws["A2"].font.name
    red = InlineFont(rFont=fn, sz=11, b=True, color="FFFF0000")
    base = InlineFont(rFont=fn, sz=11)
    ws["A2"].value = CellRichText(TextBlock(base, "ประเภท: สินค้า      "),
                                  TextBlock(red, "** ตัวอย่างข้อมูลสมมติ ไม่ใช่ข้อมูลจริง **"))
    wf = wb[SHEET_FIND]
    wf["A3"].value = SAMPLE_QUERY
    ft = wf["A1"].font.name
    base2 = InlineFont(rFont=ft, sz=18, b=True)
    small = InlineFont(rFont=ft, sz=12, b=True, color="FFFF0000")
    wf["A1"].value = CellRichText(TextBlock(base2, "ค้นหาว่าของอยู่ที่ไหน  "), TextBlock(small, "(ตัวอย่างข้อมูลสมมติ)"))
    return len(rows)


# ===========================================================================
# ตรวจ: ไม่มีอะไรอ้างถึงชีตที่ลบ
# ===========================================================================
def scan_dead_refs(wb, deleted):
    """นับจุดที่ยังอ้างชีตที่ถูกลบ: สูตร / hyperlink / defined names / data validation / print titles"""
    hits = []
    pats = [(d, re.compile(r"(?:'" + re.escape(d.replace("'", "''")) + r"'|(?<![\w'])" + re.escape(d) + r")!"))
            for d in deleted]
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    for d, p in pats:
                        if p.search(v):
                            hits.append((ws.title, c.coordinate, "formula", d))
                if c.hyperlink is not None:
                    t = (c.hyperlink.location or "") + (c.hyperlink.target or "")
                    for d, p in pats:
                        if p.search(t):
                            hits.append((ws.title, c.coordinate, "link", d))
        for dv in ws.data_validations.dataValidation:
            for d, p in pats:
                if p.search(str(dv.formula1) + str(dv.formula2)):
                    hits.append((ws.title, str(dv.sqref), "validation", d))
        for attr in (ws.print_area, ws.print_title_rows, ws.print_title_cols):
            for d, p in pats:
                if attr and p.search(str(attr)):
                    hits.append((ws.title, str(attr), "print", d))
    for name, dn in wb.defined_names.items():
        for d, p in pats:
            if p.search(str(dn.attr_text)):
                hits.append(("-", name, "definedName", d))
    return hits


# ===========================================================================
# main
# ===========================================================================
def finalize_view(wb):
    """เรียงชีต + ให้เปิดมาที่ชีต ค้นหา (เลือกแท็บเดียวพอ)"""
    wb._sheets = [wb[n] for n in SHEET_ORDER] + [s for s in wb._sheets if s.title not in SHEET_ORDER]
    for s in wb.worksheets:
        s.sheet_view.tabSelected = False
    wb.active = wb.sheetnames.index(SHEET_FIND)
    wb[SHEET_FIND].sheet_view.tabSelected = True
    wb.calculation.fullCalcOnLoad = True            # ให้ Excel คำนวณสูตรทั้งหมดตอนเปิด


def build_real(src):
    wb = openpyxl.load_workbook(src)
    for need in (SHEET_MAP, SHEET_LOG, SHEET_PROD):
        if need not in wb.sheetnames:
            raise SystemExit(f"ไม่พบชีต '{need}' ในไฟล์ {src}")
    loc_names, shelf_sheets = read_csv_frames()

    # ชื่อ Location ใน dropdown ต้องตรงกับข้อความที่อยู่บนผังจริงของเดย์ (ถ้าไม่ตรง ข้าม dropdown)
    texts = {str(c.value).strip() for row in wb[SHEET_MAP].iter_rows() for c in row if isinstance(c.value, str)}
    missing = [n for n in loc_names if n not in texts]
    use_dropdown = not missing

    deleted = delete_old_sheets(wb, shelf_sheets)
    removed, kept = remove_dead_links(wb, wb[SHEET_MAP])
    log = wb[SHEET_LOG]
    lh, lk = remove_dead_links(wb, log)           # ลิงก์ "← กลับผัง" ใน รวม (ชี้ผังคลัง ซึ่งยังอยู่ = คงไว้)
    extend_log_sheet(log)
    font_name = log["A5"].font.name or "Tahoma"
    build_find_sheet(wb, loc_names if use_dropdown else [], font_name)
    if use_dropdown:
        add_location_dropdown(log, f"'{SHEET_FIND}'!${LIST_COL}$1:${LIST_COL}${len(loc_names)}")
    finalize_view(wb)

    dead = scan_dead_refs(wb, [d for d in deleted if d != SHEET_FIND])
    return wb, dict(deleted=len(deleted), removed=removed, kept=kept, log_removed=lh, log_kept=lk,
                    dropdown=use_dropdown, missing=missing, n_loc=len(loc_names), dead=dead)


def main():
    ap = argparse.ArgumentParser(description="สร้าง LAY_OUT_02.xlsx (ว่าง) + LAY_OUT_02_mockup.xlsx (ข้อมูลสมมติ)")
    ap.add_argument("--src", default=str(OUT_REAL), help="ไฟล์ต้นทาง (ค่าเริ่มต้น LAY_OUT_02.xlsx ในโฟลเดอร์โปรเจค)")
    args = ap.parse_args()
    src = Path(args.src).resolve()
    if not src.exists():
        raise SystemExit(f"ไม่พบไฟล์ {src}")
    check_src_safe(src)

    # ทำงานบนสำเนาชั่วคราว (ไม่แตะไฟล์ต้นทาง; ถ้าต้นทางคือไฟล์ที่จะเขียนทับ ก็อ่านจบก่อนค่อยเขียน)
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp) / "src.xlsx"
        shutil.copyfile(src, work)
        wb, info = build_real(work)
    if info["dead"]:
        raise SystemExit(f"ยังมีจุดอ้างชีตที่ลบแล้ว {len(info['dead'])} จุด เช่น {info['dead'][:3]}")
    wb.save(OUT_REAL)
    print(f"เขียน {OUT_REAL.name} ({OUT_REAL.stat().st_size / 1024:.0f} KB)")

    # ไฟล์ตัวอย่าง: เปิดไฟล์จริงที่เพิ่งบันทึก แล้วเติมข้อมูลสมมติ
    wb2 = openpyxl.load_workbook(OUT_REAL)
    n = fill_mock(wb2, read_products(wb2))
    finalize_view(wb2)
    wb2.save(OUT_MOCK)
    print(f"เขียน {OUT_MOCK.name} ({OUT_MOCK.stat().st_size / 1024:.0f} KB) ข้อมูลสมมติ {n} แถว")

    print(f"ลบชีต {info['deleted']} ชีต | ลิงก์ที่ลบจาก ผังคลัง {info['removed']} (เหลือ {info['kept']}) | "
          f"ลิงก์ใน รวม ลบ {info['log_removed']} เหลือ {info['log_kept']}")
    if info["dropdown"]:
        print(f"dropdown Location: ใส่แล้ว ({info['n_loc']} ชื่อ)")
    else:
        print(f"dropdown Location: ข้าม เพราะชื่อเหล่านี้ไม่ตรงกับข้อความบนผัง: {info['missing'][:5]}")


if __name__ == "__main__":
    sys.exit(main())
