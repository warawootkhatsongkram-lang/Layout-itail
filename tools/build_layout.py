#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_layout.py - สร้าง LAY_OUT_02.xlsx (ผังคลังสินค้า + ใบจดสินค้ารายชั้นวาง)

วิธีใช้ (รันจาก root โปรเจค):
    1. แก้ layout_config.csv  (เปิดด้วย Excel ได้ บันทึกเป็น CSV UTF-8)
    2. python3 tools/build_layout.py
    3. ได้ไฟล์ LAY_OUT_02.xlsx ใหม่ (เขียนทับของเดิม)

ความปลอดภัยของข้อมูลที่ลูกน้องกรอกไว้ (รันซ้ำได้ ข้อมูลไม่หาย):
    - ก่อนเขียนทับ ถ้ามี LAY_OUT_02.xlsx เดิม จะสำรองไว้ที่ backup/LAY_OUT_02_ปปปปดดวว-ชชนนวว.xlsx เสมอ
    - อ่านข้อมูลที่กรอกในใบจดเดิม (ระดับชั้น/รหัส/จำนวน/หมายเหตุ แถว 6-35 + ผู้จด/วันที่จด)
      แล้วเติมกลับในใบจดชื่อชีตเดียวกันของไฟล์ใหม่ (คอลัมน์ รายการ/หน่วย เป็นสูตร ไม่แตะ)
    - ใบจดที่มีข้อมูลแต่ถูกลบ/เปลี่ยนชื่อใน CSV แล้ว จะไม่ทิ้ง: รวมไว้ในชีต
      "ข้อมูลค้าง (ชั้นที่ถูกลบ)" (ก่อนชีตสินค้า) และเตือนใน console
    - ถ้าเปลี่ยนชื่อชั้นแล้วอยากให้ข้อมูลตามไป ให้ย้ายข้อมูลจากชีตข้อมูลค้างกลับเอง

ตัวเลือกเสริม (ปกติไม่ต้องใช้ ใช้ทดสอบ):
    python3 tools/build_layout.py --config path/to/config.csv --out path/to/out.xlsx

อินพุต : layout_config.csv  + ชีต "สินค้า" จาก LAY_OUT_01.xlsx (อ่านอย่างเดียว ไม่แก้ไฟล์ต้นฉบับ)
เอาต์พุต: LAY_OUT_02.xlsx  = ชีต "ผังคลัง" + ใบจด 1 ชีตต่อ 1 กรอบ + ชีต "สินค้า"

คอลัมน์ใน layout_config.csv:
    ลำดับ      ลำดับของชีตใบจด (เรียงตามลำดับในไฟล์)
    ชื่อในผัง   ข้อความที่แสดงในกรอบ
    ชื่อชีต     ชื่อแท็บของใบจด (<=31 ตัวอักษร, ไม่ซ้ำ) ถ้าเว้นว่างและทำชีต=Y สคริปต์ตั้งให้อัตโนมัติ
    ประเภท     ชั้นวาง / ตัวอักษร / สินค้า / ป้าย  (กำหนดสีเริ่มต้น)
    ทำชีต       Y = สร้างใบจด + ลิงก์ในผัง / N = เป็นป้ายเฉยๆ
    แถว        แถวบนชีตผัง เช่น 12  (ถ้าต้อง merge หลายแถวเขียน 58~60)
    คอลัมน์    B=ผนังซ้าย D=กลาง F=ผนังขวา  (ถ้าต้อง merge หลายคอลัมน์เขียน B:F)
    สีพื้น      hex 6 หลัก เช่น FFC000 (เว้นว่าง = ใช้สีตามประเภท)

เปลี่ยนเลย์เอาต์ = แก้แถว/คอลัมน์ใน CSV แล้วรันใหม่ ไม่ต้องแตะโค้ด
"""
import argparse
import csv
import os
import re
import shutil
import sys
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins

# ---------- ตำแหน่งไฟล์ (อิงจาก root โปรเจค = โฟลเดอร์แม่ของ tools/) ----------
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_CSV = os.path.join(ROOT, "layout_config.csv")
SRC_XLSX = os.path.join(ROOT, "LAY_OUT_01.xlsx")
OUT_XLSX = os.path.join(ROOT, "LAY_OUT_02.xlsx")

MAP_NAME = "ผังคลัง"      # ชื่อชีตผัง (ชีตแรก)
STRAY_NAME = "ข้อมูลค้าง (ชั้นที่ถูกลบ)"   # ชีตเก็บข้อมูลของใบจดที่ไม่มีใน CSV แล้ว
DEFAULT_WHO = "ผู้จด: ______________________"      # ข้อความเริ่มต้นช่องผู้จด
DEFAULT_WHEN = "วันที่จด: ____/____/______"         # ข้อความเริ่มต้นช่องวันที่จด
PRODUCT_NAME = "สินค้า"   # ชื่อชีตสินค้า (ชีตสุดท้าย และชีตต้นทางใน LAY_OUT_01.xlsx)

# ---------- ค่าคงที่หน้าตา ----------
FONT_NAME = "Tahoma"      # ฟอนต์เดียวทั้งไฟล์
FRAME_COLS = ("B", "D", "F")   # คอลัมน์กรอบ (ผนังซ้าย / กลาง / ผนังขวา)
AISLE_COLS = ("C", "E")        # คอลัมน์ทางเดิน (แคบ)
FRAME_W, AISLE_W, ROW_H = 24, 4, 28
GRID_TOP = 4                   # แถวแรกของกรอบ
COL_LABELS = {"B": "ผนังซ้าย", "D": "กลาง", "F": "ผนังขวา"}

# สีเริ่มต้นตามประเภท (พื้นกรอบ / สีแท็บใบจดที่เข้มขึ้นเล็กน้อย)
TYPE_FILL = {"ชั้นวาง": "DDEBF7", "ตัวอักษร": "EDEDED", "สินค้า": "FFF2CC", "ป้าย": "FFFFFF"}
TYPE_TAB = {"ชั้นวาง": "9DC3E6", "ตัวอักษร": "BFBFBF", "สินค้า": "FFD966", "ป้าย": "808080"}
LINK_COLOR = "1F3864"          # น้ำเงินเข้มสำหรับลิงก์

ROWS_DATA = 30                 # จำนวนแถวข้อมูลในใบจด
HDR_ROW = 5
FIRST = HDR_ROW + 1            # แถวข้อมูลแรก = 6
LAST = HDR_ROW + ROWS_DATA     # แถวข้อมูลสุดท้าย = 35

THIN = Side(style="thin", color="000000")
MED = Side(style="medium", color="000000")


def font(**kw):
    kw.setdefault("name", FONT_NAME)
    kw.setdefault("size", 10)
    return Font(**kw)


def fill(hex6):
    return PatternFill("solid", start_color=hex6, end_color=hex6)


# ---------- ตั้งชื่อชีต ----------
def sanitize_sheet_name(name, suffix=""):
    """ชื่อในผัง -> ชื่อชีตที่ Excel ยอมรับ (ยุบช่องว่าง, / -> -, * -> x, ตัด \\ ? [ ] :, <=31 ตัว)"""
    s = re.sub(r"\s+", " ", name).strip()
    s = s.replace("/", "-").replace("*", "x")
    s = re.sub(r"[\\?\[\]:]", "", s)
    s = s.strip("'")
    keep = 31 - len(suffix)
    return s[:keep].rstrip() + suffix


# ---------- อ่าน CSV ----------
def parse_span(text, is_col):
    """'12' -> (12,12) | '58~60' -> (58,60) | 'B' -> (2,2) | 'B:F' -> (2,6)"""
    t = text.strip().upper() if is_col else text.strip()
    if is_col:
        parts = t.split(":")
        a = column_index_from_string(parts[0])
        b = column_index_from_string(parts[-1])
    else:
        parts = t.split("~")
        a, b = int(parts[0]), int(parts[-1])
    if b < a:
        a, b = b, a
    return a, b


def read_config():
    items = []
    with open(CONFIG_CSV, encoding="utf-8-sig", newline="") as f:
        for n, r in enumerate(csv.DictReader(f), start=2):
            name = (r.get("ชื่อในผัง") or "").strip()
            if not name:
                continue  # ข้ามบรรทัดว่าง
            try:
                r1, r2 = parse_span(r["แถว"], False)
                c1, c2 = parse_span(r["คอลัมน์"], True)
            except Exception:
                raise SystemExit(f"CSV บรรทัด {n}: แถว/คอลัมน์ อ่านไม่ได้ ({name})")
            typ = (r.get("ประเภท") or "").strip()
            if typ not in TYPE_FILL:
                raise SystemExit(f"CSV บรรทัด {n}: ประเภท '{typ}' ไม่รู้จัก (ใช้ ชั้นวาง/ตัวอักษร/สินค้า/ป้าย)")
            color = (r.get("สีพื้น") or "").strip().lstrip("#").upper()
            if color and not re.fullmatch(r"[0-9A-F]{6}", color):
                raise SystemExit(f"CSV บรรทัด {n}: สีพื้น '{color}' ต้องเป็น hex 6 หลัก")
            items.append({
                "seq": int(r.get("ลำดับ") or 0) or n,
                "name": re.sub(r"\s+", " ", name),
                "sheet": (r.get("ชื่อชีต") or "").strip(),
                "type": typ,
                "make": (r.get("ทำชีต") or "").strip().upper() == "Y",
                "r1": r1, "r2": r2, "c1": c1, "c2": c2,
                "color": color,
            })
    items.sort(key=lambda x: x["seq"])  # เรียงตามลำดับ (sort คงลำดับเดิมถ้าเลขซ้ำ)
    return items


def validate(items):
    """ตรวจชื่อชีตและตำแหน่งซ้อนทับ ก่อนสร้างไฟล์"""
    seen_sheet = {MAP_NAME.lower(), PRODUCT_NAME.lower(), STRAY_NAME.lower()}
    for it in items:
        if not it["make"]:
            continue
        if not it["sheet"]:
            it["sheet"] = sanitize_sheet_name(it["name"])
        s = it["sheet"]
        if len(s) > 31 or re.search(r"[\\/?*\[\]:]", s) or s.startswith("'") or s.endswith("'"):
            raise SystemExit(f"ชื่อชีต '{s}' ไม่ถูกต้อง (เกิน 31 ตัว หรือมีอักขระต้องห้าม)")
        if s.lower() in seen_sheet:
            raise SystemExit(f"ชื่อชีต '{s}' ซ้ำ (หรือชนกับชีต ผังคลัง/สินค้า)")
        seen_sheet.add(s.lower())
    used = {}
    for it in items:
        for r in range(it["r1"], it["r2"] + 1):
            for c in range(it["c1"], it["c2"] + 1):
                if (r, c) in used:
                    raise SystemExit(
                        f"ตำแหน่งซ้ำ {get_column_letter(c)}{r}: '{it['name']}' ชนกับ '{used[(r, c)]}'")
                used[(r, c)] = it["name"]
        if it["r1"] < GRID_TOP or get_column_letter(it["c1"]) not in FRAME_COLS \
                or get_column_letter(it["c2"]) not in FRAME_COLS:
            raise SystemExit(f"'{it['name']}' อยู่นอกกริด (แถวต้อง >= {GRID_TOP}, คอลัมน์ต้องเป็น {'/'.join(FRAME_COLS)})")


# ---------- สร้างชีตผัง ----------
def link(cell, sheet, ref="A1", color=LINK_COLOR):
    """ใส่ hyperlink ภายในไฟล์ไปยังชีต"""
    loc = "'" + sheet.replace("'", "''") + "'!" + ref
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=loc, display=str(cell.value))


def build_map(ws, items):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "404040"
    ws.column_dimensions["A"].width = 2
    for c in FRAME_COLS:
        ws.column_dimensions[c].width = FRAME_W
    for c in AISLE_COLS:
        ws.column_dimensions[c].width = AISLE_W
    ws.column_dimensions["G"].width = 3
    ws.column_dimensions["H"].width = 30

    # หัวเรื่อง + คำแนะนำ + ชื่อคอลัมน์ผนัง
    ws["B1"] = "ผังคลังสินค้า"
    ws["B1"].font = font(size=18, bold=True)
    ws["B2"] = "กดที่ชื่อกรอบเพื่อไปยังใบจดของชั้นนั้น"
    ws["B2"].font = font(size=10, italic=True, color="595959")
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 18
    for c, t in COL_LABELS.items():
        ws[f"{c}3"] = t
        ws[f"{c}3"].font = font(bold=True, color="595959")
        ws[f"{c}3"].alignment = Alignment(horizontal="center")

    # วางกรอบตามแถว/คอลัมน์ใน CSV
    max_row = GRID_TOP
    cells = {}
    for it in items:
        c1, c2, r1, r2 = it["c1"], it["c2"], it["r1"], it["r2"]
        fc = it["color"] or TYPE_FILL[it["type"]]
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cl = ws.cell(row=r, column=c)
                cl.fill = fill(fc)
                cl.border = Border(left=MED if c == c1 else None, right=MED if c == c2 else None,
                                   top=MED if r == r1 else None, bottom=MED if r == r2 else None)
        if (r1, c1) != (r2, c2):
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        cell = ws.cell(row=r1, column=c1, value=it["name"])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if it["make"]:
            cell.font = font(underline="single", color=LINK_COLOR)
            link(cell, it["sheet"])
            cells[it["sheet"]] = cell.coordinate
        else:
            cell.font = font(bold=True, size=11 if it["name"].lower() == "office" else 10)
        max_row = max(max_row, r2)
    for r in range(GRID_TOP, max_row + 1):
        ws.row_dimensions[r].height = ROW_H

    # legend คอลัมน์ H
    ws["H4"] = "สีของกรอบ"
    ws["H4"].font = font(bold=True)
    legend = [("ชั้นวาง", "ชั้นวาง"), ("ตัวอักษร (โซนตัวอักษร)", "ตัวอักษร"),
              ("สินค้า (กรอบชื่อสินค้า)", "สินค้า"), ("ป้าย / พื้นที่ (ไม่มีใบจด)", "ป้าย")]
    for i, (txt, typ) in enumerate(legend):
        c = ws.cell(row=5 + i, column=8, value=txt)
        c.fill = fill(TYPE_FILL[typ])
        c.border = Border(left=MED, right=MED, top=MED, bottom=MED)
        c.font = font()
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws["H9"] = "สีอื่น = สีตามผังเดิม"
    ws["H10"] = "ตัวอักษรขีดเส้นใต้ = กดเพื่อไปใบจด"
    for a in ("H9", "H10"):
        ws[a].font = font(size=9, color="595959")
        ws[a].alignment = Alignment(wrap_text=True, vertical="top")

    # ตั้งค่าพิมพ์ A4 แนวตั้ง กว้าง 1 หน้า
    ws.freeze_panes = f"A{GRID_TOP}"
    ws.print_area = f"A1:H{max(max_row, 10)}"
    ws.print_title_rows = "1:3"
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_options.horizontalCentered = True
    return cells


# ---------- สร้างใบจด ----------
def build_sheet(ws, it, back_ref):
    ws.sheet_properties.tabColor = TYPE_TAB[it["type"]]
    widths = {"A": 7, "B": 10, "C": 18, "D": 40, "E": 9, "F": 10, "G": 22}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    # หัวใบจด
    ws.merge_cells("A1:F1")
    ws["A1"] = f"ใบจดสินค้า: {it['name']}"
    ws["A1"].font = font(size=16, bold=True)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30
    ws["G1"] = "← กลับผัง"
    ws["G1"].font = font(underline="single", color=LINK_COLOR, bold=True)
    ws["G1"].alignment = Alignment(horizontal="center", vertical="center")
    link(ws["G1"], MAP_NAME, back_ref)

    ws.merge_cells("A2:G2")
    ws["A2"] = f"ประเภท: {it['type']}"
    ws.merge_cells("A3:D3")
    ws["A3"] = DEFAULT_WHO
    ws.merge_cells("E3:G3")
    ws["E3"] = DEFAULT_WHEN
    for a in ("A2", "A3", "E3"):
        ws[a].font = font(size=11)
        ws[a].alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 24
    ws.row_dimensions[4].height = 8

    # หัวตาราง
    heads = ["ลำดับ", "ระดับชั้น", "รหัสสินค้า", "รายการ", "หน่วย", "จำนวน", "หมายเหตุ"]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(row=HDR_ROW, column=i, value=h)
        c.font = font(bold=True)
        c.fill = fill("D9D9D9")
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[HDR_ROW].height = 26

    # แถวข้อมูล: รายการ/หน่วย ดึงจากชีตสินค้าด้วย VLOOKUP
    q = "'" + PRODUCT_NAME + "'!$A:$C"
    for n, r in enumerate(range(FIRST, LAST + 1), start=1):
        ws.row_dimensions[r].height = 24
        ws.cell(row=r, column=1, value=n)
        ws.cell(row=r, column=4,
                value=f'=IF(C{r}="","",IFERROR(VLOOKUP(C{r}&"",{q},2,FALSE),"ไม่พบรหัส"))')
        ws.cell(row=r, column=5,
                value=f'=IF(C{r}="","",IFERROR(VLOOKUP(C{r}&"",{q},3,FALSE),""))')
        for col in range(1, 8):
            c = ws.cell(row=r, column=col)
            c.border = border
            c.font = font(size=11)
            if col in (1, 2, 5):
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col == 6:
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = Alignment(vertical="center", wrap_text=(col in (4, 7)))
        ws.cell(row=r, column=3).number_format = "@"   # รหัสเป็น text กัน 5.6E+11
        ws.cell(row=r, column=1).font = font(size=11, color="595959")

    # จำนวน: ต้องเป็นตัวเลข >= 0
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                        showErrorMessage=True, errorStyle="stop",
                        errorTitle="จำนวนไม่ถูกต้อง",
                        error="กรุณากรอกจำนวนเป็นตัวเลขตั้งแต่ 0 ขึ้นไปเท่านั้น")
    ws.add_data_validation(dv)
    dv.add(f"F{FIRST}:F{LAST}")

    # หมายเหตุใต้ตาราง
    ws.merge_cells(start_row=LAST + 2, start_column=1, end_row=LAST + 2, end_column=7)
    ws.cell(row=LAST + 2, column=1,
            value="ระดับชั้น: นับจากล่างขึ้นบน 1,2,3… หรือเขียน พื้น").font = font(size=9, italic=True, color="595959")

    # ตั้งค่าพิมพ์ A4 แนวตั้ง
    ws.print_area = f"A1:G{LAST + 2}"
    ws.print_title_rows = f"{HDR_ROW}:{HDR_ROW}"
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_options.horizontalCentered = True
    ws.freeze_panes = f"A{FIRST}"


# ---------- คัดลอกชีตสินค้า ----------
def build_products(ws):
    """คัดลอกค่าทุกเซลล์จากชีต 'สินค้า' ใน LAY_OUT_01.xlsx (อ่านอย่างเดียว ไม่ save ทับต้นฉบับ)"""
    src = load_workbook(SRC_XLSX, read_only=True)
    try:
        s = src[PRODUCT_NAME]
        n = 0
        for row in s.iter_rows():
            n += 1
            for cell in row:
                if cell.value is None:
                    continue
                c = ws.cell(row=n, column=cell.column, value=cell.value)
                if isinstance(cell.value, str):
                    c.number_format = "@"          # รหัส/ข้อความคงเป็น text
                elif cell.number_format and cell.number_format != "General":
                    c.number_format = cell.number_format
    finally:
        src.close()
    for c in ws[1]:
        c.font = font(bold=True)
        c.fill = fill("D9D9D9")
    for c in ws["A"][1:]:
        c.font = font(size=10)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{n}"
    for k, v in {"A": 18, "B": 50, "C": 8, "D": 14}.items():
        ws.column_dimensions[k].width = v
    ws.sheet_properties.tabColor = "70AD47"
    return n


# ---------- สำรองไฟล์เดิม + ย้ายข้อมูลที่กรอกไว้ ----------
DATA_COLS = (2, 3, 6, 7)   # B ระดับชั้น, C รหัส, F จำนวน, G หมายเหตุ (D/E เป็นสูตร ไม่แตะ)


def backup_existing(out_path):
    """ถ้ามีไฟล์เดิม สำเนาไว้ที่ backup/ (ข้างไฟล์ผลลัพธ์) แล้วคืน path ของ backup"""
    if not os.path.exists(out_path):
        return None
    bdir = os.path.join(os.path.dirname(os.path.abspath(out_path)), "backup")
    os.makedirs(bdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(out_path))[0]
    dst = os.path.join(bdir, f"{stem}_{datetime.now():%Y%m%d-%H%M%S}.xlsx")
    shutil.copy2(out_path, dst)
    return dst


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def read_old_data(out_path):
    """อ่านข้อมูลที่กรอกในไฟล์เดิม -> (data, who, stray)
    data[ชื่อชีต] = {แถว: {คอลัมน์: ค่า}} | who[ชื่อชีต] = (A3, E3) เฉพาะที่ถูกแก้จากค่าเริ่มต้น
    stray = แถวจากชีตข้อมูลค้างรอบก่อน (เก็บต่อ ไม่ให้หาย)"""
    data, who, stray = {}, {}, []
    wb = load_workbook(out_path)
    for ws in wb.worksheets:
        if ws.title in (MAP_NAME, PRODUCT_NAME):
            continue
        if ws.title == STRAY_NAME:
            for r in ws.iter_rows(min_row=2, values_only=True):
                if any(not blank(v) for v in r[:5]):
                    stray.append(list(r[:5]))
            continue
        rows = {}
        for r in range(FIRST, LAST + 1):
            vals = {c: ws.cell(row=r, column=c).value for c in DATA_COLS}
            if any(not blank(v) for v in vals.values()):
                rows[r] = vals
        if rows:
            data[ws.title] = rows
        a3, e3 = ws["A3"].value, ws["E3"].value
        if (a3 and a3 != DEFAULT_WHO) or (e3 and e3 != DEFAULT_WHEN):
            who[ws.title] = (a3, e3)
    wb.close()
    return data, who, stray


def build_stray(ws, rows):
    """ชีตรวมข้อมูลของใบจดที่ไม่มีใน CSV ใหม่แล้ว"""
    ws.sheet_properties.tabColor = "C00000"
    heads = ["ชื่อชีตเดิม", "ระดับชั้น", "รหัส", "จำนวน", "หมายเหตุ"]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = font(bold=True)
        c.fill = fill("F8CBAD")
    for n, r in enumerate(rows, start=2):
        for i, v in enumerate(r, start=1):
            c = ws.cell(row=n, column=i, value=v)
            c.font = font()
            if i == 3:
                c.number_format = "@"
    for k, v in {"A": 30, "B": 12, "C": 20, "D": 10, "E": 40}.items():
        ws.column_dimensions[k].width = v
    ws.freeze_panes = "A2"


def main():
    global CONFIG_CSV
    ap = argparse.ArgumentParser(description="สร้าง LAY_OUT_02.xlsx จาก layout_config.csv")
    ap.add_argument("--config", default=CONFIG_CSV, help="ไฟล์ config CSV (ค่าเริ่มต้น layout_config.csv)")
    ap.add_argument("--out", default=OUT_XLSX, help="ไฟล์ผลลัพธ์ (ค่าเริ่มต้น LAY_OUT_02.xlsx)")
    args = ap.parse_args()
    CONFIG_CSV = args.config
    out_path = args.out

    items = read_config()
    validate(items)

    # 1) สำรองไฟล์เดิมก่อนเสมอ  2) อ่านข้อมูลที่กรอกไว้ (ถ้าอ่านไม่ได้ ไม่เขียนทับ)
    bak = backup_existing(out_path)
    data, who, stray = {}, {}, []
    if bak:
        try:
            data, who, stray = read_old_data(out_path)
        except Exception as e:
            raise SystemExit(f"อ่านไฟล์เดิมไม่ได้ ({e}) - ยังไม่เขียนทับ (สำรองไว้แล้วที่ {bak})")

    wb = Workbook()
    ws_map = wb.active
    ws_map.title = MAP_NAME
    cells = build_map(ws_map, items)
    sheets = [it for it in items if it["make"]]
    names = {it["sheet"] for it in sheets}

    # 3) เติมข้อมูลเดิมกลับใบจดชื่อเดียวกัน (ไม่เติม D/E ที่เป็นสูตร)
    moved_sheets = moved_rows = 0
    for it in sheets:
        ws = wb.create_sheet(it["sheet"])
        build_sheet(ws, it, cells[it["sheet"]])
        if it["sheet"] in data:
            moved_sheets += 1
            for r, vals in data[it["sheet"]].items():
                moved_rows += 1
                for c, v in vals.items():
                    if not blank(v):
                        ws.cell(row=r, column=c, value=v)
        if it["sheet"] in who:
            a3, e3 = who[it["sheet"]]
            if a3:
                ws["A3"] = a3
            if e3:
                ws["E3"] = e3

    # 4) ใบจดเดิมที่มีข้อมูลแต่ไม่มีใน config ใหม่ -> ชีตข้อมูลค้าง
    new_stray = 0
    for old, rows in data.items():
        if old in names:
            continue
        for r in sorted(rows):
            v = rows[r]
            stray.append([old, v[2], v[3], v[6], v[7]])
            new_stray += 1
    for old, (a3, e3) in who.items():
        if old not in names:
            stray.append([old, None, None, None, f"{a3 or ''} {e3 or ''}".strip()])
            new_stray += 1
    if stray:
        build_stray(wb.create_sheet(STRAY_NAME), stray)
    n = build_products(wb.create_sheet(PRODUCT_NAME))
    wb.save(out_path)

    print(f"สำเร็จ: {out_path}")
    print(f"  กรอบในผัง {len(items)} (ใบจด {len(sheets)}, ป้าย {len(items) - len(sheets)}) | ชีตทั้งหมด {len(wb.sheetnames)} | แถวชีตสินค้า {n}")
    print(f"  สำรองไฟล์เดิม: {bak or '(ไม่มีไฟล์เดิม)'}")
    print(f"  ย้ายข้อมูลเดิมกลับ: {moved_sheets} ชีต / {moved_rows} แถว")
    if stray:
        print(f"  !! ข้อมูลค้าง {len(stray)} แถว (ใหม่รอบนี้ {new_stray}) อยู่ในชีต \"{STRAY_NAME}\" - ตรวจและย้ายเอง")
    else:
        print("  ข้อมูลค้าง: 0 แถว")


if __name__ == "__main__":
    main()
