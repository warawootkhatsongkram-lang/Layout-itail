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

คอลัมน์ใน layout_config.csv (ตำแหน่ง/ขนาดเป็น "ช่องกริด" 1 ช่อง = 6pt เป็นสี่เหลี่ยมจัตุรัส):
    ลำดับ      ลำดับของชีตใบจด (เรียงตามลำดับในไฟล์)
    ชื่อในผัง   ข้อความที่แสดงในกรอบ (แถวตกแต่งใช้เป็นคำอธิบายในไฟล์เท่านั้น ไม่แสดงบนผัง)
    ชื่อชีต     ชื่อแท็บของใบจด (<=31 ตัวอักษร, ไม่ซ้ำ) ถ้าเว้นว่างและทำชีต=Y สคริปต์ตั้งให้อัตโนมัติ
    ประเภท     กรอบ: ชั้นวาง / ตัวอักษร / สินค้า / ป้าย  |  ตกแต่ง: ผนัง / วงเล็บ / ลูกศรขึ้น / ลูกศรลง / แถบ
    ทำชีต       Y = สร้างใบจด + ลิงก์ในผัง / N = เป็นป้ายเฉยๆ (แถวตกแต่งใส่ N)
    X, Y       มุมซ้ายบนของรูป เริ่มนับ 1 (X=1 คือคอลัมน์ A, Y=1 คือแถวแรกของกริดใต้หัวเรื่อง)
    กว้าง, สูง   ขนาดเป็นจำนวนช่อง  (กรอบ = merge ช่วง X..X+กว้าง-1, Y..Y+สูง-1)
    สีพื้น      hex 6 หลัก (กรอบเก็บของว่าง = น้ำเงิน 4472C4 / ป้ายว่าง = พื้นขาวไม่มีขอบ / ตัวอักษรขาวบนสีเข้ม ดำบนสีอ่อน)

วิธีวาดแถวตกแต่ง (ใช้ตำแหน่ง X,Y,กว้าง,สูง เหมือนกรอบ):
    ผนัง      เส้นหนาสีแดง: กว้าง=1 -> เส้นตั้งตามขอบซ้ายของคอลัมน์ X ยาว สูง ช่อง
                            สูง=1 -> เส้นนอนตามขอบบนของแถว Y ยาว กว้าง ช่อง
    วงเล็บ    ปีกกาเส้นบางสีแดง ทิศตามคำในชื่อ: หันซ้าย/หันขวา (ปลายแหลมชี้ซ้าย/ขวา, ตั้ง) หรือ หันขึ้น/หันลง (นอน)
              กว้างควรเป็นเลขคู่ (ตั้ง) / กว้างเป็นเลขคู่และสูงตามต้องการ (นอน)
    ลูกศรขึ้น/ลง  ก้านพื้นแดง + หัว ▲/▼ สีแดง (หัวสูง 3 แถว)
    แถบ       เซลล์พื้นสี (ใช้สีพื้น) มีขอบบาง
    กรอบใหญ่ที่มีกรอบเล็กอยู่ข้างใน (เช่น office ครอบ ชั้นวางหมึก) ทำได้: ข้อความอยู่ในส่วนบนที่ไม่มีกรอบเล็ก
    กรอบอื่นห้ามทับกัน (ชนกัน = error)

เปลี่ยนเลย์เอาต์ = แก้ X/Y/กว้าง/สูง ใน CSV แล้วรันใหม่ ไม่ต้องแตะโค้ด
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
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break

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
CELL_PT = 6                    # 1 ช่องกริดของผัง = 6pt (สี่เหลี่ยมจัตุรัส)
COL_W = 1.09                   # ความกว้างคอลัมน์ผัง (Excel = 8px = 6pt)
HEAD_ROWS = 3                  # แถวหัวเรื่องเหนือกริด (แถว Y=1 ของ CSV = แถวที่ HEAD_ROWS+1 ในชีต)
LEG_X, LEG_W = 69, 18          # คอลัมน์เริ่ม/ความกว้างของ legend (ขวาของผนังขวา)
MAP_BLUE = "4472C4"            # พื้นกรอบเก็บของ (น้ำเงินตามผังเดิม)

# สีแท็บของใบจดตามประเภท
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
FRAME_TYPES = ("ชั้นวาง", "ตัวอักษร", "สินค้า", "ป้าย")             # กรอบ/ป้ายที่เป็นข้อความ
DECOR_TYPES = ("ผนัง", "วงเล็บ", "ลูกศรขึ้น", "ลูกศรลง", "แถบ")      # ตกแต่ง (ไม่มีข้อความ ไม่มีใบจด)
REQUIRED_COLS = ("ลำดับ", "ชื่อในผัง", "ชื่อชีต", "ประเภท", "ทำชีต", "X", "Y", "กว้าง", "สูง", "สีพื้น")


def read_config():
    items = []
    with open(CONFIG_CSV, encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        miss = [k for k in REQUIRED_COLS if k not in (rd.fieldnames or [])]
        if miss:
            raise SystemExit(f"CSV ขาดคอลัมน์ {', '.join(miss)} (schema ใหม่: {','.join(REQUIRED_COLS)})")
        for n, r in enumerate(rd, start=2):
            name = (r.get("ชื่อในผัง") or "").strip()
            if not name:
                continue  # ข้ามบรรทัดว่าง
            typ = (r.get("ประเภท") or "").strip()
            if typ not in FRAME_TYPES + DECOR_TYPES:
                raise SystemExit(f"CSV บรรทัด {n}: ประเภท '{typ}' ไม่รู้จัก "
                                 f"(ใช้ {'/'.join(FRAME_TYPES + DECOR_TYPES)})")
            try:
                x, y, w, h = (int(float(r[k])) for k in ("X", "Y", "กว้าง", "สูง"))
            except Exception:
                raise SystemExit(f"CSV บรรทัด {n}: X/Y/กว้าง/สูง ต้องเป็นตัวเลข ({name})")
            if min(x, y, w, h) < 1:
                raise SystemExit(f"CSV บรรทัด {n}: X/Y/กว้าง/สูง ต้อง >= 1 ({name})")
            color = (r.get("สีพื้น") or "").strip().lstrip("#").upper()
            if color and not re.fullmatch(r"[0-9A-F]{6}", color):
                raise SystemExit(f"CSV บรรทัด {n}: สีพื้น '{color}' ต้องเป็น hex 6 หลัก")
            decor = typ in DECOR_TYPES
            items.append({
                "seq": int(r.get("ลำดับ") or 0) or n,
                "name": re.sub(r"\s+", " ", name),
                "sheet": (r.get("ชื่อชีต") or "").strip(),
                "type": typ,
                "decor": decor,
                "make": (not decor) and (r.get("ทำชีต") or "").strip().upper() == "Y",
                "x": x, "y": y, "w": w, "h": h,
                "color": color,
            })
    items.sort(key=lambda x: x["seq"])  # เรียงตามลำดับ (sort คงลำดับเดิมถ้าเลขซ้ำ)
    return items


def rect_of(it):
    return (it["x"], it["y"], it["x"] + it["w"] - 1, it["y"] + it["h"] - 1)


def rect_hit(a, b):
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def rect_in(a, b):
    """a อยู่ในรูป b ทั้งก้อน"""
    return a[0] >= b[0] and a[1] >= b[1] and a[2] <= b[2] and a[3] <= b[3]


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
    frames = [it for it in items if not it["decor"]]
    # กรอบชนกรอบ: ห้ามทับกัน ยกเว้นกรอบใหญ่ (เช่น office) ที่ครอบกรอบเล็กไว้ทั้งก้อน
    for it in frames:
        it["mrows"] = (it["y"], it["y"] + it["h"] - 1)       # แถวที่ merge (ค่าปกติ = ทั้งกรอบ)
    for i, a in enumerate(frames):
        for b in frames[i + 1:]:
            ra, rb = rect_of(a), rect_of(b)
            if not rect_hit(ra, rb):
                continue
            if rect_in(rb, ra):
                big, small = a, b
            elif rect_in(ra, rb):
                big, small = b, a
            else:
                raise SystemExit(f"ตำแหน่งซ้ำ: '{a['name']}' ชนกับ '{b['name']}'")
            if small["y"] <= big["y"]:
                raise SystemExit(f"'{small['name']}' ซ้อนใน '{big['name']}' ต้องอยู่ต่ำกว่าแถวบนสุดของกรอบใหญ่")
            big["mrows"] = (big["y"], min(big["mrows"][1], small["y"] - 1))  # merge เฉพาะส่วนบนที่ไม่มีกรอบเล็ก
    # ลูกศร/แถบ เป็นเซลล์พื้นสี: ห้ามทับกรอบ (เส้นผนัง/วงเล็บ เป็นแค่เส้นขอบ จึงไม่ตรวจ)
    for d in items:
        if d["type"] in ("ลูกศรขึ้น", "ลูกศรลง", "แถบ"):
            for fr in frames:
                if rect_hit(rect_of(d), rect_of(fr)):
                    raise SystemExit(f"'{d['name']}' ({d['type']}) ทับกรอบ '{fr['name']}'")


# ---------- สร้างชีตผัง ----------
def link(cell, sheet, ref="A1", color=LINK_COLOR):
    """ใส่ hyperlink ภายในไฟล์ไปยังชีต"""
    loc = "'" + sheet.replace("'", "''") + "'!" + ref
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=loc, display=str(cell.value))


def rel_lum(hex6):
    """ความสว่างสัมพัทธ์ (WCAG) ของสี hex"""
    def ch(v):
        v /= 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hex6[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def shade(hex6, k=0.55):
    """สีเข้มขึ้น (ใช้เป็นสีขอบกรอบ)"""
    return "".join(f"{int(int(hex6[i:i + 2], 16) * k):02X}" for i in (0, 2, 4))


def text_width(text, size):
    """ประมาณความกว้างข้อความ (pt) ของ Tahoma ตัวหนา: สระ/วรรณยุกต์ไทยบน-ล่างไม่กินที่"""
    em = 0.0
    for ch in text:
        o = ord(ch)
        if o == 0x0E31 or 0x0E34 <= o <= 0x0E3A or 0x0E47 <= o <= 0x0E4E:
            continue
        if ch == " ":
            em += 0.3
        elif 0x0E00 <= o <= 0x0E7F:
            em += 0.56
        elif ch.isdigit():
            em += 0.70                      # เผื่อฟอนต์สำรองที่กว้างกว่า Tahoma
        elif ch.isupper():
            em += 0.78
        elif ch.isalpha():
            em += 0.64
        else:
            em += 0.5
    return em * size


def fit_font(text, w_pt, h_pt, smax=9.0, smin=6.5):
    """เลือกขนาดตัวอักษรใหญ่สุด (<=smax) ที่ข้อความแนวนอน wrap แล้วพอดีกรอบ"""
    inner = max(w_pt - 5, 6)
    size = smax
    while size >= smin:
        lines, cur = 0, 0.0
        sp = text_width(" ", size)
        for word in text.split(" "):
            ww = text_width(word, size)
            if ww > inner:                       # คำยาวกว่ากรอบ: ตัดกลางคำ
                lines += int(ww // inner) + (1 if cur == 0 else 0)
                cur = ww % inner
            elif cur == 0:
                lines, cur = max(lines, 1), ww
            elif cur + sp + ww <= inner:
                cur += sp + ww
            else:
                lines, cur = lines + 1, ww
        if lines * size * 1.3 <= h_pt - 1:
            return size
        size -= 0.5
    return smin


def build_map(ws, items):
    """วาดผังตามพิกัดช่องกริด (1 ช่อง = CELL_PT pt): กรอบ merge + ตกแต่ง (ผนัง/วงเล็บ/ลูกศร/แถบ)"""
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "404040"

    def sr(y):            # แถว Y ของ CSV -> แถวจริงในชีต
        return y + HEAD_ROWS
    max_y = max(it["y"] + it["h"] - 1 for it in items)
    last_col = max(LEG_X + LEG_W - 1, max(it["x"] + it["w"] for it in items))
    for c in range(1, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = COL_W
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 6
    for y in range(1, max_y + 2):
        ws.row_dimensions[sr(y)].height = CELL_PT

    # หัวเรื่อง + คำแนะนำ (ข้อความล้นไปทางขวาตามเซลล์ว่าง)
    ws["A1"] = "ผังคลังสินค้า"
    ws["A1"].font = font(size=18, bold=True)
    ws["A1"].alignment = Alignment(vertical="center")
    ws["A2"] = "กดที่ชื่อกรอบเพื่อไปยังใบจดของชั้นนั้น"
    ws["A2"].font = font(size=10, italic=True, color="595959")
    ws["A2"].alignment = Alignment(vertical="center")

    # ---- เส้นขอบ: เก็บเป็น "ขอบ" ระหว่างเซลล์ ลำดับความสำคัญสูงชนะ (กรอบ < วงเล็บ < ผนัง < แถบ)
    eh, ev = {}, {}     # eh[(แถว,คอลัมน์)] = ขอบบนของเซลล์ | ev[(แถว,คอลัมน์)] = ขอบซ้ายของเซลล์

    def put(d, key, side, prio):
        if key not in d or prio >= d[key][0]:
            d[key] = (prio, side)

    def box_edges(x, y, w, h, side, prio):
        for c in range(x, x + w):
            put(eh, (sr(y), c), side, prio)
            put(eh, (sr(y + h), c), side, prio)
        for r in range(y, y + h):
            put(ev, (sr(r), x), side, prio)
            put(ev, (sr(r), x + w), side, prio)

    def style_fill(x, y, w, h, hex6):
        for r in range(y, y + h):
            for c in range(x, x + w):
                ws.cell(row=sr(r), column=c).fill = fill(hex6)

    def merge(x, y, w, h):
        if w > 1 or h > 1:
            ws.merge_cells(start_row=sr(y), start_column=x, end_row=sr(y + h - 1), end_column=x + w - 1)

    cells = {}
    P_FRAME, P_BRACE, P_WALL, P_BAR = 1, 2, 3, 4
    for it in items:
        x, y, w, h, typ = it["x"], it["y"], it["w"], it["h"], it["type"]
        color = it["color"]
        if typ in FRAME_TYPES:
            fc = color or (MAP_BLUE if it["make"] else "FFFFFF")      # ป้ายไม่ระบุสี = พื้นขาว
            y1, y2 = it["mrows"]
            merge(x, y1, w, y2 - y1 + 1)
            style_fill(x, y, w, h, fc)
            if fc != "FFFFFF":
                box_edges(x, y, w, h, Side(style="medium", color=shade(fc)), P_FRAME)
            cell = ws.cell(row=sr(y), column=x, value=it["name"])
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            tcol = "000000" if rel_lum(fc) >= 0.3 else "FFFFFF"
            low = it["name"].lower()
            if low == "office":
                size = 40.0
            elif it["name"].startswith("โซนรหัส") and h >= 6:
                size = 14.0
            else:
                size = fit_font(it["name"], w * CELL_PT, (y2 - y1 + 1) * CELL_PT)
            cell.font = font(bold=True, size=size, color=tcol)
            if it["make"]:
                link(cell, it["sheet"])
                cells[it["sheet"]] = cell.coordinate
        elif typ == "ผนัง":
            side = Side(style="thick", color=color or "FF0000")
            if w == 1:      # เส้นตั้ง ตามขอบซ้ายของคอลัมน์ X
                for r in range(y, y + h):
                    put(ev, (sr(r), x), side, P_WALL)
            else:           # เส้นนอน ตามขอบบนของแถว Y
                for c in range(x, x + w):
                    put(eh, (sr(y), c), side, P_WALL)
        elif typ == "วงเล็บ":
            side = Side(style="thin", color=color or "FF0000")
            d = it["name"]
            if h >= w:      # วงเล็บตั้ง: ขีดเส้นกลาง (spine) + ขีดสั้นหัว/ท้าย + ปลายแหลมกลาง
                half = w // 2
                tip_left = "ซ้าย" in d
                tip_cols = range(x, x + half) if tip_left else range(x + half, x + w)
                tick_cols = range(x + half, x + w) if tip_left else range(x, x + half)
                for r in range(y, y + h):
                    put(ev, (sr(r), x + half), side, P_BRACE)
                for c in tick_cols:
                    put(eh, (sr(y), c), side, P_BRACE)
                    put(eh, (sr(y + h), c), side, P_BRACE)
                for c in tip_cols:
                    put(eh, (sr(y + h // 2), c), side, P_BRACE)
            else:           # วงเล็บนอน
                half = h // 2
                tip_up = "ขึ้น" in d
                tip_rows = range(y, y + half) if tip_up else range(y + h - half, y + h)
                tick_rows = range(y + half, y + h) if tip_up else range(y, y + h - half)
                spine_y = y + half if tip_up else y + h - half
                for c in range(x, x + w):
                    put(eh, (sr(spine_y), c), side, P_BRACE)
                for r in tick_rows:
                    put(ev, (sr(r), x), side, P_BRACE)
                    put(ev, (sr(r), x + w), side, P_BRACE)
                for r in tip_rows:
                    put(ev, (sr(r), x + w // 2), side, P_BRACE)
        elif typ in ("ลูกศรขึ้น", "ลูกศรลง"):
            red = color or "FF0000"
            up = typ == "ลูกศรขึ้น"
            hd = 3                              # หัวลูกศร = 3 แถว x กว้างเต็ม
            hy = y if up else y + h - hd
            sy = y + hd if up else y
            merge(x, hy, w, hd)
            head = ws.cell(row=sr(hy), column=x, value="▲" if up else "▼")
            head.font = font(size=20, bold=True, color=red)
            head.alignment = Alignment(horizontal="center", vertical="center")
            mid = x + w // 2                    # ก้านลูกศร = คอลัมน์กลาง
            merge(mid, sy, 1, h - hd)
            style_fill(mid, sy, 1, h - hd, red)
        elif typ == "แถบ":
            bc = color or "FFC000"
            style_fill(x, y, w, h, bc)
            box_edges(x, y, w, h, Side(style="thin", color=shade(bc, 0.5)), P_BAR)

    # ---- ลงเส้นขอบจริง (หลัง merge เพื่อให้เซลล์ที่ถูก merge รับเส้นได้ครบ)
    touched = set()
    for (r, c) in list(eh):
        touched.update([(r, c), (r - 1, c)])
    for (r, c) in list(ev):
        touched.update([(r, c), (r, c - 1)])
    for (r, c) in touched:
        if r < 1 or c < 1:
            continue
        g = lambda d, k: (d[k][1] if k in d else None)
        ws.cell(row=r, column=c).border = Border(
            top=g(eh, (r, c)), bottom=g(eh, (r + 1, c)), left=g(ev, (r, c)), right=g(ev, (r, c + 1)))

    # ---- legend ขวาของผัง
    def legend(row, text, fc, bold=False, tcol="000000", box=False):
        merge(LEG_X, row, LEG_W, 3)
        c = ws.cell(row=sr(row), column=LEG_X, value=text)
        c.font = font(size=9, bold=bold, color=tcol)
        c.alignment = Alignment(horizontal="center" if fc else "left", vertical="center", wrap_text=True)
        if fc:
            style_fill(LEG_X, row, LEG_W, 3, fc)
            bs = Side(style="medium", color=shade(fc))
            for rr in range(row, row + 3):
                for cc in range(LEG_X, LEG_X + LEG_W):
                    ws.cell(row=sr(rr), column=cc).border = Border(
                        left=bs if cc == LEG_X else None, right=bs if cc == LEG_X + LEG_W - 1 else None,
                        top=bs if rr == row else None, bottom=bs if rr == row + 2 else None)
    legend(4, "คำอธิบาย", None, bold=True)
    legend(8, "กรอบเก็บของ (กดแล้วไปใบจด)", MAP_BLUE, tcol="FFFFFF")
    legend(12, "สีอื่น / ป้าย = ตามผังเดิม", None)

    # ตั้งค่าพิมพ์ A4 แนวตั้ง กว้าง 1 หน้า
    last_row = sr(max_y + 1)
    # ตัดหน้ากระดาษที่ช่องว่างระหว่างกรอบ (ไม่ผ่ากลางกรอบ): หน้าแรกจุ ~122 แถว หน้าถัดไป ~130 แถว
    spans = [f["mrows"] for f in items if f["type"] in FRAME_TYPES]
    spans += [(d["y"], d["y"] + d["h"] - 1) for d in items
              if d["type"] in ("ลูกศรขึ้น", "ลูกศรลง", "แถบ") or (d["type"] == "วงเล็บ" and d["w"] > d["h"])]
    safe = [r for r in range(1, max_y) if not any(a <= r < b for a, b in spans)]
    start, cap = 0, 122
    while max_y - start > cap:
        cand = [r for r in safe if start + cap // 2 < r <= start + cap]
        if not cand:
            break
        start = max(cand)
        ws.row_breaks.append(Break(id=sr(start)))
        cap = 130
    ws.freeze_panes = f"A{HEAD_ROWS + 1}"
    ws.print_area = f"A1:{get_column_letter(last_col)}{last_row}"
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.3, bottom=0.3)
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
    nd = sum(1 for it in items if it["decor"])
    print(f"  กรอบในผัง {len(items) - nd} (ใบจด {len(sheets)}, ป้าย {len(items) - nd - len(sheets)}) + ตกแต่ง {nd} | ชีตทั้งหมด {len(wb.sheetnames)} | แถวชีตสินค้า {n}")
    print(f"  สำรองไฟล์เดิม: {bak or '(ไม่มีไฟล์เดิม)'}")
    print(f"  ย้ายข้อมูลเดิมกลับ: {moved_sheets} ชีต / {moved_rows} แถว")
    if stray:
        print(f"  !! ข้อมูลค้าง {len(stray)} แถว (ใหม่รอบนี้ {new_stray}) อยู่ในชีต \"{STRAY_NAME}\" - ตรวจและย้ายเอง")
    else:
        print("  ข้อมูลค้าง: 0 แถว")


if __name__ == "__main__":
    main()
