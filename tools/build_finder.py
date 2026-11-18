#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_finder.py — สร้างเว็บ "หาของในคลัง" (finder.html) ไฟล์เดียว จาก LAY_OUT_02.xlsx

วิธีใช้ (จากโฟลเดอร์โปรเจค):
    python3 tools/build_finder.py
    (ทดสอบกับไฟล์อื่น: python3 tools/build_finder.py --src path/to/file.xlsx --out path/to/out.html)

โครงใหม่ (เดย์ตัดสินใจ): ผังและตำแหน่งสินค้า "ตามไฟล์ LAY_OUT_02.xlsx" ไฟล์เดียว
    - ชีต "ผังคลัง"  = ผัง (วาดด้วยเซลล์: merge + สีพื้น + เส้นขอบ) -> กรอบ/ตกแต่งบนแผนที่เว็บ
    - ชีต "รวม"      = ตำแหน่งสินค้า (B=Location ต้องตรงชื่อกรอบบนผัง, C=รหัสสินค้า, E=หมายเหตุ)
    - ชีต "สินค้า"   = ชื่อ/หน่วยของรหัส (ชื่อรายการในชีตรวมเป็นสูตร จึงดึงชื่อจากชีตนี้เอง)
    - ชีตอื่น (ค้นหา, กรอกป้าย, ป้ายแบบ1, ป้ายแบบ2) ไม่ใช้

สคริปต์นี้ทำอะไร:
    อ่าน 3 ชีตข้างบน -> ฝัง JSON ลงใน tools/finder_template.html (แทนที่ /*__DATA__*/) -> finder.html
    ข้อมูลที่ฝังนี้เป็น "ตัวสำรองสุดท้าย" ตอนเปิดหน้าเว็บแบบออฟไลน์/เปิดจากไฟล์ (file://)
    ปกติเว็บจะโหลด LAY_OUT_02.xlsx สดๆ เอง (หรือจาก Supabase ถ้าตั้งค่าไว้) ไม่ต้องรันสคริปต์นี้ทุกครั้งที่แก้ข้อมูล

กติกาอ่านผัง (ต้องเหมือนฝั่ง JavaScript ใน finder_template.html ทุกข้อ — แก้ที่หนึ่งต้องแก้ทั้งสองที่):
    - พิกัดเป็น "หน่วยช่อง": กว้าง 1 หน่วย = 8px (คอลัมน์) / สูง 1 หน่วย = 6pt (แถว) ตามขนาดจริงในชีต
      x,y = มุมซ้ายบนนับจาก 1; แถว 1-3 เป็นหัวเรื่อง ไม่นับ (แถว 4 = y 1)
    - กรอบ = ช่อง merge ที่มีข้อความ (ยกเว้น ▲ ▼ และกล่องคำอธิบายทางขวาของผัง) สี = สีพื้นเซลล์ (ไม่มี = ขาว)
      กรอบที่สีพื้นยาวลงไปต่ำกว่าช่อง merge (เช่น office ที่ครอบกรอบเล็ก) ถือว่ากรอบสูงเท่าสีพื้น
    - เส้นขอบชนิด thick/thin = เส้นตกแต่ง (ผนัง/วงเล็บ) ; ลูกศร = ช่อง merge ▲/▼ + ก้านสีแดง
      เซลล์สีพื้นที่ไม่ใช่กรอบ (ไม่ merge) = แถบสี
    - กรอบเก็บของทุกกรอบแตะดูรายการได้ (ใบจดใหม่ไม่มีคอลัมน์บอกประเภท)

สคริปต์นี้ไม่แก้ไฟล์ต้นทางใดๆ (อ่านอย่างเดียว) เขียนเฉพาะ finder.html
"""
import argparse
import colorsys
import html
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SRC = ROOT / "LAY_OUT_02.xlsx"
TEMPLATE_PATH = Path(__file__).resolve().parent / "finder_template.html"
DEFAULT_OUT = ROOT / "finder.html"
PLACEHOLDER = "/*__DATA__*/"

SHEET_MAP = "ผังคลัง"
SHEET_ALL = "รวม"
SHEET_PROD = "สินค้า"
LEGEND_TITLE = "คำอธิบาย"     # กล่องคำอธิบายทางขวาของผัง (ไม่ใช่กรอบ) — ข้ามทุกอย่างที่อยู่ตั้งแต่คอลัมน์นี้ไปทางขวา
HEAD_ROWS = 3                  # แถวหัวเรื่องเหนือกริดผัง (ไม่นับเป็นพิกัด)
UNIT_W_PX = 8                  # 1 หน่วยกว้าง = 8px
UNIT_H_PT = 6                  # 1 หน่วยสูง = 6pt
DEF_COL_PX = 64                # คอลัมน์ที่ไม่ระบุความกว้าง (Excel default 8.43 ตัวอักษร)
DEF_ROW_PT = 15                # แถวที่ไม่ระบุความสูง
MDW = 7                        # ความกว้างเลขหนึ่งตัวของฟอนต์ default (ใช้แปลงความกว้างคอลัมน์เป็น px เหมือน SheetJS)
LINE_STYLES = ("thin", "thick")  # ชนิดเส้นขอบที่นับเป็นเส้นตกแต่ง (เส้น medium = ขอบกรอบ ไม่นับ)
UNKNOWN_CODE = "ไม่พบรหัส"


def norm_key(text):
    """ชื่อกรอบแบบเทียบกัน: ยุบช่องว่างซ้อน + ตัวพิมพ์เล็ก"""
    return " ".join(str(text).split()).lower()


def r3(v):
    """ปัดทศนิยม 3 ตำแหน่ง (กันเลขลอยต่างกันระหว่าง Python/JS) และตัด .0 ออก"""
    v = round(float(v) + 0.0, 3)
    return int(v) if v == int(v) else v


# ---------------------------------------------------------------------------
# สีและสไตล์เซลล์: อ่านจาก XML ดิบในไฟล์ xlsx โดยตรง
# (ไม่ใช้ openpyxl อ่านเส้นขอบ เพราะ openpyxl ตอนโหลดจะเขียนเส้นขอบรอบช่อง merge ใหม่จากเซลล์มุมซ้ายบน -> ผลเพี้ยน
#  และฝั่ง JavaScript ก็อ่าน XML ดิบเช่นกัน จึงได้ผลตรงกัน)
# ---------------------------------------------------------------------------
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]


def read_theme_colors(zf):
    """ชุดสี theme ของไฟล์ -> list ตามเลข theme ของ Excel (0=lt1,1=dk1,2=lt2,3=dk2,4..9=accent1..6)"""
    try:
        xml = zf.read("xl/theme/theme1.xml").decode("utf-8", "ignore")
    except KeyError:
        return []
    m = re.search(r"<a:clrScheme[^>]*>(.*?)</a:clrScheme>", xml, re.S)
    if not m:
        return []
    out = {}
    for name, body in re.findall(r"<a:(dk1|lt1|dk2|lt2|accent[1-6])>(.*?)</a:\1>", m.group(1), re.S):
        v = re.search(r'(?:lastClr|val)="([0-9A-Fa-f]{6})"', body)
        if v:
            out[name] = v.group(1).upper()
    return [out.get(n) for n in THEME_ORDER]


def apply_tint(hex6, tint):
    r, g, b = (int(hex6[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return "%02X%02X%02X" % (int(r * 255 + 0.5), int(g * 255 + 0.5), int(b * 255 + 0.5))


def color_hex(attrs, theme, auto_black=False):
    """แอตทริบิวต์ของแท็ก <color>/<fgColor> -> 'RRGGBB' หรือ None (ไม่รู้จัก)"""
    rgb = attrs.get("rgb")
    if rgb and re.fullmatch(r"[0-9A-Fa-f]{8}|[0-9A-Fa-f]{6}", rgb):
        return rgb[-6:].upper()
    if attrs.get("theme") is not None and attrs.get("theme").isdigit():
        i = int(attrs["theme"])
        if i < len(theme) and theme[i]:
            tint = float(attrs.get("tint") or 0)
            return apply_tint(theme[i], tint) if tint else theme[i]
        return None
    if auto_black and attrs.get("auto") in ("1", "true"):
        return "000000"
    return None


def read_styles(zf):
    """styles.xml -> (xf_fill, xf_border): xf_fill[i] = สีพื้น 'RRGGBB'|None ; xf_border[i] = {'top'|'bottom'|'left'|'right': (style, 'RRGGBB')}"""
    theme = read_theme_colors(zf)
    root = ET.fromstring(zf.read("xl/styles.xml"))
    def kids(tag):
        el = root.find(NS + tag)
        return list(el) if el is not None else []

    fills = []
    for fill in kids("fills"):
        pf = fill.find(NS + "patternFill")
        fg = pf.find(NS + "fgColor") if pf is not None else None
        ok = pf is not None and pf.get("patternType") == "solid" and fg is not None
        fills.append(color_hex(fg.attrib, theme) if ok else None)
    borders = []
    for b in kids("borders"):
        d = {}
        for side in ("top", "bottom", "left", "right"):
            el = b.find(NS + side)
            if el is not None and el.get("style"):
                col = el.find(NS + "color")
                hexv = color_hex(col.attrib, theme, True) if col is not None else None
                d[side] = (el.get("style"), hexv or "000000")
        borders.append(d)
    xf_fill, xf_border = [], []
    for xf in kids("cellXfs"):
        fi, bi = int(xf.get("fillId") or 0), int(xf.get("borderId") or 0)
        xf_fill.append(fills[fi] if fi < len(fills) else None)
        xf_border.append(borders[bi] if bi < len(borders) else {})
    return xf_fill, xf_border


def sheet_xml_path(zf, name):
    """หาไฟล์ XML ของชีตตามชื่อ (workbook.xml + rels)"""
    wbx = zf.read("xl/workbook.xml").decode("utf-8")
    rid = None
    for tag in re.findall(r"<sheet\s[^>]*>", wbx):
        nm = re.search(r'\sname="([^"]*)"', tag)
        if nm and html.unescape(nm.group(1)) == name:
            m = re.search(r'\sr:id="([^"]*)"', tag) or re.search(r'\s[\w]*:?id="([^"]*)"', tag)
            rid = m.group(1) if m else None
    if rid is None:
        raise SystemExit("ไม่พบชีต %s ในไฟล์" % name)
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    for tag in re.findall(r"<Relationship\s[^>]*>", rels):
        if re.search(r'\sId="%s"' % re.escape(rid), tag):
            t = re.search(r'\sTarget="([^"]*)"', tag).group(1)
            return t[1:] if t.startswith("/") else "xl/" + t
    raise SystemExit("หาไฟล์ของชีต %s ไม่เจอ" % name)


def read_cell_styles(zf, name, xf_fill, xf_border):
    """ชีต -> {(แถว, คอลัมน์): (สีพื้น|None, ขอบ{})} เฉพาะเซลล์ที่มีอยู่ใน XML"""
    xml = zf.read(sheet_xml_path(zf, name)).decode("utf-8")
    out = {}
    for attrs in re.findall(r"<c\s([^>]*?)/?>", xml):
        r = re.search(r'\br="([A-Z]+)(\d+)"', attrs)
        s = re.search(r'\bs="(\d+)"', attrs)
        if not r:
            continue
        i = int(s.group(1)) if s else 0
        col = 0
        for ch in r.group(1):
            col = col * 26 + (ord(ch) - 64)
        out[(int(r.group(2)), col)] = (xf_fill[i] if i < len(xf_fill) else None,
                                       xf_border[i] if i < len(xf_border) else {})
    return out


# ---------------------------------------------------------------------------
# ขนาดแถว/คอลัมน์ -> พิกัดสะสม (หน่วยช่อง)
# ---------------------------------------------------------------------------
def width_to_px(w):
    return int(((256 * w + int(128 / MDW)) / 256.0) * MDW)


def grid_sizes(ws, max_r, max_c):
    """คืน (left, top): left[c] = ขอบซ้ายของคอลัมน์ c, top[r] = ขอบบนของแถว r (r>=HEAD_ROWS+1; ขอบบนแถว 4 = 0)"""
    sf = ws.sheet_format
    dflt_px = width_to_px(sf.defaultColWidth) if sf.defaultColWidth else DEF_COL_PX
    dflt_pt = sf.defaultRowHeight or DEF_ROW_PT
    colpx = {}
    for dim in ws.column_dimensions.values():
        if dim.width is None:
            continue
        lo, hi = dim.min or 1, dim.max or dim.min or 1
        for c in range(lo, min(hi, max_c + 2) + 1):
            colpx[c] = 0 if dim.hidden else width_to_px(dim.width)
    left = {1: 0.0}
    for c in range(1, max_c + 2):
        left[c + 1] = left[c] + colpx.get(c, dflt_px) / float(UNIT_W_PX)
    top = {HEAD_ROWS + 1: 0.0}
    for r in range(HEAD_ROWS + 1, max_r + 2):
        rd = ws.row_dimensions.get(r)
        h = rd.height if rd is not None and rd.height is not None else dflt_pt
        if rd is not None and rd.hidden:
            h = 0
        top[r + 1] = top[r] + h / float(UNIT_H_PT)
    return left, top


# ---------------------------------------------------------------------------
# ผังคลัง -> frames + decor
# ---------------------------------------------------------------------------
def read_map(wb, zf):
    """อ่านชีต ผังคลัง -> (frames, decor)  กติกาเหมือน parseMap() ใน finder_template.html"""
    ws = wb[SHEET_MAP]
    xf_fill, xf_border = read_styles(zf)
    sty = read_cell_styles(zf, SHEET_MAP, xf_fill, xf_border)      # (แถว,คอลัมน์) -> (สีพื้น, ขอบ)
    max_r = max([ws.max_row] + [k[0] for k in sty])
    max_c = max([ws.max_column] + [k[1] for k in sty])
    left, top = grid_sizes(ws, max_r, max_c)

    def fill_of(r, c):
        return sty[(r, c)][0] if (r, c) in sty else None

    def text_of(r, c):
        v = ws.cell(r, c).value
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return " ".join(str(v).split())

    # ช่วง merge (r1,c1,r2,c2) + ช่องเดี่ยวที่มีข้อความ (กรอบ 1 ช่อง ไม่ merge)
    ranges, merge_id = [], {}
    for m in ws.merged_cells.ranges:
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                merge_id[(r, c)] = len(ranges)
        ranges.append((m.min_row, m.min_col, m.max_row, m.max_col))
    for row in ws.iter_rows():
        for cell in row:
            if (cell.row, cell.column) not in merge_id and cell.row > HEAD_ROWS and text_of(cell.row, cell.column):
                ranges.append((cell.row, cell.column, cell.row, cell.column))
    ranges.sort(key=lambda x: (x[0], x[1]))

    legend_col = None
    for (r1, c1, r2, c2) in ranges:
        if text_of(r1, c1) == LEGEND_TITLE:
            legend_col = c1
            break

    frames, heads, blocks, covered = [], [], [], set()
    for (r1, c1, r2, c2) in ranges:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                covered.add((r, c))
        if r2 <= HEAD_ROWS:
            continue                                  # แถวหัวเรื่อง
        if legend_col is not None and c1 >= legend_col:
            continue                                  # กล่องคำอธิบาย
        txt = text_of(r1, c1)
        fill = fill_of(r1, c1)
        if txt in ("▲", "▼"):
            heads.append((r1, c1, r2, c2, txt))
        elif txt:
            # ขยายลงล่างถ้าสีพื้นยังต่อเนื่อง (ดูเฉพาะคอลัมน์ซ้าย-ขวาสุดของกรอบ และต้องไม่ใช่ช่อง merge อื่น)
            rext = r2
            while fill and all((rext + 1, cc) not in merge_id and fill_of(rext + 1, cc) == fill for cc in (c1, c2)):
                rext += 1
            for r in range(r2 + 1, rext + 1):
                for c in range(c1, c2 + 1):
                    covered.add((r, c))
            frames.append({
                "id": "f%d" % (len(frames) + 1),
                "name": txt,
                "store": True,
                "x": r3(left[c1] + 1), "y": r3(top[max(r1, HEAD_ROWS + 1)] + 1),
                "w": r3(left[c2 + 1] - left[c1]),
                "h": r3(top[rext + 1] - top[r1]),
                "lh": r3(top[r2 + 1] - top[r1]),
                "color": fill or "FFFFFF",
            })
        elif fill and fill != "FFFFFF":
            blocks.append((r1, c1, r2, c2, fill))     # ช่อง merge สีพื้นไม่มีข้อความ (เช่น ก้านลูกศร)

    decor = []

    # ---- แถบสี: เซลล์สีพื้นที่ไม่ merge และไม่อยู่ในกรอบ (รวมเป็นสี่เหลี่ยมใหญ่สุด)
    fills = {}
    for (r, c), (f, _b) in sty.items():
        if r <= HEAD_ROWS or (r, c) in covered or (legend_col is not None and c >= legend_col):
            continue
        if f and f != "FFFFFF":
            fills[(r, c)] = f
    bars, active = [], {}
    for r in sorted({k[0] for k in fills}):
        runs, cs = [], sorted(c for (rr, c) in fills if rr == r)
        i = 0
        while i < len(cs):
            j = i
            while j + 1 < len(cs) and cs[j + 1] == cs[j] + 1 and fills[(r, cs[j + 1])] == fills[(r, cs[i])]:
                j += 1
            runs.append((cs[i], cs[j], fills[(r, cs[i])]))
            i = j + 1
        nxt = {}
        for k in runs:
            if k in active and active[k][1] == r - 1:
                active[k][1] = r
                nxt[k] = active[k]
            else:
                nxt[k] = [r, r]
                bars.append((k, nxt[k]))
        active = nxt
    for ((a, b, col), (ra, rb)) in bars:
        decor.append(_rect_decor("แถบ", ra, a, rb, b, col, left, top))

    # ---- ลูกศร: ▲/▼ (ช่อง merge) + ก้านสีพื้น (ช่อง merge 1 คอลัมน์ ต่อติดหัว)
    used = set()
    for (r1, c1, r2, c2, txt) in heads:
        up = txt == "▲"
        mid = c1 + (c2 - c1 + 1) // 2
        stem = None
        for k, (sr1, sc1, sr2, sc2, col) in enumerate(blocks):
            if k in used or sc1 != mid or sc2 != mid:
                continue
            if (up and sr1 == r2 + 1) or (not up and sr2 == r1 - 1):
                stem = (k, sr1, sr2, col)
                break
        if stem:
            used.add(stem[0])
            ya, yb = (r1, stem[2]) if up else (stem[1], r2)
            col = stem[3]
        else:
            ya, yb, col = r1, r2, "FF0000"
        d = _rect_decor("ลูกศรขึ้น" if up else "ลูกศรลง", ya, c1, yb, c2, col, left, top)
        d["hd"] = r3(top[r2 + 1] - top[r1])           # ความสูงหัวลูกศร
        d["sx"] = r3(left[mid]) if stem else 0        # ก้าน: ซ้ายสุด (หน่วย 0-based) และความกว้าง
        d["sw"] = r3(left[mid + 1] - left[mid]) if stem else 0
        decor.append(d)
    for k, (r1, c1, r2, c2, col) in enumerate(blocks):
        if k not in used:
            decor.append(_rect_decor("แถบ", r1, c1, r2, c2, col, left, top))

    # ---- เส้นตกแต่ง: ขอบเซลล์ชนิด thin/thick -> รวมเป็นเส้นยาว
    # ข้ามขอบที่อยู่ "ข้างใน" ช่อง merge เดียวกัน (Excel ไม่แสดงเส้นนั้น)
    strength = {"thin": 1, "thick": 2}
    hedge, vedge = {}, {}     # hedge[(แถว r, คอลัมน์ c)] = ขอบบนของแถว r | vedge[(แถว r, คอลัมน์ c)] = ขอบซ้ายของคอลัมน์ c

    def put(d, key, side):
        if side is None or side[0] not in LINE_STYLES:
            return
        old = d.get(key)
        if old is None or strength[side[0]] > strength[old[0]]:
            d[key] = side

    def same_merge(a, b):
        return a in merge_id and merge_id.get(a) == merge_id.get(b)

    for (r, c), (_f, b) in sty.items():
        if not b:
            continue
        if not same_merge((r, c), (r - 1, c)):
            put(hedge, (r, c), b.get("top"))
        if not same_merge((r, c), (r + 1, c)):
            put(hedge, (r + 1, c), b.get("bottom"))
        if not same_merge((r, c), (r, c - 1)):
            put(vedge, (r, c), b.get("left"))
        if not same_merge((r, c), (r, c + 1)):
            put(vedge, (r, c + 1), b.get("right"))
    hedge = {k: v for k, v in hedge.items() if k[0] >= HEAD_ROWS + 1}
    vedge = {k: v for k, v in vedge.items() if k[0] > HEAD_ROWS}
    lines = []
    for (r, c) in sorted(hedge):                                  # เส้นนอน: ต่อคอลัมน์ติดกัน
        v = hedge[(r, c)]
        if lines and lines[-1]["_h"] and lines[-1]["_k"] == (r, v) and lines[-1]["_end"] == c - 1:
            lines[-1]["_end"] = c
        else:
            lines.append({"_h": True, "_k": (r, v), "_start": c, "_end": c})
    for (c, r) in sorted((k[1], k[0]) for k in vedge):            # เส้นตั้ง: ต่อแถวติดกัน
        v = vedge[(r, c)]
        if lines and not lines[-1]["_h"] and lines[-1]["_k"] == (c, v) and lines[-1]["_end"] == r - 1:
            lines[-1]["_end"] = r
        else:
            lines.append({"_h": False, "_k": (c, v), "_start": r, "_end": r})
    for ln in lines:
        (pos, (style, col)) = ln["_k"]
        if ln["_h"]:
            x1, x2, y1, y2 = left[ln["_start"]], left[ln["_end"] + 1], top[pos], top[pos]
        else:
            x1, x2, y1, y2 = left[pos], left[pos], top[ln["_start"]], top[ln["_end"] + 1]
        decor.append({"type": "เส้น", "x1": r3(x1), "y1": r3(y1), "x2": r3(x2), "y2": r3(y2),
                      "color": col, "thick": style == "thick",
                      "x": r3(x1 + 1), "y": r3(y1 + 1), "w": r3(x2 - x1), "h": r3(y2 - y1)})
    return frames, decor


def _rect_decor(typ, r1, c1, r2, c2, color, left, top):
    """สี่เหลี่ยมตกแต่ง (แถบ/ลูกศร) จากช่วงแถว/คอลัมน์ -> พิกัดหน่วยช่อง"""
    r1 = max(r1, HEAD_ROWS + 1)
    return {"type": typ, "x": r3(left[c1] + 1), "y": r3(top[r1] + 1),
            "w": r3(left[c2 + 1] - left[c1]), "h": r3(top[r2 + 1] - top[r1]), "color": color}


# ---------------------------------------------------------------------------
# สินค้า + รวม
# ---------------------------------------------------------------------------
def cell_str(v):
    """ค่าเซลล์ -> ข้อความ (ตัวเลขเต็ม ไม่ใช่ 5.6e+11)"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        if v != v or v in (float("inf"), float("-inf")):
            return ""
        if v.is_integer():
            return str(int(v))
    return str(v).strip()


def read_products(wb):
    """ชีต สินค้า -> [[code, name, unit, total], ...] (ข้ามหัวตารางและ Grand Total)"""
    ws = wb[SHEET_PROD]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = cell_str(row[0] if len(row) > 0 else None)
        if not code or code == "Grand Total":
            continue
        total = row[3] if len(row) > 3 and isinstance(row[3], (int, float)) and not isinstance(row[3], bool) else 0
        if isinstance(total, float) and total.is_integer():
            total = int(total)
        out.append([code,
                    cell_str(row[1] if len(row) > 1 else None),
                    cell_str(row[2] if len(row) > 2 else None),
                    total])
    return out


def read_locations(wb, frames, products):
    """ชีต รวม -> (locations, skipped)
    หัวตารางหาจากแถวที่ B = 'Location' (ไม่เจอใช้แถว 4) ข้อมูลเริ่มแถวถัดไป ข้ามแถวที่ B,C ว่างทั้งคู่ (รวมโน้ตท้ายตาราง)"""
    ws = wb[SHEET_ALL]
    head = 4
    for r in range(1, 11):
        if norm_key(cell_str(ws.cell(r, 2).value)) == "location":
            head = r
            break
    locations, skipped = [], []
    for r in range(head + 1, ws.max_row + 1):
        loc = " ".join(cell_str(ws.cell(r, 2).value).split())
        code = cell_str(ws.cell(r, 3).value)
        if not loc and not code:
            continue
        if not code:
            skipped.append({"row": r, "why": "มี Location แต่ไม่มีรหัสสินค้า"})
            continue
        if not loc:
            skipped.append({"row": r, "why": "มีรหัสสินค้าแต่ไม่มี Location"})
            continue
        locations.append({"frame": norm_key(loc), "loc": loc, "code": code,
                          "note": cell_str(ws.cell(r, 5).value), "row": r})
    return locations, skipped


def file_date(wb):
    m = wb.properties.modified
    return m.replace(microsecond=0).isoformat() + "Z" if m else None


def parse_workbook(path):
    """อ่านไฟล์ -> dict {frames, decor, products, locations, skipped, fileDate}; ขาดชีตจำเป็น = SystemExit"""
    wb = openpyxl.load_workbook(path)
    zf = zipfile.ZipFile(path)
    miss = [n for n in (SHEET_MAP, SHEET_ALL, SHEET_PROD) if n not in wb.sheetnames]
    if miss:
        raise SystemExit("ไฟล์ %s ขาดชีต: %s" % (path, ", ".join(miss)))
    frames, decor = read_map(wb, zf)
    if not frames:
        raise SystemExit("อ่านชีต ผังคลัง ไม่เจอกรอบเลย (กรอบต้องเป็นช่อง merge ที่มีชื่อกรอบอยู่ข้างใน)")
    products = read_products(wb)
    locations, skipped = read_locations(wb, frames, products)
    return {"fileDate": file_date(wb), "frames": frames, "decor": decor,
            "products": products, "locations": locations, "skipped": skipped}


def main():
    ap = argparse.ArgumentParser(description="สร้าง finder.html จาก LAY_OUT_02.xlsx")
    ap.add_argument("--src", default=str(DEFAULT_SRC), help="ไฟล์ xlsx ต้นทาง (ค่าเริ่มต้น LAY_OUT_02.xlsx)")
    ap.add_argument("--out", default=str(DEFAULT_OUT), help="ไฟล์ html ที่จะเขียน (ค่าเริ่มต้น finder.html)")
    ap.add_argument("--json", help="(ทดสอบ) เขียนผลที่อ่านได้เป็น JSON ไปที่ไฟล์นี้ แล้วหยุด ไม่สร้าง html")
    args = ap.parse_args()

    src, out = Path(args.src), Path(args.out)
    data = parse_workbook(src)
    data["source"] = src.name
    data["generatedAt"] = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    if args.json:
        Path(args.json).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    else:
        # escape "</" กันปิด <script> ก่อนเวลา
        payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
        template = TEMPLATE_PATH.read_text(encoding="utf-8")
        if template.count(PLACEHOLDER) != 1:
            raise SystemExit("template ต้องมี %s เพียง 1 จุด" % PLACEHOLDER)
        out.write_text(template.replace(PLACEHOLDER, payload), encoding="utf-8")

    keys = {norm_key(f["name"]) for f in data["frames"]}
    unmatched = sorted({l["loc"] for l in data["locations"] if l["frame"] not in keys})
    kinds = {}
    for d in data["decor"]:
        kinds[d["type"]] = kinds.get(d["type"], 0) + 1
    print("อ่านจาก %s" % src)
    print("กรอบบนผัง %d | ตกแต่ง %d %s" % (len(data["frames"]), len(data["decor"]), kinds))
    print("สินค้า %d รายการ | ตำแหน่งสินค้า (ชีต รวม) %d แถว | ข้ามแถว %d" % (
        len(data["products"]), len(data["locations"]), len(data["skipped"])))
    if unmatched:
        print("คำเตือน: Location ที่ไม่ตรงชื่อกรอบบนผัง %d ชื่อ: %s" % (len(unmatched), ", ".join(unmatched[:10])))
    if not args.json:
        print("เขียน %s (%.0f KB)" % (out, out.stat().st_size / 1024))


if __name__ == "__main__":
    sys.exit(main())
