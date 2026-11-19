#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_location_list.py — อัปเดต dropdown "Location" (ชีต รวม คอลัมน์ B) ให้ตรงกับชื่อกรอบบนผังที่เดย์วาดด้วย shape

วิธีใช้:
    python3 tools/update_location_list.py INPUT.xlsx OUTPUT.xlsx
    (ไม่แก้ไฟล์ INPUT เด็ดขาด — เขียนไฟล์ใหม่ที่ OUTPUT ; ถ้า INPUT ไม่มีชีต "ผังให้เดย์ปรับเองได้" + shape จะหยุดโดยไม่เขียนอะไร)

ทำอะไร (แก้ XML ดิบใน zip เท่านั้น — ไม่ใช้ openpyxl บันทึก เพราะ openpyxl ทำให้ shape หาย):
    1) อ่านชื่อกรอบจากชีต shape (ใช้ตัวอ่านเดียวกับ finder: build_finder.read_shape_map)
       -> ชื่อซ้ำเหลือชื่อเดียว (เทียบด้วย norm_key) เรียงชื่อแบบธรรมชาติ P1,P2,P10 (--sort natural, ค่าเริ่มต้น)
          (หรือ --sort map = ตามตำแหน่งบนผัง บนลงล่าง ซ้ายไปขวา) ; ตัดป้ายที่ไม่ใช่ที่เก็บของ (NOT_STORAGE) ออก
    2) เขียนรายการลงชีต ค้นหา คอลัมน์ L (ซ่อนอยู่) แถว 1..N เป็น inline string ; ลบเซลล์ L เก่าที่เกิน N
    3) แก้ทุก data validation ที่ชี้ ค้นหา!$L$1:$L$82 (ทุกชีต + definedNames) ให้เป็น $L$1:$L$N
    ชิ้นส่วนอื่นใน zip ไม่ถูกแตะ (ไบต์เท่าเดิม) รวมทั้ง drawing ของผัง, sharedStrings, calcChain
"""
import argparse
import hashlib
import html
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_finder as bf  # noqa: E402  (ใช้ตัวอ่านผัง shape ตัวเดียวกับเว็บ)

SHEET_FIND = "ค้นหา"       # ชีตที่เก็บรายการ (คอลัมน์ L ซ่อนอยู่)
# ป้ายบนผังที่ไม่ใช่ที่เก็บของ -> ไม่ใส่ใน dropdown (ตามข้อตัดสินใจเดิมใน spec.md) — บนเว็บยังเป็นกรอบแตะได้ตามปกติ
NOT_STORAGE = ("โซนรหัส 5", "โซนรหัส 6", "โซนรหัส 7", "office", "โต๊ะนั้ง")
LIST_COL = "L"
LIST_COL_NUM = 12
CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
CELL = r'<c\s[^>]*?(?:/>|>.*?</c>)'
RANGE_RE = re.compile(r"((?:'ค้นหา'|ค้นหา)!\$L\$1:\$L\$)(\d+)")


def location_names(frames, mode):
    """ชื่อกรอบไม่ซ้ำ -> list เรียงตาม mode"""
    seen, rows = set(), []
    for f in frames:
        k = bf.norm_key(f["name"])
        if k in seen:
            continue
        seen.add(k)
        rows.append(f)
    if mode == "natural":
        def nat(f):
            return [(0, int(t), "") if t.isdigit() else (1, 0, t.lower()) for t in re.split(r"(\d+)", f["name"]) if t != ""]
        rows.sort(key=nat)
    else:
        rows.sort(key=lambda f: (round(f["y"]), f["x"]))      # บนลงล่าง แล้วซ้ายไปขวา
    return [CTRL.sub("", f["name"]) for f in rows]


def col_of(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n


def rewrite_list(xml, names):
    """แก้ XML ของชีต ค้นหา: ใส่ L1..LN ใหม่ ลบ L เก่าที่เกิน N (แก้เฉพาะเซลล์ L ที่เกี่ยวข้อง ที่เหลือไบต์เท่าเดิม)"""
    n = len(names)
    seen_rows = set()

    def fix_row(m):
        head, r, body = m.group(1), int(m.group(2)), m.group(3)
        if r > max(n, 1) and not re.search(r'<c r="%s%d"[ />]' % (LIST_COL, r), body):
            return m.group(0)
        seen_rows.add(r)
        cells = list(re.finditer(CELL, body, re.S))
        out, placed = [], False
        want = None
        if 1 <= r <= n:
            want = '<c r="%s%d" t="inlineStr"><is><t%s>%s</t></is></c>' % (
                LIST_COL, r, ' xml:space="preserve"' if names[r - 1] != names[r - 1].strip() else "", escape(names[r - 1]))
        pos = 0
        for c in cells:
            ref = re.match(r'<c\s+r="([A-Z]+\d+)"', c.group(0)).group(1)
            col = col_of(ref)
            out.append(body[pos:c.start()])
            pos = c.end()
            if col == LIST_COL_NUM:                       # เซลล์ L เดิม: แทนที่ (หรือลบถ้าเกิน N)
                if want and not placed:
                    out.append(want)
                    placed = True
                continue
            if col > LIST_COL_NUM and want and not placed:
                out.append(want)
                placed = True
            out.append(c.group(0))
        out.append(body[pos:])
        if want and not placed:
            out.append(want)
        return head + "".join(out) + "</row>"

    new = re.sub(r'(<row r="(\d+)"[^>]*?(?<!/)>)(.*?)</row>', fix_row, xml, flags=re.S)
    missing = [r for r in range(1, n + 1) if r not in seen_rows]
    if missing:
        raise SystemExit("ชีต %s ไม่มีแถว %s ... จึงใส่รายการไม่ครบ (หยุด ไม่ได้เขียนอะไร)" % (SHEET_FIND, missing[:5]))
    return new


def main():
    ap = argparse.ArgumentParser(description="อัปเดตรายการ Location (ชีต ค้นหา คอลัมน์ L + dropdown) จากชื่อกรอบในชีต shape")
    ap.add_argument("src", help="ไฟล์ xlsx ต้นทาง (ไม่ถูกแก้)")
    ap.add_argument("dst", help="ไฟล์ xlsx ที่จะเขียนใหม่")
    ap.add_argument("--sort", choices=("map", "natural"), default="natural",
                    help="natural = เรียงชื่อ A,B,..,P1,P2,P10 หาง่ายใน dropdown (ค่าเริ่มต้น) / map = ตามตำแหน่งบนผัง")
    args = ap.parse_args()
    src, dst = os.path.abspath(args.src), os.path.abspath(args.dst)
    if src == dst or (os.path.exists(dst) and os.path.samefile(src, dst)):
        raise SystemExit("ปลายทางต้องไม่ใช่ไฟล์ต้นทาง (ไม่แก้ไฟล์เดิมในที่เดิม)")

    with zipfile.ZipFile(src) as zin:
        shape = bf.read_shape_map(zin)
        if shape is None:
            raise SystemExit("ไฟล์ %s ไม่มีชีต \"%s\" ที่มี shape — หยุด ไม่ได้เขียนอะไร" % (src, bf.SHEET_SHAPES))
        skip = {bf.norm_key(x) for x in NOT_STORAGE}
        names = [x for x in location_names(shape[0], args.sort) if bf.norm_key(x) not in skip]
        n = len(names)
        find_path = bf.find_sheet_xml(zin, SHEET_FIND)
        if find_path is None:
            raise SystemExit("ไม่พบชีต \"%s\" ในไฟล์ — หยุด ไม่ได้เขียนอะไร" % SHEET_FIND)

        edits, notes = {}, []
        edits[find_path] = rewrite_list(zin.read(find_path).decode("utf-8"), names).encode("utf-8")
        # แก้ range ของ dropdown ทุกที่ (ทุกชีต + definedNames ใน workbook.xml)
        for info in zin.infolist():
            if not re.fullmatch(r"xl/(worksheets/[^/]+|workbook)\.xml", info.filename):
                continue
            text = (edits.get(info.filename) or zin.read(info.filename)).decode("utf-8")
            hits = RANGE_RE.findall(text)
            if hits:
                text = RANGE_RE.sub(lambda m: m.group(1) + str(n), text)
                notes.append("%s: %d จุด (%s -> %d)" % (info.filename, len(hits), ",".join(sorted({h[1] for h in hits})), n))
                edits[info.filename] = text.encode("utf-8")
        if not notes:
            print("คำเตือน: ไม่พบ data validation ที่ชี้ %s!$L$1:$L$.. ในไฟล์ (ใส่รายการแล้ว แต่ไม่มี dropdown ให้อัปเดต)" % SHEET_FIND)

        tmp = dst + ".tmp"
        with zipfile.ZipFile(tmp, "w") as zout:
            for info in zin.infolist():
                data = edits.get(info.filename)
                if data is None:
                    data = zin.read(info.filename)
                zi = zipfile.ZipInfo(info.filename, info.date_time)
                zi.compress_type, zi.external_attr, zi.create_system = info.compress_type, info.external_attr, info.create_system
                zout.writestr(zi, data)

    # ---- ตรวจผลที่เขียน ----
    changed = []
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(tmp) as b:
        if a.namelist() != b.namelist():
            os.remove(tmp)
            raise SystemExit("รายชื่อชิ้นส่วนใน zip ไม่ตรงกับต้นฉบับ (ผิดปกติ) — ยกเลิก")
        for nm in a.namelist():
            if a.read(nm) != b.read(nm):
                changed.append(nm)
        dr = bf.shape_drawing_path(a)
        same_draw = a.read(dr) == b.read(dr)
        sha_a, sha_b = hashlib.sha256(a.read(dr)).hexdigest(), hashlib.sha256(b.read(dr)).hexdigest()
        out_xml = b.read(find_path).decode("utf-8")
        got = []
        for r in range(1, n + 2):
            m = re.search(r'<c r="L%d"[^>]*?(?:/>|>(.*?)</c>)' % r, out_xml, re.S)
            got.append(html.unescape(re.sub(r"<[^>]+>", "", m.group(1) or "")) if m else None)
    if got[:n] != names or got[n] is not None:
        os.remove(tmp)
        raise SystemExit("ตรวจหลังเขียนแล้วรายการใน L ไม่ตรง (ผิดปกติ) — ยกเลิก ไม่ได้เขียนไฟล์ปลายทาง")
    if not same_draw:
        os.remove(tmp)
        raise SystemExit("drawing ของผังเปลี่ยนไป (ผิดปกติ) — ยกเลิก")
    os.replace(tmp, dst)

    print("รายการ Location %d ชื่อ (เรียงแบบ %s) -> ชีต %s คอลัมน์ L1..L%d" % (n, args.sort, SHEET_FIND, n))
    print("แก้ range dropdown:", "; ".join(notes) if notes else "-")
    print("ชิ้นส่วนที่เปลี่ยนใน zip (%d ชิ้น): %s" % (len(changed), ", ".join(changed)))
    print("ชิ้นส่วนอื่นไบต์เท่าเดิมทั้งหมด")
    print("drawing %s sha256 เข้า=%s ออก=%s (%s)" % (dr, sha_a, sha_b, "เท่ากัน" if same_draw else "ต่างกัน!"))
    print("เขียน %s" % dst)


if __name__ == "__main__":
    sys.exit(main())
