#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_finder.py — สร้างเว็บ "หาของในคลัง" (finder.html) ไฟล์เดียว จากไฟล์ชุดเดียวกับที่ใช้สร้างผัง Excel

วิธีใช้ (จากโฟลเดอร์โปรเจค):
    1. แก้ layout_config.csv  (เพิ่ม/ลบ/ย้ายกรอบ ฯลฯ)
    2. python3 tools/build_layout.py   -> ได้ LAY_OUT_02.xlsx (ใบจดรายชั้น)
    3. python3 tools/build_finder.py   -> ได้ finder.html (เว็บค้นหา)
    ผังใน Excel กับผังในเว็บจะตรงกัน เพราะอ่านจาก layout_config.csv ตัวเดียวกัน

สิ่งที่สคริปต์นี้ทำ:
    - อ่าน layout_config.csv  -> frames (กรอบบนผัง: X/Y/กว้าง/สูง/สี/ประเภท) + decor (ผนัง/วงเล็บ/ลูกศร/แถบ)
    - อ่านชีต "สินค้า" ใน LAY_OUT_01.xlsx (ข้าม Grand Total) -> products
    - สร้าง "ตำแหน่งสมมติ" (mockLocations) เพราะยังไม่มีข้อมูลจริง
        * จับคู่ชื่อกรอบกับสินค้าด้วย MOCK_KEYWORDS (ค้นคำใน DESCRIPTION)
        * สินค้าที่ไม่ตรงกรอบไหนเลย -> สุ่มลงกรอบชั้นวาง/ตัวอักษร (seed คงที่ 42 ผลซ้ำได้)
    - ฝัง JSON ลงใน tools/finder_template.html (แทนที่ /*__DATA__*/) -> finder.html

ข้อมูลจริง: เมื่อลูกน้องกรอกใบจดใน LAY_OUT_02.xlsx แล้ว ให้เปิด finder.html
กด "นำเข้าไฟล์ใบจด (.xlsx)" เว็บจะอ่านชีตตามชื่อชีตของแต่ละกรอบ (ไม่ต้องรันสคริปต์นี้ใหม่)
รันสคริปต์นี้ใหม่เฉพาะตอนผัง (CSV) หรือรายการสินค้า (LAY_OUT_01) เปลี่ยน

สคริปต์นี้ไม่แก้ไฟล์ต้นทางใดๆ (อ่านอย่างเดียว) เขียนเฉพาะ finder.html
"""
import csv
import json
import random
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = ROOT / "layout_config.csv"
PRODUCT_XLSX = ROOT / "LAY_OUT_01.xlsx"
PRODUCT_SHEET = "สินค้า"
TEMPLATE_PATH = Path(__file__).resolve().parent / "finder_template.html"
OUT_PATH = ROOT / "finder.html"

PLACEHOLDER = "/*__DATA__*/"
RANDOM_SEED = 42

# ---------------------------------------------------------------------------
# คำค้นสำหรับจับคู่ "ชื่อกรอบในผัง" -> สินค้าที่ DESCRIPTION มีคำนั้น
# (ไม่สนตัวพิมพ์เล็ก/ใหญ่ ช่องว่างซ้อนถูกยุบเป็นช่องเดียว)
# กรอบที่ไม่มีสินค้าตรงในข้อมูลจริงให้ใส่ list ว่าง (ไม่จับมั่ว)
# ลำดับการจับคู่ = ลำดับกรอบใน CSV; สินค้า 1 ตัวเข้ากรอบแรกที่ match เท่านั้น
# หมายเหตุ: ห้ามใช้ 'OPP' เดี่ยวๆ เพราะไปติดคำว่า CHOPPER
# ---------------------------------------------------------------------------
MOCK_KEYWORDS = {
    "กล่องอะไฟล่": ["อะไหล่"],
    "กล่องพลาสติก": ["กล่องพลาสติก"],
    "หมวก": ["หมวก"],
    "ป้าย Incomming เดือน": ["ป้าย", "INCOMING"],
    "ถุงมือส้ม 7/1/2": ["ถุงมือส้ม", "ถุงมือยางสีส้ม"],
    "กิโลพังฟิลม์ 300mm*600mm": [],  # ไม่มีสินค้าชื่อใกล้เคียงในข้อมูลจริง
    "ฟิล์มแลป": [],                  # ไม่มีสินค้าชื่อใกล้เคียงในข้อมูลจริง
    "กระดาษชำระ": ["กระดาษชำระ", "ทิชชู่"],
    "เทป OPP": ["เทป OPP", "สก็อตเทป"],
    "ถุงHD ฟ้า * 12 นิ้ว": ["ถุงพลาสติก HD ขนาด"],
    "ฟิล์มยืเม้วนใหยญ่": ["ฟิล์มพัน"],
    "ถุงมือยางเกรด": ["ถุงมือยาง"],
    "เชือกไนล่อน": ["เชือกไนล่อน", "เชือกอวน"],
    "ชั้นวางเอี้ยม ฟ้า/น้ำเงิน": ["เอี๊ยม"],
    "ถุง 20 * 30 ชมพู": ["HD20X30"],
    "รองเท้าซีลสีฟ้า-1": ["BOOT ฟ้า"],
    "รองเท้าซีลสีฟ้า-2": ["BOOT ขาว", "BOOT ดำ", "BOOT น้ำตาล"],
    "ถุงพลาสติก PE 14*25*0.20": ["ถุงพลาสติก PE"],
    "สเปร์ย พ่นรถยนต์": ["สีสเปรย์", "สเปรย์", "แลคเกอร์"],
    "พาเลท": ["พาเลท", "PALLET"],   # ไม่มีในข้อมูลจริง (พาเลตที่มีคือฟิล์มพัน ซึ่งเข้ากรอบฟิล์มก่อนแล้ว)
    "เชื่อก/เชือกพลาสติก": ["เชือกฟาง", "สายรัด"],
    "สายพาน": ["สายพาน"],
    "ท่อ 32 มิล": ["ท่อ PVC", "ท่ออ่อน", "แคล้ม"],
    "ไม้กวาดแข็ง": ["ไม้กวาดแข็ง", "ไม้กวาดหยักไย่", "ไม้กวาดพลาสติก"],
    "ไม้กวาดอ่อน": ["ไม้กวาดอ่อน"],
    "ชั้นวางสแตนเลส": ["สแตนเลส"],
    "ชั้นวางหมึก": ["หมึก"],
}

# ประเภทที่รับสินค้า / ประเภทที่ใช้สุ่ม
# (สี default ตามประเภทถูกกำหนดใน CSS ของ finder_template.html เพื่อรองรับโหมดมืด)
STORE_TYPES = ("สินค้า", "ชั้นวาง")
RANDOM_TYPES = ("ชั้นวาง", "ตัวอักษร")


def norm(text):
    """ยุบช่องว่างซ้อน + ตัวพิมพ์เล็ก เพื่อเทียบข้อความแบบไม่สนรูปแบบ"""
    return " ".join(str(text).split()).lower()


# ประเภทตกแต่ง (ไม่ใช่กรอบ ไม่มีใบจด ไม่ต้องกดได้) - ต้องตรงกับ build_layout.py
DECOR_TYPES = ("ผนัง", "วงเล็บ", "ลูกศรขึ้น", "ลูกศรลง", "แถบ")


def read_frames():
    """อ่าน layout_config.csv (UTF-8 BOM) -> (frames, decor)
    พิกัด X,Y = มุมซ้ายบน นับจาก 1, กว้าง/สูง = จำนวนช่อง (หน่วยเดียวกับ build_layout.py)
    frames = กรอบ/ป้าย (store=True คือกดได้) | decor = ผนัง/วงเล็บ/ลูกศร/แถบ"""
    frames, decor = [], []
    seen_ids = set()
    with open(CSV_PATH, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            name = (row["ชื่อในผัง"] or "").strip()
            if not name:
                continue  # ข้ามบรรทัดว่าง (เหมือน build_layout.py)
            seq = int(row["ลำดับ"])
            typ = row["ประเภท"].strip()
            x, y, w, h = (int(float(row[k])) for k in ("X", "Y", "กว้าง", "สูง"))
            color = row["สีพื้น"].strip().lstrip("#").upper()
            if typ in DECOR_TYPES:
                decor.append({"type": typ, "name": name, "x": x, "y": y, "w": w, "h": h, "color": color})
                continue
            store = row["ทำชีต"].strip().upper() == "Y"
            sheet = row["ชื่อชีต"].strip()
            if store:
                if not sheet:
                    raise SystemExit(f"กรอบลำดับ {seq} ทำชีต=Y แต่ไม่มีชื่อชีต")
                frame_id = sheet
            else:
                frame_id = f"sign-{seq}"
            if frame_id in seen_ids:
                raise SystemExit(f"id กรอบซ้ำ: {frame_id}")
            seen_ids.add(frame_id)
            frames.append({
                "id": frame_id,
                "name": name,
                "type": typ,
                "store": store,
                "x": x, "y": y, "w": w, "h": h,
                "lh": h,            # จำนวนแถวส่วนบนที่ใช้วางข้อความ (กรอบใหญ่ที่ครอบกรอบเล็ก เช่น office)
                "color": color,
            })
    # กรอบใหญ่ที่มีกรอบเล็กอยู่ข้างใน: ข้อความอยู่เฉพาะส่วนบนที่ไม่มีกรอบเล็ก (ตามกติกาเดียวกับ build_layout.py)
    for big in frames:
        for sm in frames:
            if sm is big:
                continue
            inside = (sm["x"] >= big["x"] and sm["y"] > big["y"] and
                      sm["x"] + sm["w"] <= big["x"] + big["w"] and sm["y"] + sm["h"] <= big["y"] + big["h"])
            if inside:
                big["lh"] = min(big["lh"], sm["y"] - big["y"])
    return frames, decor


def read_products():
    """อ่านชีตสินค้า -> [[code, name, unit, total], ...] (ข้าม header และ Grand Total)"""
    wb = openpyxl.load_workbook(PRODUCT_XLSX, read_only=True, data_only=True)
    ws = wb[PRODUCT_SHEET]
    products = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[0]
        if code is None or str(code).strip() == "" or str(code).strip() == "Grand Total":
            continue
        name = row[1] if row[1] is not None else ""
        unit = row[2] if row[2] is not None else ""
        total = row[3] if isinstance(row[3], (int, float)) else 0
        if isinstance(total, float) and total.is_integer():
            total = int(total)
        products.append([str(code).strip(), str(name).strip(), str(unit).strip(), total])
    wb.close()
    return products


def make_mock_locations(frames, products):
    """จับคู่สินค้า-กรอบด้วยคำค้น แล้วสุ่มส่วนที่เหลือ -> (locations, สถิติ)"""
    rng = random.Random(RANDOM_SEED)
    by_name = {f["name"]: f for f in frames if f["store"]}

    # ตรวจ MOCK_KEYWORDS ว่าชื่อกรอบมีจริงและเป็นประเภทที่รับสินค้า
    for name in MOCK_KEYWORDS:
        f = by_name.get(name)
        if f is None:
            raise SystemExit(f"MOCK_KEYWORDS: ไม่พบกรอบชื่อ '{name}' ใน CSV (ชื่อเปลี่ยน/สะกดผิด?)")
        if f["type"] not in STORE_TYPES:
            raise SystemExit(f"MOCK_KEYWORDS: กรอบ '{name}' ประเภท {f['type']} ไม่ใช่สินค้า/ชั้นวาง")

    # ลำดับการตรวจ = ลำดับกรอบใน CSV
    matchers = []
    for f in frames:
        if f["store"] and f["name"] in MOCK_KEYWORDS:
            kws = [norm(k) for k in MOCK_KEYWORDS[f["name"]]]
            matchers.append((f, kws))

    # กรอบที่ใช้สุ่ม: ชั้นวาง/ตัวอักษรที่ทำชีต และไม่ได้มีธีมสินค้าเฉพาะ (ไม่อยู่ใน MOCK_KEYWORDS)
    pool = [f for f in frames
            if f["store"] and f["type"] in RANDOM_TYPES and f["name"] not in MOCK_KEYWORDS]
    if not pool:
        raise SystemExit("ไม่มีกรอบชั้นวาง/ตัวอักษรสำหรับสุ่มสินค้า")

    matched_count = {f["id"]: 0 for f, _ in matchers}
    locations = []
    random_count = 0
    for code, name, _unit, total in products:
        text = norm(f"{code} {name}")
        target = None
        for f, kws in matchers:
            if any(k and k in text for k in kws):
                target = f
                break
        if target is not None:
            matched_count[target["id"]] += 1
        else:
            target = rng.choice(pool)
            random_count += 1
        level = str(rng.randint(1, 4))
        locations.append({"frame": target["id"], "level": level, "code": code,
                          "qty": total, "note": ""})
    return locations, matched_count, random_count


def main():
    frames, decor = read_frames()
    products = read_products()
    locations, matched_count, random_count = make_mock_locations(frames, products)

    data = {
        "generatedAt": datetime.now().replace(microsecond=0).isoformat(),
        "frames": frames,
        "decor": decor,
        "products": products,
        "mockLocations": locations,
    }
    # escape "</" กันปิด <script> ก่อนเวลา
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    if template.count(PLACEHOLDER) != 1:
        raise SystemExit(f"template ต้องมี {PLACEHOLDER} เพียง 1 จุด")
    OUT_PATH.write_text(template.replace(PLACEHOLDER, payload), encoding="utf-8")

    store_frames = [f for f in frames if f["store"]]
    print(f"กรอบทั้งหมด {len(frames)} (กรอบเก็บของ {len(store_frames)}, ป้าย {len(frames) - len(store_frames)}) + ตกแต่ง {len(decor)}")
    print(f"สินค้า {len(products)} รายการ")
    print("จับคู่ต่อกรอบ:")
    for f in frames:
        if f["id"] in matched_count:
            print(f"  {matched_count[f['id']]:4d}  {f['name']}")
    print(f"จับคู่รวม {sum(matched_count.values())} รายการ / สุ่ม {random_count} รายการ")
    print(f"เขียน {OUT_PATH} ({OUT_PATH.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    sys.exit(main())
