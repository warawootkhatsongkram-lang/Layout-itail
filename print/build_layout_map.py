# -*- coding: utf-8 -*-
"""สร้างแผนที่ผังคลัง A3 แนวตั้ง พื้นขาว (สำหรับผู้บริหาร) -> print/layout_map_A3.html
แล้วสั่ง Chrome แบบไม่เปิดหน้าต่าง พิมพ์เป็น PDF และแปลงเป็น PNG 200dpi

ข้อมูลผัง: อ่านจากข้อมูลที่ฝังใน finder.html (ชุดเดียวกับหน้าเว็บ มาจากชีต "ผังให้เดย์ปรับเองได้"
ใน LAY_OUT_02.xlsx) — อ่านอย่างเดียว ไม่แตะไฟล์ Excel
ผังเปลี่ยน: build finder.html ใหม่ก่อน (tools/build_finder.py) แล้วรันไฟล์นี้

วิธีใช้:  python print/build_layout_map.py
"""
import html
import json
import math
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from build_qr_poster import Iso  # ใช้ตัววาดภาพ isometric ชุดเดียวกับป้าย QR

SRC = os.path.join(ROOT, "finder.html")
OUT_HTML = os.path.join(HERE, "layout_map.html")
OUT_PDF = os.path.join(HERE, "layout_map_A3.pdf")
OUT_PNG = os.path.join(HERE, "layout_map_A3.png")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

TH_MONTH = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]


def load_layout():
    with open(SRC, encoding="utf-8") as f:
        h = f.read()
    m = re.search(r'<script[^>]*id="app-data"[^>]*>(.*?)</script>', h, re.S)
    if not m:
        sys.exit("ไม่พบข้อมูลผังใน finder.html")
    return json.loads(m.group(1))


def th_date(iso):
    # เวลาไฟล์เป็น UTC -> เวลาไทย -> วันที่ พ.ศ.
    if not iso:
        return "-"
    t = datetime.strptime(iso, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc) + timedelta(hours=7)
    return f"{t.day} {TH_MONTH[t.month - 1]} {t.year + 543}"


# ---------- กลุ่มสี (ตามเฉดของสีใน Excel แบบเดียวกับหน้าเว็บ) -> สีสำหรับพื้นขาว ----------
PAL = {
    "blue":   ("#E7EFFA", "#2A4D80", "#0E2A4F"),
    "yellow": ("#FFE680", "#A87C00", "#3F2E00"),
    "red":    ("#FCE3E0", "#B42318", "#7A1A12"),
    "green":  ("#DBF3E5", "#1E8449", "#0F4D2A"),
    "purple": ("#ECE5F5", "#5B3F8C", "#3A2560"),
    "white":  ("#FFFFFF", "#94A3B8", "#1E293B"),
    "dark":   ("#334155", "#1E293B", "#FFFFFF"),
}


def group(hexv):
    hexv = (hexv or "FFFFFF").upper()
    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    mx, mn = max(r, g, b), min(r, g, b)
    if mx - mn < 30:
        return "white" if mx > 200 else "dark"
    if mx == r and g > 150 and b < 100:
        return "yellow"
    if mx == r and g < 140:
        return "red"
    if mx == g and r < 120:
        return "green"
    if mx == r and b > 120:
        return "purple"
    if b >= r and b >= g and r > 100 and g < 110:
        return "purple"
    if mx == b or (b > r and g < 120):
        return "blue" if (mx + mn) / 2 > 90 else "dark"
    return "blue"


# ---------- เส้นทางรูปทรงแบบ Office (เหมือนหน้าเว็บ) ----------
def clamp(v, lo, hi):
    return min(hi, max(lo, v))


def down_arrow(w, h, a1, a2):
    ss = min(w, h)
    a1 = clamp(50000 if a1 is None else a1, 0, 100000)
    a2 = clamp(50000 if a2 is None else a2, 0, 100000 * h / ss if ss else 0)
    y1, dx, hc = h - ss * a2 / 100000, w * a1 / 200000, w / 2
    return (f"M{hc - dx},0 L{hc + dx},0 L{hc + dx},{y1} L{w},{y1} L{hc},{h} "
            f"L0,{y1} L{hc - dx},{y1} Z")


def brace(w, h, a1, a2):
    a2 = clamp(50000 if a2 is None else a2, 0, 100000)
    q3, ss = min(100000 - a2, a2) / 2, min(w, h)
    a1 = clamp(8333 if a1 is None else a1, 0, q3 * h / ss if ss else 0)
    y1, y3 = ss * a1 / 100000, h * a2 / 100000
    y2, y4, r = y3 - y1, h - y1, w / 2
    A = f" A{r},{y1} 0 0 "
    return (f"M0,0{A}1 {r},{y1} L{r},{y2}{A}0 {w},{y3}{A}0 {r},{y3 + y1} "
            f"L{r},{y4}{A}1 0,{h}")


# ---------- วาดผัง (หน่วย = ช่องกริดของ Excel, 1 หน่วย = S มม.) ----------
def map_svg(d, S):
    frames, decor = d["frames"], d["decor"]
    xs, ys = [], []
    for f in frames:
        xs += [f["x"] - 1, f["x"] - 1 + f["w"]]
        ys += [f["y"] - 1, f["y"] - 1 + f["h"]]
    for e in decor:
        if e["type"] == "เส้น":
            xs += [e["x1"], e["x2"]]
            ys += [e["y1"], e["y2"]]
        else:
            xs += [e["x"] - 1, e["x"] - 1 + e["w"]]
            ys += [e["y"] - 1, e["y"] - 1 + e["h"]]
    pad = 1.2
    x0, y0 = min(xs) - pad, min(ys) - pad
    W, H = max(xs) - x0 + pad, max(ys) - y0 + pad
    mm = 1 / S                       # 1 มม. บนกระดาษ ในหน่วยผัง
    out = []

    # พื้นที่ office วาดก่อน (เป็นพื้นหลังของกรอบเล็กที่อยู่ข้างใน)
    order = sorted(frames, key=lambda f: 0 if f["w"] * f["h"] > 600 else 1)
    labels = []
    for f in order:
        L, T, w, h = f["x"] - 1, f["y"] - 1, f["w"], f["h"]
        big = w * h > 600
        name = html.escape(f["name"])
        if f.get("nf"):                         # กล่องข้อความเฉยๆ ไม่มีพื้น
            out.append(f'<rect class="nf" x="{L:.3f}" y="{T:.3f}" width="{w:.3f}" height="{h:.3f}" rx="{0.6 * mm:.3f}" '
                       f'stroke-width="{0.25 * mm:.3f}" stroke-dasharray="{0.8 * mm:.3f} {0.6 * mm:.3f}"/>')
            fill, edge, ink = PAL["white"]
        else:
            fill, edge, ink = PAL[group(f.get("color"))]
            if big:
                out.append(f'<rect x="{L:.3f}" y="{T:.3f}" width="{w:.3f}" height="{h:.3f}" fill="url(#hatch)" '
                           f'stroke="{edge}" stroke-width="{0.5 * mm:.3f}"/>')
            else:
                sw = 0 if f.get("nb") else 0.35 * mm
                out.append(f'<rect x="{L:.3f}" y="{T:.3f}" width="{w:.3f}" height="{h:.3f}" rx="{0.5 * mm:.3f}" '
                           f'fill="{fill}" stroke="{edge}" stroke-width="{sw:.3f}"/>')
                if not f.get("nb") and min(w, h) >= 2.5:   # แถบสีด้านบนกรอบ ให้ดูเป็นป้ายชั้นวาง
                    out.append(f'<rect x="{L + 0.35 * mm:.3f}" y="{T + 0.35 * mm:.3f}" width="{w - 0.7 * mm:.3f}" '
                               f'height="{0.6 * mm:.3f}" fill="{edge}" opacity=".85"/>')
        tr = f.get("tr") or 0
        cx, cy = L + w / 2, T + h / 2
        if big:                               # office: ข้อความชิดบน ไม่ทับกรอบเล็กข้างใน
            cy = T + 6.5
        along, across = (h, w) if tr in (90, 270) else (w, h)
        grp = f"{round(L * 2) / 2}|{round(w)}|{tr}|{'s' if h < 3 else ''}"
        labels.append(f'<text class="lbl{" big" if big else ""}" x="{cx:.3f}" y="{cy:.3f}" fill="{ink}" '
                      f'data-along="{along - 1.6 * mm:.3f}" data-across="{across:.3f}" data-grp="{grp}" '
                      f'transform="rotate({tr} {cx:.3f} {cy:.3f})">{name}</text>')

    # ตกแต่ง: ผนัง / วงเล็บแบ่งโซน / ลูกศร / วงกลม / แถบประตู
    for e in decor:
        if e["type"] == "เส้น":
            wall = e.get("lw", 0) >= 2.5 or e.get("thick")
            sw = (1.5 if wall else 0.4) * mm
            col = "#1E293B" if wall else "#D14343"
            out.append(f'<line x1="{e["x1"]:.3f}" y1="{e["y1"]:.3f}" x2="{e["x2"]:.3f}" y2="{e["y2"]:.3f}" '
                       f'stroke="{col}" stroke-width="{sw:.3f}" stroke-linecap="square"/>')
            continue
        g, gw, gh = e["geom"], e["gw"], e["gh"]
        tf = (f'translate({e["cx"]:.3f} {e["cy"]:.3f}) rotate({e["rot"]}) scale({-1 if e.get("fh") else 1} 1) '
              f'translate({-gw / 2:.3f} {-gh / 2:.3f})')
        if g == "rightBrace":
            shp = f'<path d="{brace(gw, gh, e.get("adj1"), e.get("adj2"))}" fill="none" stroke="#D14343" stroke-width="{0.45 * mm:.3f}"/>'
        elif g == "downArrow":
            shp = f'<path d="{down_arrow(gw, gh, e.get("adj1"), e.get("adj2"))}" fill="#E5484D" stroke="#B42318" stroke-width="{0.2 * mm:.3f}"/>'
        elif g == "ellipse":
            fill, edge, _ = PAL[group(e.get("fill"))]
            shp = (f'<ellipse cx="{gw / 2:.3f}" cy="{gh / 2:.3f}" rx="{gw / 2:.3f}" ry="{gh / 2:.3f}" fill="#BFD1EA" fill-opacity=".6" '
                   f'stroke="{edge}" stroke-width="{0.3 * mm:.3f}"/>')
        else:
            grp_ = group(e.get("fill"))
            col = "#FFC400" if grp_ == "yellow" else PAL[grp_][1]
            shp = (f'<rect width="{gw:.3f}" height="{gh:.3f}" fill="{col}" stroke="#1E293B" '
                   f'stroke-width="{0.25 * mm:.3f}"/>')
        out.append(f'<g transform="{tf}">{shp}</g>')

    svg = (f'<svg class="map" viewBox="{x0:.3f} {y0:.3f} {W:.3f} {H:.3f}" width="{W * S:.2f}mm" height="{H * S:.2f}mm" '
           f'xmlns="http://www.w3.org/2000/svg">'
           f'<defs><pattern id="hatch" width="{2.2 * mm:.3f}" height="{2.2 * mm:.3f}" patternUnits="userSpaceOnUse" '
           f'patternTransform="rotate(45)"><rect width="{2.2 * mm:.3f}" height="{2.2 * mm:.3f}" fill="#F1F5F9"/>'
           f'<line x1="0" y1="0" x2="0" y2="{2.2 * mm:.3f}" stroke="#CBD5E1" stroke-width="{0.35 * mm:.3f}"/></pattern></defs>'
           + "".join(out) + "".join(labels) + "</svg>")
    return svg, W * S, H * S


# ---------- ภาพคลัง isometric (โทนสว่าง) ----------
def hero_svg():
    floor = Iso(ox=96, oy=60, s=0.95)
    floor.floor_grid(-50, -40, 120, 110, 12)
    a = Iso(ox=96, oy=60, s=0.95)
    a.rack(6, -12, 4, seed=7)
    a.rack(6, -40, 4, seed=13)
    b = Iso(ox=96, oy=60, s=0.95)
    b.rack(6, -12, 4, seed=3)
    b.rack(6, -40, 4, seed=17)
    m = Iso(ox=96, oy=60, s=0.95)
    m.forklift(44, 40)
    flip = 'transform="translate(192 0) scale(-1 1)"'
    body = floor.out + a.out + [f"<g {flip}>"] + b.out + ["</g>"] + m.out
    return ('<svg class="hero" viewBox="-30 -30 252 140" xmlns="http://www.w3.org/2000/svg">'
            '<defs><linearGradient id="fd" x1="0" y1="0" x2="0" y2="1">'
            '<stop offset="0" stop-color="#fff" stop-opacity="1"/><stop offset=".72" stop-color="#fff" stop-opacity="1"/>'
            '<stop offset="1" stop-color="#fff" stop-opacity="0"/></linearGradient>'
            '<mask id="mf"><rect x="-30" y="-30" width="252" height="140" fill="url(#fd)"/></mask></defs>'
            f'<g mask="url(#mf)">{"".join(body)}</g></svg>')


# ---------- ดัชนีชื่อกรอบ แยกตามแถว (ซ้าย/กลาง/ขวา ตามตำแหน่งบนผัง) ----------
def columns(d):
    cols = {"ซ้าย": [], "กลาง": [], "ขวา": []}
    for f in d["frames"]:
        if f.get("nf") or f["w"] * f["h"] > 600 or (f.get("nb") and "โซน" in f["name"]):
            continue
        cx = f["x"] - 1 + f["w"] / 2
        key = "ซ้าย" if cx < 22 else ("กลาง" if cx < 46 else "ขวา")
        cols[key].append(f)
    for k in cols:
        cols[k].sort(key=lambda f: f["y"])
    return cols


ICON_RACK = '<svg viewBox="0 0 24 24"><path d="M4 3v18M20 3v18M4 8h16M4 14h16M4 20h16"/><path d="M7 5h4v3H7zM13 10h4v4h-4zM7 16h5v4H7z"/></svg>'


def build_html():
    d = load_layout()
    S = 1.06                                  # มม. ต่อ 1 หน่วยผัง (ผังสูงเต็มหน้า A3)
    msvg, mw, mh = map_svg(d, S)
    cols = columns(d)
    date = th_date(d.get("fileDate"))

    def chips(lst):
        return "".join(f'<span class="chip c-{group(f.get("color"))}">{html.escape(f["name"])}</span>' for f in lst)

    zones = "".join(
        f'<div class="zone"><div class="zh"><b>แถว{k}</b></div><div class="chips">{chips(v)}</div></div>'
        for k, v in cols.items())

    return f"""<!doctype html>
<html lang="th"><head><meta charset="utf-8">
<title>ผังคลังสินค้า — LAY OUT STORE</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Kanit:wght@500;600;700&family=Sarabun:wght@400;600;700&display=block" rel="stylesheet">
<style>
@page {{ size: A3; margin: 0 }}
* {{ box-sizing: border-box; margin: 0; padding: 0 }}
html, body {{ width: 297mm; height: 420mm; background: #fff;
  -webkit-print-color-adjust: exact; print-color-adjust: exact }}
body {{ font-family: Sarabun, sans-serif; color: #0F172A }}
.page {{ position: relative; width: 297mm; height: 420mm; overflow: hidden; background: #fff }}
.hazard {{ position: absolute; left: 0; right: 0; height: 5mm;
  background: repeating-linear-gradient(-45deg, #FFC400 0 5mm, #111827 5mm 10mm) }}
.hazard.top {{ top: 0 }} .hazard.bot {{ bottom: 0 }}
.grid {{ position: absolute; inset: 0;
  background-image: linear-gradient(#EEF2F7 .25mm, transparent .25mm), linear-gradient(90deg, #EEF2F7 .25mm, transparent .25mm);
  background-size: 10mm 10mm; background-position: -0.1mm -0.1mm }}
.grid::after {{ content: ""; position: absolute; inset: 0;
  background: radial-gradient(120% 70% at 75% 30%, rgba(255,255,255,0) 0, rgba(255,255,255,.75) 60%, #fff 100%) }}

/* ---- หัวกระดาษ ---- */
.head {{ position: absolute; left: 112mm; right: 16mm; top: 22mm }}
.eyebrow {{ font: 600 10.5pt Kanit; letter-spacing: .32em; color: #A87C00; display: flex; align-items: center; gap: 3mm }}
.eyebrow i {{ display: block; width: 12mm; height: .6mm; background: #FFC400 }}
h1 {{ font: 700 54pt/1 Kanit; letter-spacing: -.01em; margin-top: 3mm; color: #0F172A }}
h1 b {{ color: #0F172A; background: linear-gradient(transparent 62%, #FFC400 62%, #FFC400 92%, transparent 92%); padding: 0 1mm }}
.sub {{ font: 600 19pt Kanit; color: #334155; margin-top: 3mm }}
.meta {{ display: flex; gap: 4mm; margin-top: 4.5mm }}
.pill {{ font: 600 10.5pt Sarabun; color: #1E293B; border: .35mm solid #CBD5E1; border-radius: 99mm; padding: 1.2mm 4mm; background: #fff }}
.pill b {{ color: #A87C00 }}

.hero {{ position: absolute; right: 6mm; top: 86mm; width: 186mm }}
.hero .grid {{ stroke: #D9E1EC; stroke-width: .25 }}
.hero .up {{ stroke: #64748B; stroke-width: .7 }} .hero .up.f {{ stroke: #334155; stroke-width: .9 }}
.hero .beam {{ stroke: #F59E0B; stroke-width: .8 }} .hero .beam.f {{ stroke: #D97706; stroke-width: 1 }}
.hero .bL {{ fill: #C9D6E8 }} .hero .bR {{ fill: #9FB3CF }} .hero .bT {{ fill: #E6EDF6 }}
.hero .yL {{ fill: #FFD43B }} .hero .yR {{ fill: #F2B705 }} .hero .yT {{ fill: #FFE680 }}
.hero .kL {{ fill: #334155 }} .hero .kR {{ fill: #1E293B }} .hero .kT {{ fill: #475569 }}
.hero polygon {{ stroke: rgba(15,23,42,.18); stroke-width: .18 }}
.hero .cage {{ stroke: #334155; stroke-width: .8 }} .hero .wheel {{ fill: #111827 }}
.hero line.grid {{ stroke: #DCE3ED; stroke-width: .25 }}

/* ---- ผัง ---- */
.mapcard {{ position: absolute; left: 14mm; top: 12mm; width: {mw + 14:.1f}mm; height: {mh + 26:.1f}mm;
  background: #fff; border: .4mm solid #CBD5E1; border-radius: 3mm;
  box-shadow: 0 1.5mm 5mm rgba(15,23,42,.10) }}
.mapcard .tag {{ position: absolute; left: 7mm; top: 4.5mm; font: 600 10pt Kanit; color: #334155; letter-spacing: .12em }}
.mapcard .tag b {{ display: inline-block; width: 2.4mm; height: 2.4mm; background: #FFC400; margin-right: 2mm; vertical-align: .1mm }}
.mapcard svg.map {{ position: absolute; left: 7mm; top: 13mm }}
.lbl {{ font-family: Kanit; font-weight: 600; text-anchor: middle; dominant-baseline: central }}
.lbl.big {{ font-weight: 700; letter-spacing: .2em }}
.nf {{ fill: none; stroke: #94A3B8 }}

/* ---- ด้านขวา ---- */
.side {{ position: absolute; left: 112mm; right: 16mm; bottom: 34mm }}
.sec {{ font: 600 13pt Kanit; color: #0F172A; display: flex; align-items: center; gap: 2.5mm; margin-bottom: 3mm }}
.sec .ic {{ width: 7mm; height: 7mm; border-radius: 1.6mm; background: #0F172A; display: grid; place-items: center }}
.sec .ic svg {{ width: 4.6mm; height: 4.6mm; fill: none; stroke: #FFC400; stroke-width: 2; stroke-linecap: round; stroke-linejoin: round }}
.zones {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 3.5mm }}
.zone {{ border: .35mm solid #E2E8F0; border-radius: 2.5mm; padding: 3.5mm; background: #FAFBFD }}
.zh {{ font: 600 12pt Kanit; color: #0F172A; border-bottom: .35mm solid #E2E8F0; padding-bottom: 1.6mm; margin-bottom: 2.2mm }}
.zh b {{ border-left: 1.2mm solid #FFC400; padding-left: 2mm }}
.chips {{ display: flex; flex-wrap: wrap; gap: 1.6mm }}
.chip {{ font: 600 10pt Kanit; padding: .7mm 2mm; border-radius: 1mm; border: .3mm solid }}
.c-blue {{ background: #E7EFFA; border-color: #A9BEDC; color: #0E2A4F }}
.c-yellow {{ background: #FFE680; border-color: #D4AA2A; color: #3F2E00 }}
.c-red {{ background: #FCE3E0; border-color: #E8A39B; color: #7A1A12 }}
.c-green {{ background: #DBF3E5; border-color: #93CFAE; color: #0F4D2A }}
.c-purple {{ background: #ECE5F5; border-color: #B9A6D8; color: #3A2560 }}

.legend {{ display: grid; grid-template-columns: 1fr 1fr; gap: 3mm 6mm; margin-top: 8mm }}
.lg {{ display: flex; align-items: center; gap: 3mm; font: 400 11pt Sarabun; color: #1E293B }}
.sw {{ width: 11mm; height: 6mm; border-radius: 1mm; flex: none }}
.sw.rack {{ background: #E7EFFA; border: .35mm solid #2A4D80; box-shadow: inset 0 .9mm 0 #2A4D80 }}
.sw.belt {{ background: #FFE680; border: .35mm solid #A87C00; box-shadow: inset 0 .9mm 0 #A87C00 }}
.sw.office {{ background: repeating-linear-gradient(45deg, #F1F5F9 0 1mm, #CBD5E1 1mm 1.3mm); border: .4mm solid #2A4D80 }}
.sw.wall {{ height: 1.5mm; background: #1E293B; border-radius: 0 }}
.sw.brace {{ height: 6mm; border: .45mm solid #D14343; border-left: 0; border-radius: 0 3mm 3mm 0; width: 4mm; margin: 0 3.5mm }}
.sw.zonet {{ background: #DBF3E5; border: 0; display: grid; place-items: center; font: 600 7pt Kanit; color: #0F4D2A }}
.sw.note {{ border: .3mm dashed #94A3B8; background: #fff }}

.tblock {{ position: absolute; right: 16mm; bottom: 13mm; left: 112mm; display: grid;
  grid-template-columns: 1.4fr 1fr 1fr; border: .4mm solid #0F172A; border-radius: 1.5mm; overflow: hidden }}
.tblock div {{ padding: 2.2mm 3.5mm; border-left: .3mm solid #CBD5E1 }}
.tblock div:first-child {{ border-left: 0; background: #0F172A; color: #fff }}
.tblock small {{ display: block; font: 400 8pt Sarabun; color: #64748B }}
.tblock div:first-child small {{ color: #FFC400 }}
.tblock span {{ font: 600 11pt Kanit }}
</style></head>
<body><div class="page">
  <div class="grid"></div>
  <div class="hazard top"></div>

  <div class="head">
    <div class="eyebrow"><i></i>WAREHOUSE LAYOUT<i></i></div>
    <h1>LAY OUT <b>STORE</b></h1>
    <div class="sub">ผังคลังสินค้า</div>
    <div class="meta"><span class="pill">ข้อมูลผัง ณ <b>{date}</b></span><span class="pill">A3 · มาตราส่วนตามผังจริง</span></div>
  </div>
  {hero_svg()}

  <div class="mapcard">
    <div class="tag"><b></b>FLOOR PLAN</div>
    {msvg}
  </div>

  <div class="side">
    <div class="sec"><span class="ic">{ICON_RACK}</span>ตำแหน่งเก็บสินค้าแต่ละแถว</div>
    <div class="zones">{zones}</div>

    <div class="legend">
      <div class="lg"><span class="sw rack"></span>ชั้นวาง / ตำแหน่งเก็บสินค้า</div>
      <div class="lg"><span class="sw belt"></span>สายพาน</div>
      <div class="lg"><span class="sw office"></span>office</div>
      <div class="lg"><span class="sw wall"></span>ผนังอาคาร</div>
      <div class="lg"><span class="sw brace"></span>ขอบเขตโซนรหัส</div>
      <div class="lg"><span class="sw zonet">โซน</span>ป้ายโซนรหัส</div>
    </div>
  </div>

  <div class="tblock">
    <div><small>เอกสาร</small><span>ผังคลังสินค้า · LAY OUT STORE</span></div>
    <div><small>ข้อมูลผัง ณ</small><span>{date}</span></div>
    <div><small>ที่มา</small><span>LAY_OUT_02.xlsx</span></div>
  </div>
  <div class="hazard bot"></div>
</div>
<script>
// ปรับขนาดตัวอักษรให้พอดีกรอบ แล้วใช้ขนาดเล็กสุดของกลุ่มเดียวกัน (กรอบแถวเดียวกัน ตัวหนังสือเท่ากัน)
(function () {{
  var best = {{}}, els = document.querySelectorAll('text.lbl');
  els.forEach(function (t) {{
    var along = +t.dataset.along, across = +t.dataset.across;
    var fs = Math.min(across * (across < 2.5 ? 0.72 : 0.58), t.classList.contains('big') ? 5 : 3.4);
    t.setAttribute('font-size', fs);
    var len = t.getComputedTextLength();
    if (len > along) fs = fs * along / len;
    t._fs = fs;
    var g = t.dataset.grp;
    best[g] = Math.min(best[g] || 99, fs);
  }});
  els.forEach(function (t) {{ t.setAttribute('font-size', best[t.dataset.grp].toFixed(3)); }});
  document.body.setAttribute('data-ready', '1');
}})();
</script>
</body></html>
"""


def render():
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(build_html())
    url = "file:///" + OUT_HTML.replace("\\", "/")
    subprocess.run(
        [CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
         "--virtual-time-budget=10000", f"--print-to-pdf={OUT_PDF}", url],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    import pymupdf
    doc = pymupdf.open(OUT_PDF)
    doc[0].get_pixmap(dpi=200).save(OUT_PNG)
    print("pages:", len(doc), "size(pt):", doc[0].rect)
    print("wrote", OUT_HTML, OUT_PDF, OUT_PNG, sep="\n  ")


if __name__ == "__main__":
    render()
