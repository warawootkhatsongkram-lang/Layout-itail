# -*- coding: utf-8 -*-
"""สร้างป้าย QR ขนาด A4 (ธีม "คลังกลางคืน") -> print/qr_poster.html
แล้วสั่ง Chrome แบบไม่เปิดหน้าต่าง พิมพ์เป็น PDF และแปลงเป็น PNG 300dpi

วิธีใช้:  python print/build_qr_poster.py
เปลี่ยนลิงก์: แก้ค่า URL ด้านล่างแล้วรันใหม่
"""
import math
import os
import subprocess
import sys

import qrcode

URL = "https://layout-itail.pages.dev/"
URL_TEXT = "layout-itail.pages.dev"

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_HTML = os.path.join(HERE, "qr_poster.html")
OUT_PDF = os.path.join(HERE, "qr_poster_A4.pdf")
OUT_PNG = os.path.join(HERE, "qr_poster_A4.png")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"


# ---------- QR (แก้ไขความผิดพลาดระดับ H = ทนรอยเปื้อน/ยับได้มากสุด) ----------
def qr_svg():
    q = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, border=0)
    q.add_data(URL)
    q.make(fit=True)
    m = q.get_matrix()
    n = len(m)
    d = "".join(
        f"M{x},{y}h1v1h-1z" for y, row in enumerate(m) for x, v in enumerate(row) if v
    )
    qz = 4  # ขอบขาวรอบ QR 4 ช่องตามมาตรฐาน
    return (
        f'<svg class="qr" viewBox="{-qz} {-qz} {n + 2 * qz} {n + 2 * qz}" '
        f'shape-rendering="crispEdges" xmlns="http://www.w3.org/2000/svg">'
        f'<rect x="{-qz}" y="{-qz}" width="{n + 2 * qz}" height="{n + 2 * qz}" fill="#fff"/>'
        f'<path d="{d}" fill="#0b1017"/></svg>'
    ), n


# ---------- ภาพคลังแบบ isometric (หน่วย = มม. บนกระดาษ A4) ----------
COS = math.cos(math.radians(30))


class Iso:
    def __init__(self, ox, oy, s):
        self.ox, self.oy, self.s = ox, oy, s
        self.out = []

    def p(self, x, y, z):
        return (self.ox + (x - y) * COS * self.s, self.oy + (x + y) * 0.5 * self.s - z * self.s)

    def poly(self, pts, cls):
        pp = " ".join(f"{a:.2f},{b:.2f}" for a, b in (self.p(*q) for q in pts))
        self.out.append(f'<polygon class="{cls}" points="{pp}"/>')

    def line(self, a, b, cls):
        (x1, y1), (x2, y2) = self.p(*a), self.p(*b)
        self.out.append(f'<line class="{cls}" x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}"/>')

    def box(self, x, y, z, w, d, h, tone="b"):
        # วาด 3 หน้าที่หันมาทางคนดู: บน / หน้าซ้าย (y+d) / หน้าขวา (x+w)
        self.poly([(x, y + d, z), (x + w, y + d, z), (x + w, y + d, z + h), (x, y + d, z + h)], f"{tone}L")
        self.poly([(x + w, y, z), (x + w, y + d, z), (x + w, y + d, z + h), (x + w, y, z + h)], f"{tone}R")
        self.poly([(x, y, z + h), (x + w, y, z + h), (x + w, y + d, z + h), (x, y + d, z + h)], f"{tone}T")

    def rack(self, x0, y0, bays, bay=24, depth=10, levels=(0, 13, 26, 39), height=46, seed=1):
        L = bays * bay
        # เสาและคานด้านหลัง
        for i in range(bays + 1):
            self.line((x0 + i * bay, y0, 0), (x0 + i * bay, y0, height), "up")
        for z in levels[1:]:
            self.line((x0, y0, z), (x0 + L, y0, z), "beam")
        # สินค้าบนชั้น (สุ่มแบบคงที่ เพื่อให้ภาพเหมือนเดิมทุกครั้งที่สร้าง)
        r = seed
        for z in levels:
            for i in range(bays):
                for k in range(2):
                    r = (r * 1103515245 + 12345) & 0x7FFFFFFF
                    if r % 7 == 0:
                        continue
                    w = 9 + (r >> 5) % 3
                    h = 6 + (r >> 9) % 5
                    tone = "y" if (r >> 13) % 11 == 0 else "b"
                    bx = x0 + i * bay + 1.5 + k * 11
                    self.box(bx, y0 + 1.2, z + 0.6, min(w, 10.5), depth - 2.4, min(h, 11), tone)
        # เสาและคานด้านหน้า (ทับสินค้า ให้ดูเป็นโครงเหล็กจริง)
        for i in range(bays + 1):
            self.line((x0 + i * bay, y0 + depth, 0), (x0 + i * bay, y0 + depth, height), "up f")
            self.line((x0 + i * bay, y0, height), (x0 + i * bay, y0 + depth, height), "beam")
        for z in levels[1:]:
            self.line((x0, y0 + depth, z), (x0 + L, y0 + depth, z), "beam f")

    def forklift(self, x, y):
        self.box(x, y, 0.8, 14, 9, 7, "y")          # ตัวรถ
        self.box(x + 2, y + 1, 7.8, 7, 7, 1, "y")   # หลังคาห้องคนขับ (ต่ำ)
        for dx, dy in ((2, 1), (2, 7), (8, 1), (8, 7)):
            self.line((x + dx, y + dy, 7.8), (x + dx, y + dy, 15), "cage")
        self.box(x + 1.5, y + 0.5, 15, 8, 8, 0.8, "k")  # หลังคา
        self.box(x + 14, y + 1, 0.8, 1.2, 1.2, 20, "k")  # เสายก
        self.box(x + 14, y + 6.8, 0.8, 1.2, 1.2, 20, "k")
        self.box(x + 15.2, y + 2, 2, 9, 1.2, 0.6, "k")   # งา
        self.box(x + 15.2, y + 5.8, 2, 9, 1.2, 0.6, "k")
        self.box(x + 16, y + 0.8, 2.6, 8, 7.4, 6, "y")   # พาเลทสินค้าบนงา
        for dx, dy in ((2.5, 9), (11, 9)):              # ล้อด้านหน้า
            cx, cy = self.p(x + dx, y + dy, 1.6)
            self.out.append(f'<ellipse class="wheel" cx="{cx:.2f}" cy="{cy:.2f}" rx="{1.9 * self.s:.2f}" ry="{1.7 * self.s:.2f}"/>')

    def floor_grid(self, x0, y0, x1, y1, step):
        x = x0
        while x <= x1:
            self.line((x, y0, 0), (x, y1, 0), "grid")
            x += step
        y = y0
        while y <= y1:
            self.line((x0, y, 0), (x1, y, 0), "grid")
            y += step


def scene_svg():
    # ฉากล่าง: ทางเดินกลางคลัง — ชั้นวางขวา + ชั้นวางซ้าย (กลับด้าน) บรรจบกลางหน้า รถยกอยู่กลางทางเดิน
    Y0 = 230
    floor = Iso(ox=105, oy=Y0, s=0.9)
    floor.floor_grid(-60, -60, 160, 160, 12)
    right = Iso(ox=105, oy=Y0, s=0.9)
    right.rack(8, -14, 5, seed=7)
    right.rack(8, -44, 5, seed=13)
    left = Iso(ox=105, oy=Y0, s=0.9)   # วาดแบบเดียวกันแล้วกลับซ้าย-ขวา
    left.rack(8, -14, 5, seed=3)
    left.rack(8, -44, 5, seed=17)
    mid = Iso(ox=105, oy=Y0, s=0.9)
    mid.forklift(50, 48)
    # ฉากบน: ชั้นวางจาง ๆ สองมุมบน
    hi = Iso(ox=150, oy=-30, s=0.8)
    hi.rack(-20, 0, 6, seed=5)
    hi.rack(-20, 34, 6, seed=9)
    hi2 = Iso(ox=150, oy=-30, s=0.8)
    hi2.rack(-20, 0, 6, seed=21)
    hi2.rack(-20, 34, 6, seed=23)
    flip = 'transform="translate(210 0) scale(-1 1)"'
    lo = type("G", (), {})()
    lo.out = (floor.out + right.out + [f'<g {flip}>'] + left.out + ["</g>"] + mid.out)
    hi.out = hi.out + [f'<g {flip}>'] + hi2.out + ["</g>"]
    return (
        '<svg class="scene" viewBox="0 0 210 297" preserveAspectRatio="xMidYMid slice" xmlns="http://www.w3.org/2000/svg">'
        '<defs>'
        '<linearGradient id="fadeLo" x1="0" y1="0" x2="0" y2="1">'
        '<stop offset="0" stop-color="#fff" stop-opacity="0"/>'
        '<stop offset=".66" stop-color="#fff" stop-opacity="0"/>'
        '<stop offset=".82" stop-color="#fff" stop-opacity=".95"/>'
        '<stop offset="1" stop-color="#fff" stop-opacity="1"/></linearGradient>'
        '<mask id="mLo"><rect width="210" height="297" fill="url(#fadeLo)"/></mask>'
        '<linearGradient id="fadeHi" x1="0" y1="0" x2="0" y2="1">'
        '<stop offset="0" stop-color="#fff" stop-opacity=".75"/>'
        '<stop offset=".22" stop-color="#fff" stop-opacity=".25"/>'
        '<stop offset=".34" stop-color="#fff" stop-opacity="0"/></linearGradient>'
        '<mask id="mHi"><rect width="210" height="297" fill="url(#fadeHi)"/></mask>'
        '</defs>'
        f'<g mask="url(#mHi)" class="hi">{"".join(hi.out)}</g>'
        f'<g mask="url(#mLo)">{"".join(lo.out)}</g>'
        '</svg>'
    )


# ---------- ไอคอนเส้น (stroke) ----------
ICON_CAM = '<svg viewBox="0 0 24 24"><path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.6"/></svg>'
ICON_SCAN = '<svg viewBox="0 0 24 24"><path d="M4 9V4h5M15 4h5v5M20 15v5h-5M9 20H4v-5"/><path d="M8 8h3v3H8zM13 13h3v3h-3zM13 8h3M8 16h3"/></svg>'
ICON_PIN = '<svg viewBox="0 0 24 24"><path d="M12 21s-6-5.6-6-11a6 6 0 0 1 12 0c0 5.4-6 11-6 11z"/><circle cx="12" cy="10" r="2.3"/></svg>'
ICON_GLOBE = '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3c2.6 2.6 3.8 5.6 3.8 9s-1.2 6.4-3.8 9c-2.6-2.6-3.8-5.6-3.8-9S9.4 5.6 12 3z"/></svg>'


def build_html():
    qr, n = qr_svg()
    return f"""<!doctype html>
<html lang="th"><head><meta charset="utf-8">
<title>ป้าย QR หาของในคลัง (A4)</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Kanit:wght@500;600;700&family=Sarabun:wght@400;600;700&display=block" rel="stylesheet">
<style>
@page {{ size: A4; margin: 0; }}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html, body {{ background: #3a3f46; }}
body {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
:root {{
  --ink: #0a0f16; --ink2: #111a26; --steel: #2c4361; --steel-hi: #4d6d93;
  --safety: #FFC20E; --safety-lo: #e0a400; --fg: #ffffff; --muted: #a9b8ca;
}}
.page {{
  position: relative; width: 210mm; height: 297mm; margin: 0 auto; overflow: hidden;
  background:
    radial-gradient(95mm 85mm at 50% 47%, rgba(255,194,14,.20), rgba(255,194,14,0) 70%),
    radial-gradient(160mm 120mm at 50% 105%, #16263a 0%, rgba(22,38,58,0) 70%),
    linear-gradient(180deg, #0d1521 0%, var(--ink) 45%, #070b11 100%);
  color: var(--fg); font-family: 'Sarabun', sans-serif;
}}
@media screen {{ .page {{ margin: 10mm auto; box-shadow: 0 10px 40px rgba(0,0,0,.5); }} }}

/* ฉากคลัง */
.scene {{ position: absolute; inset: 0; width: 100%; height: 100%; }}
.scene .grid {{ stroke: #1d3049; stroke-width: .25; }}
.scene .up {{ stroke: #3d5b80; stroke-width: .7; }}
.scene .up.f {{ stroke: #5b81ad; stroke-width: .9; }}
.scene .beam {{ stroke: #e0a400; stroke-width: .55; opacity: .75; }}
.scene .beam.f {{ stroke: #FFC20E; stroke-width: .75; opacity: .95; }}
.scene .bT {{ fill: #29425f; stroke: #4f74a0; stroke-width: .18; }}
.scene .bL {{ fill: #172a40; stroke: #3b5a80; stroke-width: .18; }}
.scene .bR {{ fill: #0f1d2e; stroke: #34506f; stroke-width: .18; }}
.scene .yT {{ fill: #ffd84a; stroke: #fff0a8; stroke-width: .2; }}
.scene .yL {{ fill: #f2ad00; stroke: #ffd84a; stroke-width: .2; }}
.scene .yR {{ fill: #b88400; stroke: #e0a400; stroke-width: .2; }}
.scene .kT {{ fill: #3a4656; }} .scene .kL {{ fill: #252f3c; }} .scene .kR {{ fill: #1a222c; }}
.scene .cage {{ stroke: #c9d4e2; stroke-width: .45; }}
.scene .wheel {{ fill: #05080c; stroke: #3a4656; stroke-width: .3; }}
.scene .hi {{ opacity: .9; }}

/* แถบเหลือง-ดำ */
.hazard {{ position: absolute; left: 0; right: 0; height: 7mm;
  background: repeating-linear-gradient(-45deg, var(--safety) 0 5mm, #0b1017 5mm 10mm); }}
.hazard.top {{ top: 0; }} .hazard.bot {{ bottom: 0; }}

.content {{ position: absolute; inset: 7mm 0; display: flex; flex-direction: column; align-items: center; }}

/* หัว */
.eyebrow {{ margin-top: 14mm; display: flex; align-items: center; gap: 3mm;
  font: 600 10.5pt 'Kanit', sans-serif; letter-spacing: .32em; color: var(--safety); }}
.eyebrow i {{ display: block; width: 14mm; height: .5mm; background: var(--safety); opacity: .7; }}
h1 {{ margin-top: 2mm; font: 700 64pt/1.15 'Kanit', sans-serif; letter-spacing: .01em;
  text-shadow: 0 1.2mm 4mm rgba(0,0,0,.65); }}
h1 b {{ color: var(--safety); font-weight: 700; }}
.sub {{ margin-top: 1mm; font: 600 17pt/1.35 'Sarabun', sans-serif; color: #dfe7f1; text-align: center; }}

/* QR */
.qrwrap {{ position: relative; margin-top: 9mm; width: 112mm; height: 112mm; }}
.qrcard {{ position: absolute; inset: 0; background: #fff; border-radius: 5mm; padding: 2mm;
  box-shadow: 0 0 0 1.4mm #0b1017, 0 0 0 2.2mm var(--safety), 0 4mm 18mm rgba(0,0,0,.6), 0 0 26mm rgba(255,194,14,.35); }}
.qr {{ display: block; width: 100%; height: 100%; }}
.corner {{ position: absolute; width: 16mm; height: 16mm; border: 2.2mm solid var(--safety); }}
.c1 {{ left: -7.5mm; top: -7.5mm; border-right: 0; border-bottom: 0; border-top-left-radius: 4mm; }}
.c2 {{ right: -7.5mm; top: -7.5mm; border-left: 0; border-bottom: 0; border-top-right-radius: 4mm; }}
.c3 {{ left: -7.5mm; bottom: -7.5mm; border-right: 0; border-top: 0; border-bottom-left-radius: 4mm; }}
.c4 {{ right: -7.5mm; bottom: -7.5mm; border-left: 0; border-top: 0; border-bottom-right-radius: 4mm; }}
.chip {{ position: absolute; left: 50%; bottom: -6.5mm; transform: translateX(-50%);
  display: flex; align-items: center; gap: 2.2mm; padding: 2mm 6mm 2.2mm;
  background: var(--safety); color: #1a1300; border-radius: 99mm; border: .8mm solid #0b1017;
  font: 700 15pt/1 'Kanit', sans-serif; white-space: nowrap; box-shadow: 0 2mm 6mm rgba(0,0,0,.5); }}
.chip svg {{ width: 6.5mm; height: 6.5mm; fill: none; stroke: #1a1300; stroke-width: 2.1; stroke-linecap: round; stroke-linejoin: round; }}

/* ขั้นตอน */
.steps {{ margin-top: 15mm; width: 182mm; display: grid; grid-template-columns: repeat(3, 1fr); gap: 4mm; }}
.step {{ position: relative; padding: 4.5mm 3.5mm 4mm 3.5mm; border-radius: 4mm; text-align: center;
  background: linear-gradient(180deg, rgba(23,36,54,.92), rgba(10,16,24,.92));
  border: .35mm solid #2f4766; box-shadow: inset 0 .4mm 0 rgba(255,255,255,.08), 0 2mm 6mm rgba(0,0,0,.45); }}
.num {{ position: absolute; top: -3.6mm; left: 50%; transform: translateX(-50%);
  width: 7.2mm; height: 7.2mm; border-radius: 50%; background: var(--safety); color: #1a1300;
  font: 700 12pt/7.2mm 'Kanit', sans-serif; border: .6mm solid #0b1017; }}
.ico {{ width: 13mm; height: 13mm; margin: 1mm auto 1.5mm; border-radius: 3mm; display: grid; place-items: center;
  background: linear-gradient(160deg, #22364f, #0f1a28); border: .3mm solid #3d5b80; }}
.ico svg {{ width: 8.5mm; height: 8.5mm; fill: none; stroke: var(--safety); stroke-width: 1.7; stroke-linecap: round; stroke-linejoin: round; }}
.step h3 {{ font: 600 13.5pt/1.25 'Kanit', sans-serif; color: #fff; }}
.step p {{ margin-top: .8mm; font: 400 10.5pt/1.35 'Sarabun', sans-serif; color: var(--muted); }}

/* ลิงก์ */
.link {{ margin-top: 6mm; display: flex; align-items: center; gap: 3mm;
  padding: 2.6mm 6mm; border-radius: 99mm; background: rgba(8,13,20,.85); border: .35mm solid #34506f; }}
.link svg {{ width: 6mm; height: 6mm; fill: none; stroke: var(--safety); stroke-width: 1.7; }}
.link span {{ font: 400 11pt 'Sarabun', sans-serif; color: var(--muted); }}
.link b {{ font: 600 14pt 'Kanit', sans-serif; letter-spacing: .03em; color: #fff; }}
</style></head>
<body>
<div class="page">
  {scene_svg()}
  <div class="hazard top"></div>
  <div class="content">
    <div class="eyebrow"><i></i>WAREHOUSE FINDER<i></i></div>
    <h1>LAY OUT <b>STORE</b></h1>

    <div class="qrwrap">
      <div class="qrcard">{qr}</div>
      <span class="corner c1"></span><span class="corner c2"></span>
      <span class="corner c3"></span><span class="corner c4"></span>
      <div class="chip">{ICON_SCAN}สแกนที่นี่</div>
    </div>

    <div class="steps">
      <div class="step"><div class="num">1</div><div class="ico">{ICON_CAM}</div>
        <h3>เปิดกล้องมือถือ</h3><p>ส่องให้เห็น QR ทั้งกรอบ</p></div>
      <div class="step"><div class="num">2</div><div class="ico">{ICON_SCAN}</div>
        <h3>แตะลิงก์ที่ขึ้นมา</h3><p>ไม่ต้องติดตั้งแอป</p></div>
      <div class="step"><div class="num">3</div><div class="ico">{ICON_PIN}</div>
        <h3>พิมพ์รหัสหรือชื่อ</h3><p>ผังคลังจะกะพริบบอกตำแหน่ง</p></div>
    </div>

    <div class="link">{ICON_GLOBE}<span>สแกนไม่ได้? เปิดเว็บ</span><b>{URL_TEXT}</b></div>
  </div>
  <div class="hazard bot"></div>
</div>
</body></html>
"""


def render():
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(build_html())
    url = "file:///" + OUT_HTML.replace("\\", "/")
    subprocess.run(
        [CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
         "--virtual-time-budget=8000", f"--print-to-pdf={OUT_PDF}", url],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    import pymupdf
    doc = pymupdf.open(OUT_PDF)
    doc[0].get_pixmap(dpi=300).save(OUT_PNG)
    print("pages:", len(doc), "size(pt):", doc[0].rect)
    print("wrote", OUT_HTML, OUT_PDF, OUT_PNG, sep="\n  ")


if __name__ == "__main__":
    if "--html-only" in sys.argv:
        with open(OUT_HTML, "w", encoding="utf-8") as f:
            f.write(build_html())
    else:
        render()
