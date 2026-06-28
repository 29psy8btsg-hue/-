#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
9月号 Excel 最終更新スクリプト
- 6月装飾（あじさい・梅雨フレーム）→ 削除
- 9月行事別イラストをカレンダーセルに追加
- 9月季節装飾（紅葉・お月見）を追加
- 行事名ふりがな（rPh）を修正
"""
import os, re, shutil, zipfile, io
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw

# ── パス設定 ──────────────────────────────────────────────────────────────
SRC  = "/home/user/-/夢通信_2026年09月号.xlsx"
WRK  = "/tmp/sep_final_work"
DIR  = os.path.join(WRK, "xlsx_x")
OUT  = "/home/user/-/夢通信_2026年09月号.xlsx"
ICO  = "/home/user/-/assets/ico"

# ── 名前空間 ──────────────────────────────────────────────────────────────
XDR  = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A    = "http://schemas.openxmlformats.org/drawingml/2006/main"
R    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG  = "http://schemas.openxmlformats.org/package/2006/relationships"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XMLSP= "{http://www.w3.org/XML/1998/namespace}space"

def reg_ns(path):
    txt = open(path, encoding="utf-8", errors="replace").read()
    for pref, uri in re.findall(r'xmlns:?(\w*)="([^"]+)"', txt):
        try: ET.register_namespace(pref, uri)
        except Exception: pass

# ── 準備 ──────────────────────────────────────────────────────────────────
if os.path.exists(WRK): shutil.rmtree(WRK)
os.makedirs(DIR)
with zipfile.ZipFile(SRC) as z: z.extractall(DIR)
print("展開 OK")

# ── 1. 季節装飾イメージ作成 ───────────────────────────────────────────────
MEDIA = os.path.join(DIR, "xl/media")

def load_icon(name, size=(280, 280)):
    """assets/ico からアイコン読み込み・リサイズ (RGBA)"""
    img = Image.open(os.path.join(ICO, name)).convert("RGBA")
    img.thumbnail(size, Image.LANCZOS)
    out = Image.new("RGBA", size, (0,0,0,0))
    x = (size[0] - img.width)  // 2
    y = (size[1] - img.height) // 2
    out.paste(img, (x, y), img)
    return out

def save_media(img, name):
    path = os.path.join(MEDIA, name)
    img.save(path, "PNG")
    return path

# 1-a 大型ヘッダー装飾（行2列5 → 9月: 紅葉＋お月見コンポジット）
def make_autumn_big():
    W, H = 300, 240
    bg = Image.new("RGBA", (W, H), (0,0,0,0))
    maple = load_icon("maple.png", (160,160))
    moon  = load_icon("tsukimi.png", (130,130))
    bg.paste(maple, (0,  40),  maple)
    bg.paste(moon,  (150, 80), moon)
    # 小さい紅葉を背景に
    sm = load_icon("maple.png", (70,70))
    bg.paste(sm, (200, 10), sm)
    bg.paste(sm, (20,  150), sm)
    return bg

save_media(make_autumn_big(), "image34.png")
print("image34 (ヘッダー装飾) OK")

# 1-b 行19 小型装飾 (4枚)
for num, ico in [(35,"maple.png"), (36,"moon.png"),
                 (37,"tsukimi.png"), (38,"star.png")]:
    img = load_icon(ico, (100, 80))
    save_media(img, f"image{num}.png")
print("image35-38 (行19 小型装飾) OK")

# 1-c 秋フレーム（行34 用 × 1種 → rId で共有）
def make_autumn_frame(W=540, H=390):
    bg = Image.new("RGBA", (W, H), (255, 255, 255, 255))
    draw = ImageDraw.Draw(bg)
    # 外枠
    draw.rounded_rectangle([4, 4, W-5, H-5], radius=18,
                            outline="#E07820", width=6)
    draw.rounded_rectangle([12, 12, W-13, H-13], radius=14,
                            outline="#F5B942", width=3)
    # 四隅・エッジに紅葉アイコンを配置
    maple = load_icon("maple.png", (64, 64))
    moon  = load_icon("moon.png",  (52, 52))
    star  = load_icon("star.png",  (44, 44))
    positions_maple = [(6,6),(W-70,6),(6,H-70),(W-70,H-70)]
    positions_moon  = [(W//2-26, 4), (W//2-26, H-56)]
    positions_star  = [(80,4),(160,4),(240,4),(320,4),(400,4),
                       (80,H-48),(160,H-48),(240,H-48),(320,H-48),(400,H-48)]
    for (px,py) in positions_maple: bg.paste(maple,(px,py),maple)
    for (px,py) in positions_moon:  bg.paste(moon, (px,py),moon)
    for (px,py) in positions_star:  bg.paste(star, (px,py),star)
    return bg

save_media(make_autumn_frame(), "image39.png")
print("image39 (秋フレーム) OK")

# 1-d カレンダーイベントイラスト (9枚)
event_icons = [
    # (file_num, icon_file)
    (40, "rhythm.png"),   # D11 リズム遊び週間
    (41, "maple.png"),    # G11 秋の自然観察
    (42, "bousai.png"),   # D13 防災週間
    (43, "imo.png"),      # G13 芋掘り
    (44, "present.png"),  # D15 敬老の日プレゼント作り週間
    (45, "ohagi.png"),    # G15 おはぎづくり
    (46, "tsukimi.png"),  # D17 お月見会
    (47, "kouen.png"),    # G17 アスレチック公園
    (48, "bowling.png"),  # D19 ボッチャ大会
]
for num, ico in event_icons:
    img = load_icon(ico, (240, 240))
    save_media(img, f"image{num}.png")
print("image40-48 (イベントアイコン) OK")

# ── 2. drawing1.xml.rels 更新 ────────────────────────────────────────────
RELS = os.path.join(DIR, "xl/drawings/_rels/drawing1.xml.rels")
reg_ns(RELS)
rt = ET.parse(RELS); rroot = rt.getroot()

# 削除する rId（6月装飾）
JUNE_RIDS = {"rId5","rId19","rId20","rId21","rId22","rId30","rId31","rId32","rId33"}
for rel in list(rroot):
    if rel.get("Id") in JUNE_RIDS:
        rroot.remove(rel)

# 新規 rId 追加
new_rels = [
    ("rId34","image34.png"),  # ヘッダー装飾
    ("rId35","image35.png"),  # 行19 maple
    ("rId36","image36.png"),  # 行19 moon
    ("rId37","image37.png"),  # 行19 tsukimi
    ("rId38","image38.png"),  # 行19 star
    ("rId39","image39.png"),  # 秋フレーム（共有）
    ("rId40","image40.png"),  # rhythm
    ("rId41","image41.png"),  # maple (秋自然)
    ("rId42","image42.png"),  # bousai
    ("rId43","image43.png"),  # imo
    ("rId44","image44.png"),  # present
    ("rId45","image45.png"),  # ohagi
    ("rId46","image46.png"),  # tsukimi
    ("rId47","image47.png"),  # kouen
    ("rId48","image48.png"),  # bowling
]
for rid, fname in new_rels:
    el = ET.SubElement(rroot, f"{{{PKG}}}Relationship")
    el.set("Id", rid)
    el.set("Type", f"{R}/image")
    el.set("Target", f"../media/{fname}")

rt.write(RELS, encoding="UTF-8", xml_declaration=True)
print("drawing1.xml.rels 更新 OK")

# ── 3. drawing1.xml 更新 ─────────────────────────────────────────────────
DR = os.path.join(DIR, "xl/drawings/drawing1.xml")
reg_ns(DR)
dt = ET.parse(DR); droot = dt.getroot()

# 3-a 6月装飾アンカーを削除
JUNE_NAMES = {"図 19","図 8","図 17","図 23","図 25",
              "図 55","図 59","図 61","図 62"}
removed = 0
for an in list(droot):
    cnv = an.find(f".//{{{XDR}}}cNvPr")
    if cnv is not None and cnv.get("name") in JUNE_NAMES:
        droot.remove(an); removed += 1
print(f"削除したアンカー数: {removed}")

# 3-b アンカーテンプレート生成ヘルパー
MSVIZ  = "http://schemas.microsoft.com/office/drawing/2010/main"
MSCRID = "http://schemas.microsoft.com/office/drawing/2014/main"
_anchor_id_counter = [100]

def make_anchor(rid, name,
                fc, fr, fco, fro,
                tc, tr, tco, tro):
    """twoCellAnchor を生成して返す"""
    aid = _anchor_id_counter[0]; _anchor_id_counter[0] += 1
    an = ET.Element(f"{{{XDR}}}twoCellAnchor")
    an.set("editAs", "oneCell")
    # from
    f = ET.SubElement(an, f"{{{XDR}}}from")
    for tag, val in [("col",str(fc)),("colOff",str(fco)),
                     ("row",str(fr)),("rowOff",str(fro))]:
        ET.SubElement(f, f"{{{XDR}}}{tag}").text = val
    # to
    t = ET.SubElement(an, f"{{{XDR}}}to")
    for tag, val in [("col",str(tc)),("colOff",str(tco)),
                     ("row",str(tr)),("rowOff",str(tro))]:
        ET.SubElement(t, f"{{{XDR}}}{tag}").text = val
    # pic
    pic = ET.SubElement(an, f"{{{XDR}}}pic")
    nvp = ET.SubElement(pic, f"{{{XDR}}}nvPicPr")
    cnv = ET.SubElement(nvp, f"{{{XDR}}}cNvPr")
    cnv.set("id", str(aid)); cnv.set("name", name)
    cnvp = ET.SubElement(nvp, f"{{{XDR}}}cNvPicPr")
    pl = ET.SubElement(cnvp, f"{{{A}}}picLocks")
    pl.set("noChangeAspect", "1")
    bf = ET.SubElement(pic, f"{{{XDR}}}blipFill")
    blip = ET.SubElement(bf, f"{{{A}}}blip")
    blip.set(f"{{{R}}}embed", rid)
    blip.set("cstate", "print")
    stretch = ET.SubElement(bf, f"{{{A}}}stretch")
    ET.SubElement(stretch, f"{{{A}}}fillRect")
    sp = ET.SubElement(pic, f"{{{XDR}}}spPr")
    xfrm = ET.SubElement(sp, f"{{{A}}}xfrm")
    ET.SubElement(xfrm, f"{{{A}}}off").set("x","0"); ET.SubElement(xfrm, f"{{{A}}}off")  # dummy
    ET.SubElement(sp, f"{{{A}}}prstGeom").set("prst","rect")
    ET.SubElement(ET.SubElement(sp, f"{{{A}}}prstGeom"), f"{{{A}}}avLst") \
        if sp.find(f"{{{A}}}prstGeom") is None else None
    # fix: rebuild spPr
    sp.clear()
    xfrm = ET.SubElement(sp, f"{{{A}}}xfrm")
    off = ET.SubElement(xfrm, f"{{{A}}}off"); off.set("x","0"); off.set("y","0")
    ext = ET.SubElement(xfrm, f"{{{A}}}ext"); ext.set("cx","500000"); ext.set("cy","400000")
    pg = ET.SubElement(sp, f"{{{A}}}prstGeom"); pg.set("prst","rect")
    ET.SubElement(pg, f"{{{A}}}avLst")
    ET.SubElement(an, f"{{{XDR}}}clientData")
    return an

# 3-c 新規アンカーを追加
# ── ヘッダー装飾（row=2, col=5 → col=6, row=5: 元あじさいと同位置）
droot.append(make_anchor("rId34", "秋装飾ヘッダー",
    5, 2, 806711, 68038,   # from
    6, 5, 719234, 21028))  # to

# ── 行19 小型装飾（元あじさい小と同位置）
droot.append(make_anchor("rId35","秋装飾_紅葉",  0,19,644594,91741,  1,20,355762,172641))
droot.append(make_anchor("rId36","秋装飾_月",   2,19,512538,109290, 2,20,928829,184547))
droot.append(make_anchor("rId37","秋装飾_月見", 3,19,969065,74555,  4,20,375047,125853))
droot.append(make_anchor("rId38","秋装飾_星",   6,19,338241,76574,  6,20,762000,157453))

# ── 行34 秋フレーム（元6月フレームと同位置）
droot.append(make_anchor("rId39","秋フレーム_1",  3,34,760843,110122, 6,35,17389,162935))
droot.append(make_anchor("rId39","秋フレーム_2",  0,34,107578,106797, 2,35,611842,159610))
droot.append(make_anchor("rId39","秋フレーム_3",  2,34,576060,120863, 3,35,360906,170020))
droot.append(make_anchor("rId39","秋フレーム_4",  6,34,32149,135315,  6,35,394138,164224))

# ── カレンダー内イベントイラスト（行8〜18内）
# 配置座標：from/to が同じ col/row（セル内配置）
# colOff/rowOff は June 版に倣い適度なパディング
CAL_ICONS = [
    # (rId, name, fc, fr, fco, fro, tc, tr, tco, tro)
    ("rId40","イラスト_リズム遊び",  3,10, 80000, 30000, 3,11, 880000, 470000),
    ("rId41","イラスト_秋の自然観察",6,10, 60000, 30000, 6,11, 870000, 470000),
    ("rId42","イラスト_防災週間",    3,12, 80000, 30000, 3,13, 880000, 470000),
    ("rId43","イラスト_芋掘り",      6,12, 60000, 30000, 6,13, 870000, 470000),
    ("rId44","イラスト_敬老の日",    3,14, 80000, 30000, 3,15, 880000, 470000),
    ("rId45","イラスト_おはぎ",      6,14, 60000, 30000, 6,15, 870000, 470000),
    ("rId46","イラスト_お月見",      3,16, 80000, 30000, 3,17, 880000, 470000),
    ("rId47","イラスト_アスレチック",6,16, 60000, 30000, 6,17, 870000, 470000),
    ("rId48","イラスト_ボッチャ",    3,18, 80000, 30000, 3,19, 880000, 470000),
]
for args in CAL_ICONS:
    droot.append(make_anchor(*args))

dt.write(DR, encoding="UTF-8", xml_declaration=True)
print("drawing1.xml 更新 OK")

# ── 4. sharedStrings.xml ふりがな修正 ────────────────────────────────────
SS = os.path.join(DIR, "xl/sharedStrings.xml")
reg_ns(SS)
sst = ET.parse(SS); sroot = sst.getroot()
sis = list(sroot)

def q(tag): return f"{{{MAIN}}}{tag}"

def make_si(text, ruby_list=None, font_id="16"):
    """
    シェアードストリング要素を作成
    ruby_list: [(sb, eb, reading), ...]  rPh リスト
    """
    si = ET.Element(q("si"))
    t = ET.SubElement(si, q("t"))
    t.set(XMLSP, "preserve")
    t.text = text
    if ruby_list:
        for sb, eb, rd in ruby_list:
            rph = ET.SubElement(si, q("rPh"))
            rph.set("sb", str(sb)); rph.set("eb", str(eb))
            rt = ET.SubElement(rph, q("t"))
            rt.text = rd
    pp = ET.SubElement(si, q("phoneticPr"))
    pp.set("fontId", font_id)
    pp.set("type",   "Hiragana")
    pp.set("alignment", "distributed")
    return si

# ふりがな付き 9月イベント定義
# 「敬老の日\nプレゼント作り週間」の文字位置:
#   敬(0)老(1)の(2)日(3)\n(4)プ(5)レ(6)ゼ(7)ン(8)ト(9)作(10)り(11)週(12)間(13)
GREETING_TEXT = (
    "日頃より夢門塾ゆうゆう日吉をご利用いただき誠にありがとうございます。"
    "朝晩は過ごしやすくなり、秋の気配が感じられる季節になりました。"
    "気温差で体調を崩さないよう、体調第一で過ごしましょう！"
)
SI_NEW = {
    8:  make_si("秋の自然観察", [(0,1,"あき"),(2,4,"しぜん"),(4,6,"かんさつ")]),
    9:  make_si("アスレチック公園",   [(6,8,"こうえん")]),
    10: make_si("ボッチャ大会",        [(4,6,"たいかい")]),
    11: make_si("芋掘り",              [(0,1,"いも"),(1,2,"ほ")]),
    13: make_si("リズム遊び週間",      [(3,4,"あそ"),(5,7,"しゅうかん")]),
    14: make_si("お月見会",            [(1,3,"つきみ"),(3,4,"かい")]),
    15: make_si("敬老の日\nプレゼント作り週間",
                [(0,2,"けいろう"),(3,4,"ひ"),(10,11,"つく"),(12,14,"しゅうかん")]),
    16: make_si(GREETING_TEXT),
    17: make_si("おはぎづくり"),
}

for idx, new_si in SI_NEW.items():
    # 既存 si を置換
    old = sis[idx]
    parent = sroot
    pos = list(parent).index(old)
    parent.remove(old)
    parent.insert(pos, new_si)

sst.write(SS, encoding="UTF-8", xml_declaration=True)
print("sharedStrings.xml ふりがな修正 OK")

# ── 5. 不要メディアファイルを削除 ─────────────────────────────────────────
june_media = ["image5.jpg",
              "image19.jpeg","image20.jpeg","image21.jpeg","image22.jpeg",
              "image30.jpeg","image31.jpeg","image32.jpeg","image33.jpeg"]
for f in june_media:
    p = os.path.join(MEDIA, f)
    if os.path.exists(p): os.remove(p)
print(f"{len(june_media)} 個の6月メディアを削除")

# ── 6. 再 ZIP ────────────────────────────────────────────────────────────
if os.path.exists(OUT): os.remove(OUT)
ct = os.path.join(DIR, "[Content_Types].xml")
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.write(ct, "[Content_Types].xml")
    for root_d, _, files in os.walk(DIR):
        for f in files:
            full = os.path.join(root_d, f)
            arc  = os.path.relpath(full, DIR)
            if arc == "[Content_Types].xml": continue
            z.write(full, arc)

size = os.path.getsize(OUT)
print(f"出力 OK: {OUT}  ({size:,} bytes)")
