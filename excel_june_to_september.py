#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""令和8年6月のExcelを9月版に変換（ロゴ・書式・図形を保持、カレンダーのイラストは削除）"""
import re, os, shutil, zipfile
import xml.etree.ElementTree as ET

SRC_DIR = "xlsx_x"
OUT_XLSX = "夢通信_2026年09月号.xlsx"

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
A    = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR  = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"

def register_ns_from(path):
    txt = open(path, encoding="utf-8").read()
    for pref, uri in re.findall(r'xmlns:?(\w*)="([^"]+)"', txt):
        try: ET.register_namespace(pref, uri)
        except Exception: pass

# ---------------- 1) sharedStrings.xml（イベント名・あいさつ） ----------------
SS = os.path.join(SRC_DIR, "xl/sharedStrings.xml")
register_ns_from(SS)
sst = ET.parse(SS); sroot = sst.getroot()
sis = list(sroot)

GREETING = ("日頃より夢門塾ゆうゆう日吉をご利用いただき誠にありがとうございます。"
            "朝晩は過ごしやすくなり、秋の気配が感じられる季節になりました。"
            "気温差で体調を崩さないよう、体調第一で過ごしましょう！")

# index -> 新しい表示テキスト（ふりがなは省略してプレーン化）
SI_REPLACE = {
    8:  "秋の自然観察",
    9:  "アスレチック公園",
    10: "ボッチャ大会",
    11: "芋掘り",
    13: "リズム遊び週間",
    14: "お月見会",
    15: "敬老の日\nプレゼント作り週間",
    16: GREETING,
    17: "おはぎづくり",
    # 12 防災週間 は据え置き
}
for idx, newtext in SI_REPLACE.items():
    si = sis[idx]
    for ch in list(si): si.remove(ch)
    t = ET.SubElement(si, f"{{{MAIN}}}t")
    t.set(XMLSPACE, "preserve")
    t.text = newtext
sst.write(SS, encoding="UTF-8", xml_declaration=True)
print("sharedStrings 更新 OK")

# ---------------- 2) sheet1.xml（日付セルを-1、B10削除、D18=30追加） ----------------
SH = os.path.join(SRC_DIR, "xl/worksheets/sheet1.xml")
register_ns_from(SH)
sht = ET.parse(SH); wroot = sht.getroot()
def q(tag): return f"{{{MAIN}}}{tag}"

rows = {}
sheetData = wroot.find(q("sheetData"))
for row in sheetData.findall(q("row")):
    rows[row.get("r")] = row

def col_of(ref): return re.match(r"([A-Z]+)", ref).group(1)

for row in sheetData.findall(q("row")):
    for c in list(row.findall(q("c"))):
        t = c.get("t")
        v = c.find(q("v"))
        if t == "s" or v is None:        # 文字列セルや空セルは触らない
            continue
        try: num = int(v.text)
        except (TypeError, ValueError): continue
        if 1 <= num <= 31:               # 日付セル
            newnum = num - 1
            if newnum <= 0:              # B10(1)→削除
                row.remove(c)
            else:
                v.text = str(newnum)

# D18 = 30（既存の空セルD18があれば値を設定、無ければ作成）
row18 = rows["18"]
d18 = next((c for c in row18.findall(q("c")) if c.get("r") == "D18"), None)
if d18 is not None:
    if d18.get("t"): del d18.attrib["t"]   # 文字列型を解除
    v = d18.find(q("v"))
    if v is None:
        v = ET.SubElement(d18, q("v"))
    v.text = "30"
else:
    c18 = next((c for c in row18.findall(q("c")) if c.get("r") == "C18"), None)
    newc = ET.Element(q("c")); newc.set("r", "D18")
    if c18 is not None and c18.get("s") is not None: newc.set("s", c18.get("s"))
    nv = ET.SubElement(newc, q("v")); nv.text = "30"
    idx = list(row18).index(c18) + 1 if c18 is not None else len(list(row18))
    row18.insert(idx, newc)

sht.write(SH, encoding="UTF-8", xml_declaration=True)
print("sheet1 日付 更新 OK")

# ---------------- 3) drawing1.xml（月号・おしらせ・ねらい・イラスト削除） ----------------
DR = os.path.join(SRC_DIR, "xl/drawings/drawing1.xml")
register_ns_from(DR)
dt = ET.parse(DR); droot = dt.getroot()

def shape_by_name(root, name):
    for an in root:
        cnv = an.find(f".//{{{XDR}}}cNvPr")
        if cnv is not None and cnv.get("name") == name:
            return an
    return None

def texts_of(anchor):
    return anchor.findall(f".//{{{A}}}t")

# 3-1 ヘッダー（フレーム7）：６月号→９月号
frame = shape_by_name(droot, "フレーム 7")
for t in texts_of(frame):
    if t.text == "６月号":
        t.text = "９月号"

# 3-2 おしらせ（テキスト ボックス 50）：実行順で置換
osu = shape_by_name(droot, "テキスト ボックス 50")
ts = texts_of(osu)
def set_text_at(elems, i, new):
    if 0 <= i < len(elems): elems[i].text = new
repl_osu = {
    1:  "9/5",
    5:  "「秋の自然観察」",
    14: "9/12(",
    17: "「芋掘り」",
    21: "軍手、汚れてもよい服装、お弁当、水筒、タオル、お勉強道具、着替え",
    26: "9/19(",
    29: "「クッキング（おはぎづくり）」",
    41: "9/26",
    45: "「アスレチック公園」",
}
for i, nt in repl_osu.items(): set_text_at(ts, i, nt)

# 3-3 ねらい（四角形: メモ 12）
nerai = shape_by_name(droot, "四角形: メモ 12")
tn = texts_of(nerai)
repl_nerai = {
    2:  "リズム遊び週間",
    4:  "音楽に合わせて全身を動かし、リズム感や表現力を養います。",
    # 6 防災週間 / 8 説明 は据え置き
    10: "敬老の日プレゼント作り週間",
    12: "製作活動を通して手先の器用さや集中力を養い、感謝の気持ちを伝えます。",
    14: "お月見会",
    16: "季節の行事に親しみ、友達と一緒に楽しむ経験を通して情緒を育みます。",
    18: "ボッチャ大会",
    20: "コントロール力や集中力を鍛えます。友達と協力することの大切さを学びます。",
}
for i, nt in repl_nerai.items(): set_text_at(tn, i, nt)

# 3-4 カレンダー内のイラスト（rId付き=画像）を削除（行9〜19）。ロゴ・図形・上下装飾は保持
removed = 0
for an in list(droot):
    blip = an.find(f".//{{{A}}}blip")
    if blip is None:
        continue
    frm = an.find(f"{{{XDR}}}from")
    if frm is None:
        continue
    r0 = int(frm.find(f"{{{XDR}}}row").text)  # 0-indexed
    if 8 <= r0 <= 18:                          # 1-indexed 行9〜19＝カレンダー
        droot.remove(an); removed += 1
print("削除した画像アンカー数:", removed)

dt.write(DR, encoding="UTF-8", xml_declaration=True)
print("drawing1 更新 OK")

# ---------------- 4) 再ZIP ----------------
if os.path.exists(OUT_XLSX): os.remove(OUT_XLSX)
with zipfile.ZipFile(OUT_XLSX, "w", zipfile.ZIP_DEFLATED) as z:
    # [Content_Types].xml を先頭に
    ct = os.path.join(SRC_DIR, "[Content_Types].xml")
    z.write(ct, "[Content_Types].xml")
    for root, _, files in os.walk(SRC_DIR):
        for f in files:
            full = os.path.join(root, f)
            arc = os.path.relpath(full, SRC_DIR)
            if arc == "[Content_Types].xml": continue
            z.write(full, arc)
print("出力:", OUT_XLSX)
