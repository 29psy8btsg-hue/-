#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""わたしの がんばりカード（夢門塾）を編集可能な PowerPoint(.pptx) で生成する。
A4 横・1スライド。テキストはすべて編集可能。"""
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "mumonjuku-logo.png")
FONT = "Meiryo"  # 全PC共通の見やすい日本語フォント。お好みで変更可

# 色
SKY   = RGBColor(0x4E,0xC5,0xE0); SKY_L = RGBColor(0x7D,0xD8,0xEC); SKY_D = RGBColor(0x2B,0xA8,0xC9)
NAVY  = RGBColor(0x1A,0x3A,0x6B); PINK  = RGBColor(0xE8,0x19,0x7D); PINK_L= RGBColor(0xF6,0x87,0xC0)
GREEN = RGBColor(0x22,0xC5,0x5E); GREEN_D=RGBColor(0x15,0x80,0x3d); ORANGE= RGBColor(0xFF,0x8C,0x42)
YELLOW= RGBColor(0xFF,0xC9,0x3C); PURPLE= RGBColor(0x8B,0x5C,0xF6); SOFT  = RGBColor(0xF3,0xF9,0xFC)
LINE  = RGBColor(0xC3,0xD0,0xDD); GRAY  = RGBColor(0x94,0xA3,0xB8); WHITE = RGBColor(0xFF,0xFF,0xFF)
DASH  = RGBColor(0xC7,0xD3,0xDF); INK_L = RGBColor(0x5b,0x71,0x87)

prs = Presentation()
prs.slide_width  = Cm(29.7)
prs.slide_height = Cm(21.0)
slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白
shapes = slide.shapes

def _noline(sp):
    sp.line.fill.background()

def rrect(x,y,w,h, fill=None, line=None, lw=1.0, radius=0.10, shadow=False):
    sp = shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x),Cm(y),Cm(w),Cm(h))
    try: sp.adjustments[0] = radius
    except Exception: pass
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None: _noline(sp)
    else: sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    return sp

def oval(x,y,d, fill=None, line=None, lw=1.0):
    sp = shapes.add_shape(MSO_SHAPE.OVAL, Cm(x),Cm(y),Cm(d),Cm(d))
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None: _noline(sp)
    else: sp.line.color.rgb = line; sp.line.width = Pt(lw)
    sp.shadow.inherit = False
    return sp

def text(x,y,w,h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, wrap=True, sp_after=2):
    """runs: list of lines; each line is list of (txt,size,bold,color) tuples."""
    tb = shapes.add_textbox(Cm(x),Cm(y),Cm(w),Cm(h)); tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left=Cm(0.1); tf.margin_right=Cm(0.1); tf.margin_top=Cm(0.02); tf.margin_bottom=Cm(0.02)
    for i,line in enumerate(runs):
        p = tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment = align; p.space_after=Pt(sp_after); p.space_before=Pt(0)
        for (t,sz,b,col) in line:
            r = p.add_run(); r.text=t; f=r.font
            f.size=Pt(sz); f.bold=b; f.color.rgb=col; f.name=FONT
    return tb

def dot_line(x,y,w, color=LINE):
    """記入用の点線"""
    ln = shapes.add_connector(2, Cm(x),Cm(y),Cm(x+w),Cm(y))
    ln.line.color.rgb=color; ln.line.width=Pt(1.2)
    d = ln.line._get_or_add_ln(); pd = d.makeelement(qn('a:prstDash'),{'val':'sysDash'}); d.append(pd)
    ln.shadow.inherit=False
    return ln

# ===== 背景 =====
bg = rrect(0.3,0.3,29.1,20.4, fill=WHITE, line=SKY, lw=2.2, radius=0.04)

# ===== 夢門塾 ヘッダー（空＋雲） =====
head = rrect(0.3,0.3,29.1,3.5, fill=SKY, line=None, radius=0.04)
# 雲（淡い白＝ブランドの空に溶け込む）
CLOUD = RGBColor(0xBF,0xE7,0xF3)
for (cx,cy,cd) in [(7.0,-1.4,3.2),(13.0,2.6,2.4),(22.5,-1.3,2.8)]:
    oval(cx,cy,cd, fill=CLOUD)
# ロゴ（白カードに乗せる）
rrect(0.9,0.7,3.1,2.7, fill=WHITE, line=None, radius=0.16)
if os.path.exists(LOGO):
    shapes.add_picture(LOGO, Cm(1.1),Cm(0.9), Cm(2.7),Cm(2.7))
# タイトル
text(4.5,0.75,15.5,1.7, [[("わたしの がんばりカード",28,True,WHITE)]], anchor=MSO_ANCHOR.BOTTOM)
sub = rrect(4.6,2.5,9.7,0.85, fill=RGBColor(0x7D,0xD8,0xEC), line=None, radius=0.5)
text(4.6,2.5,9.7,0.85, [[("まいにち ふりかえり ・ せいちょうの きろく",11,True,NAVY)]], align=PP_ALIGN.CENTER)
# 名前・期間
rrect(20.6,0.7,8.4,2.5, fill=WHITE, line=None, radius=0.12)
text(21.0,0.95,2.4,0.9, [[("なまえ",13,True,SKY_D)]], anchor=MSO_ANCHOR.MIDDLE)
dot_line(23.2,1.75,5.4)
text(21.0,1.95,2.4,0.9, [[("きかん",13,True,SKY_D)]], anchor=MSO_ANCHOR.MIDDLE)
dot_line(23.2,2.75,5.4)

# ===== 個別支援計画 =====
plan = [
    ("🎯 ながい もくひょう","（個別支援計画より）", SKY,  SOFT),
    ("🌱 今月の チャレンジ","（自分で きめる 近い 目標）", ORANGE, SOFT),
    ("💬 おうえん","（おうちの人・先生から）", PINK, SOFT),
]
px=0.7; pw=9.1; py=4.15; ph=1.9
for i,(t,sub,col,fill) in enumerate(plan):
    x = px + i*(pw+0.5)
    rrect(x,py,pw,ph, fill=fill, line=col, lw=1.4, radius=0.12)
    rrect(x,py+0.25,0.16,ph-0.5, fill=col, line=None, radius=0.5)  # 左アクセント
    text(x+0.35,py+0.18,pw-0.5,0.7, [[ (t,11,True,col),("  "+sub,8,False,GRAY) ]])
    dot_line(x+0.4,py+1.45,pw-0.8)

# ===== 目標ゲージ =====
text(0.7,6.25,8.2,0.8, [[("⭐ できた日に シールを はろう",16,True,PINK)]])
tag = rrect(9.0,6.38,6.5,0.66, fill=RGBColor(0xE6,0xF6,0xFB), line=None, radius=0.5)
text(9.0,6.38,6.5,0.66, [[("シールが ふえると せいちょうが 見える！",10,True,SKY_D)]], align=PP_ALIGN.CENTER)

gauges = [
    ("🗣️","おだやかな こえ","落ち着いた 声で 話す", SKY),
    ("😊","やさしい 言葉","ふわふわ 言葉を 使う", PINK_L),
    ("🤝","手・足は おだやかに","たたかない・けらない", PURPLE),
    ("✏️","すわって とりくむ","自分の 場所で おちついて", GREEN),
    ("⭐","きらり ポイント","とくい・成長した こと", YELLOW),
    ("🌱","今月の チャレンジ","自分で きめた こと", ORANGE),
]
gx=[0.7,15.25]; gw=13.75; gy0=7.05; gh=1.55; gvs=1.78
for i,(ico,n1,n2,col) in enumerate(gauges):
    col_i=i%2; row=i//2
    x=gx[col_i]; y=gy0+row*gvs
    rrect(x,y,gw,gh, fill=WHITE, line=col, lw=1.6, radius=0.14)
    # アイコン
    rrect(x+0.3,y+0.25,1.05,1.05, fill=SOFT, line=None, radius=0.22)
    text(x+0.3,y+0.18,1.05,1.05, [[(ico,17,False,NAVY)]], align=PP_ALIGN.CENTER)
    # 名前
    text(x+1.55,y+0.18,5.0,1.2, [[(n1,13,True,NAVY)],[(n2,8,False,GRAY)]], sp_after=0)
    # シール円 10個
    n=10; cd=0.60; gap=0.095
    start=x+6.5
    for k in range(n):
        cx=start+k*(cd+gap); cy=y+(gh-cd)/2
        if k==4:
            oval(cx,cy,cd, fill=RGBColor(0xFF,0xF6,0xE0), line=YELLOW, lw=1.4)
            text(cx-0.05,cy-0.02,cd+0.1,cd, [[("⭐",8,False,ORANGE)]], align=PP_ALIGN.CENTER)
        elif k==9:
            oval(cx,cy,cd, fill=RGBColor(0xFC,0xE7,0xF0), line=PINK, lw=1.4)
            text(cx-0.05,cy-0.02,cd+0.1,cd, [[("✦",9,True,PINK)]], align=PP_ALIGN.CENTER)
        else:
            oval(cx,cy,cd, fill=WHITE, line=DASH, lw=1.3)
            text(cx-0.05,cy-0.02,cd+0.1,cd, [[(str(k+1),8,True,GRAY)]], align=PP_ALIGN.CENTER)

# ===== 下段 =====
by=12.65; bh=3.0
# 見かた
lx=0.7; lw_=6.4
rrect(lx,by,lw_,bh, fill=SOFT, line=RGBColor(0xD6,0xE3,0xEC), lw=1.6, radius=0.10)
text(lx+0.35,by+0.15,lw_-0.6,0.7, [[("👀 シールの 見かた",12,True,INK_L)]])
# good bar
for k in range(5):
    oval(lx+0.4+k*0.62, by+1.05, 0.5, fill=YELLOW, line=None)
text(lx+3.7,by+0.95,2.6,0.7, [[("いっぱい ＝ とくい！",10,True,GREEN_D)]])
# few bar
oval(lx+0.4, by+1.95, 0.5, fill=YELLOW, line=None)
for k in range(1,5):
    oval(lx+0.4+k*0.62, by+1.95, 0.5, fill=WHITE, line=DASH, lw=1.3)
text(lx+3.7,by+1.85,2.6,0.9, [[("すくない ＝",10,True,ORANGE)],[("いっしょに れんしゅう",10,True,ORANGE)]], sp_after=0)

# ごほうび
rx=7.4; rw=9.4
rrect(rx,by,rw,bh, fill=RGBColor(0xFF,0xFD,0xF5), line=YELLOW, lw=1.6, radius=0.10)
text(rx+0.35,by+0.15,rw-0.6,0.7, [[("🎁 ごほうび",12,True,ORANGE)]])
rwd=[("🥉","シール 10こ"),("🥈","シール 30こ"),("🥇","ぜんぶ そろう")]
for i,(m,c) in enumerate(rwd):
    ry=by+0.95+i*0.62
    text(rx+0.35,ry-0.05,0.8,0.55,[[(m,12,False,NAVY)]])
    text(rx+1.15,ry-0.05,3.0,0.55,[[(c,10,True,ORANGE)]])
    text(rx+4.0,ry-0.05,0.6,0.55,[[("→",10,True,RGBColor(0xCB,0xB3,0x6A))]])
    dot_line(rx+4.7,ry+0.3,rw-5.1)
text(rx+0.35,by+2.62,rw-0.6,0.4, [[("＊ごほうびは 子どもと いっしょに きめると やる気アップ！",8.5,True,RGBColor(0xB0,0x9A,0x4E))]])

# メッセージ
mx=17.3; mw=11.7
rrect(mx,by,mw,bh, fill=WHITE, line=PINK_L, lw=1.6, radius=0.10)
text(mx+0.35,by+0.15,mw-0.6,0.7, [[("💬 今月の ふりかえり・がんばったね メッセージ",11.5,True,PINK)]])
dot_line(mx+0.4,by+1.35,mw-0.8, color=RGBColor(0xE3,0xB9,0xCF))
dot_line(mx+0.4,by+2.05,mw-0.8, color=RGBColor(0xE3,0xB9,0xCF))
text(mx+0.4,by+2.45,mw-0.8,0.5, [[("先生・おうちの人 ＿＿＿＿＿＿＿",10,True,GRAY)]], align=PP_ALIGN.RIGHT)

# ===== フッター =====
text(0.7,15.95,22.5,1.5,
     [[("使い方  ",9,True,SKY_D),("毎日の 終わりに 子どもと いっしょに ふり返り、できた 目標に シールを 1まい 貼る。注意の 代わりに「できた所を 見つけて すぐ ほめる」ことを 続けると、行動が 身に つきやすく なります。",9,False,GRAY)],
      [("目標①〜④は 個別支援計画の 課題から、⑤⑥は その子の 成長・チャレンジから 書きかえてOK。",9,False,GRAY)]],
     anchor=MSO_ANCHOR.TOP, sp_after=2)
text(23.5,15.95,5.5,1.2, [[("放課後等デイサービス 夢門塾",10,True,SKY_D)],[("M U M O N J U K U",8,True,GRAY)]],
     align=PP_ALIGN.RIGHT, sp_after=0)

out = os.path.join(HERE, "ganbari-card-mumonjuku.pptx")
prs.save(out)
print("saved:", out)
