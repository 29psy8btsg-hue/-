#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""わたしの がんばりカード（夢門塾）【1ヶ月仕様】を編集可能な PowerPoint(.pptx) で生成。
A4 横・1スライド＝1ヶ月。スキル別ゲージは20マス（≒1ヶ月の登園日数）。テキストは編集可能。"""
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "mumonjuku-logo.png")
FONT = "Meiryo"
SLOTS = 20  # 1ヶ月のマス数（登園日数に合わせて変更可）

SKY=RGBColor(0x4E,0xC5,0xE0); SKY_L=RGBColor(0x7D,0xD8,0xEC); SKY_D=RGBColor(0x2B,0xA8,0xC9)
NAVY=RGBColor(0x1A,0x3A,0x6B); PINK=RGBColor(0xE8,0x19,0x7D); PINK_L=RGBColor(0xF6,0x87,0xC0)
GREEN=RGBColor(0x22,0xC5,0x5E); GREEN_D=RGBColor(0x15,0x80,0x3d); ORANGE=RGBColor(0xFF,0x8C,0x42)
YELLOW=RGBColor(0xFF,0xC9,0x3C); PURPLE=RGBColor(0x8B,0x5C,0xF6); SOFT=RGBColor(0xF3,0xF9,0xFC)
LINE=RGBColor(0xC3,0xD0,0xDD); GRAY=RGBColor(0x94,0xA3,0xB8); WHITE=RGBColor(0xFF,0xFF,0xFF)
DASH=RGBColor(0xC7,0xD3,0xDF); INK_L=RGBColor(0x5b,0x71,0x87); CLOUD=RGBColor(0xBF,0xE7,0xF3)

prs = Presentation(); prs.slide_width=Cm(29.7); prs.slide_height=Cm(21.0)
slide = prs.slides.add_slide(prs.slide_layouts[6]); shapes = slide.shapes

def rrect(x,y,w,h, fill=None, line=None, lw=1.0, radius=0.10):
    sp=shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(x),Cm(y),Cm(w),Cm(h))
    try: sp.adjustments[0]=radius
    except Exception: pass
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb=fill
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb=line; sp.line.width=Pt(lw)
    sp.shadow.inherit=False; return sp

def oval(x,y,d, fill=None, line=None, lw=1.0):
    sp=shapes.add_shape(MSO_SHAPE.OVAL, Cm(x),Cm(y),Cm(d),Cm(d))
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb=fill
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb=line; sp.line.width=Pt(lw)
    sp.shadow.inherit=False; return sp

def star(x,y,d, fill, line=None, lw=1.0):
    sp=shapes.add_shape(MSO_SHAPE.STAR_5_POINT, Cm(x),Cm(y),Cm(d),Cm(d))
    sp.fill.solid(); sp.fill.fore_color.rgb=fill
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb=line; sp.line.width=Pt(lw)
    sp.shadow.inherit=False; return sp

def badge(cx,cy,d, num, fill, numcol, fs=12):
    """色付き丸＋番号（絵文字の代わり）"""
    oval(cx,cy,d, fill=fill)
    text(cx,cy-0.03,d,d, [[(str(num),fs,True,numcol)]], align=PP_ALIGN.CENTER)

def text(x,y,w,h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, wrap=True, sp_after=2):
    tb=shapes.add_textbox(Cm(x),Cm(y),Cm(w),Cm(h)); tf=tb.text_frame
    tf.word_wrap=wrap; tf.vertical_anchor=anchor
    tf.margin_left=Cm(0.1); tf.margin_right=Cm(0.1); tf.margin_top=Cm(0.02); tf.margin_bottom=Cm(0.02)
    for i,line in enumerate(runs):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align; p.space_after=Pt(sp_after); p.space_before=Pt(0)
        for (t,sz,b,col) in line:
            r=p.add_run(); r.text=t; f=r.font; f.size=Pt(sz); f.bold=b; f.color.rgb=col; f.name=FONT
    return tb

def dot_line(x,y,w, color=LINE):
    ln=shapes.add_connector(2, Cm(x),Cm(y),Cm(x+w),Cm(y))
    ln.line.color.rgb=color; ln.line.width=Pt(1.2)
    d=ln.line._get_or_add_ln(); d.append(d.makeelement(qn('a:prstDash'),{'val':'sysDash'}))
    ln.shadow.inherit=False; return ln

# ===== 背景 =====
rrect(0.3,0.3,29.1,20.4, fill=WHITE, line=SKY, lw=2.2, radius=0.04)

# ===== 夢門塾 ヘッダー =====
rrect(0.3,0.3,29.1,3.3, fill=SKY, line=None, radius=0.04)
for (cx,cy,cd) in [(7.0,-1.4,3.0),(13.0,2.5,2.2),(22.5,-1.3,2.6)]:
    oval(cx,cy,cd, fill=CLOUD)
rrect(0.9,0.65,3.0,2.6, fill=WHITE, line=None, radius=0.16)
if os.path.exists(LOGO):
    shapes.add_picture(LOGO, Cm(1.1),Cm(0.85), Cm(2.6),Cm(2.6))
text(4.4,0.65,15.8,1.6, [[("わたしの がんばりカード",27,True,WHITE)]], anchor=MSO_ANCHOR.BOTTOM)
rrect(4.5,2.35,9.6,0.8, fill=SKY_L, line=None, radius=0.5)
text(4.5,2.35,9.6,0.8, [[("1ヶ月の せいちょうきろく ・ まいにち ふりかえり",10.5,True,NAVY)]], align=PP_ALIGN.CENTER)
# 名前・月
rrect(20.5,0.6,8.5,2.45, fill=WHITE, line=None, radius=0.12)
text(20.9,0.8,2.6,0.85, [[("なまえ",13,True,SKY_D)]]); dot_line(23.2,1.6,5.5)
text(20.9,1.78,2.6,0.85, [[("　月",13,True,SKY_D)]])
text(23.0,1.78,1.5,0.85, [[("（",13,True,SKY_D)]], align=PP_ALIGN.LEFT)
dot_line(23.6,2.58,1.6); text(25.3,1.78,3.6,0.85, [[("）月",13,True,SKY_D)]])

# ===== 個別支援計画 =====
plan=[("ながい もくひょう","（個別支援計画より）",SKY),
      ("今月の チャレンジ","（自分で きめる 近い 目標）",ORANGE),
      ("おうえん","（おうちの人・先生から）",PINK)]
px=0.7; pw=9.1; py=3.85; ph=1.75
for i,(t,sub,col) in enumerate(plan):
    x=px+i*(pw+0.5)
    rrect(x,py,pw,ph, fill=SOFT, line=col, lw=1.4, radius=0.12)
    rrect(x,py+0.22,0.16,ph-0.44, fill=col, line=None, radius=0.5)
    text(x+0.35,py+0.14,pw-0.5,0.6, [[(t,11,True,col),("  "+sub,8,False,GRAY)]])
    dot_line(x+0.4,py+1.32,pw-0.8)

# ===== 目標ゲージ（1ヶ月・横いっぱい・6スキル） =====
text(0.7,5.78,9.5,0.7, [[("できた日に シールを はろう（1ヶ月）",15,True,PINK)]])
rrect(10.7,5.88,7.6,0.62, fill=RGBColor(0xE6,0xF6,0xFB), line=None, radius=0.5)
text(10.7,5.88,7.6,0.62, [[("シールが ふえると せいちょうが 見える！",10,True,SKY_D)]], align=PP_ALIGN.CENTER)

gauges=[("おだやかな こえ","落ち着いた 声で 話す",SKY),
        ("やさしい 言葉","ふわふわ 言葉を 使う",PINK_L),
        ("手・足は おだやかに","たたかない・けらない",PURPLE),
        ("すわって とりくむ","自分の 場所で おちついて",GREEN),
        ("きらり ポイント","とくい・成長した こと",YELLOW),
        ("今月の チャレンジ","自分で きめた こと",ORANGE)]
gx=0.7; gw=28.3; gy0=6.6; gh=1.28; gvs=1.5
for i,(n1,n2,col) in enumerate(gauges):
    y=gy0+i*gvs
    rrect(gx,y,gw,gh, fill=WHITE, line=col, lw=1.5, radius=0.12)
    # 番号バッジ（絵文字なし）
    bd=0.92; numcol = NAVY if col is YELLOW else WHITE
    badge(gx+0.28,y+(gh-bd)/2,bd, i+1, col, numcol)
    text(gx+1.35,y+0.12,5.3,1.05, [[(n1,12.5,True,NAVY)],[(n2,8,False,GRAY)]], sp_after=0)
    n=SLOTS; start=gx+6.9; right=gx+gw-0.3
    cd=0.92; gap=(right-start-n*cd)/(n-1)
    cy=y+(gh-cd)/2
    for k in range(n):
        cx=start+k*(cd+gap)
        if k in (4,9,14):
            oval(cx,cy,cd, fill=RGBColor(0xFF,0xF6,0xE0), line=YELLOW, lw=1.3)
            star(cx+cd*0.22,cy+cd*0.24,cd*0.56, fill=YELLOW)  # 5/10/15 の節目
        elif k==n-1:
            oval(cx,cy,cd, fill=RGBColor(0xFC,0xE7,0xF0), line=PINK, lw=1.4)
            star(cx+cd*0.18,cy+cd*0.20,cd*0.64, fill=PINK)    # ゴール(20)
        else:
            oval(cx,cy,cd, fill=WHITE, line=DASH, lw=1.2)
            text(cx,cy-0.02,cd,cd, [[(str(k+1),7.5,True,GRAY)]], align=PP_ALIGN.CENTER)

# ===== 下段 =====
by=15.7; bh=3.0
# 見かた
lx=0.7; lw_=6.4
rrect(lx,by,lw_,bh, fill=SOFT, line=RGBColor(0xD6,0xE3,0xEC), lw=1.6, radius=0.10)
text(lx+0.35,by+0.12,lw_-0.6,0.6, [[("シールの 見かた",12,True,INK_L)]])
for k in range(5): oval(lx+0.4+k*0.62, by+1.0, 0.5, fill=YELLOW)
text(lx+3.7,by+0.9,2.6,0.7, [[("いっぱい ＝ とくい！",10,True,GREEN_D)]])
oval(lx+0.4, by+1.9, 0.5, fill=YELLOW)
for k in range(1,5): oval(lx+0.4+k*0.62, by+1.9, 0.5, fill=WHITE, line=DASH, lw=1.3)
text(lx+3.7,by+1.78,2.6,0.95, [[("すくない ＝",10,True,ORANGE)],[("いっしょに れんしゅう",10,True,ORANGE)]], sp_after=0)

# ごほうび（1ヶ月）
rx=7.4; rw=9.4
rrect(rx,by,rw,bh, fill=RGBColor(0xFF,0xFD,0xF5), line=YELLOW, lw=1.6, radius=0.10)
text(rx+0.35,by+0.12,rw-3.0,0.6, [[("今月の ごほうび",12,True,ORANGE)]])
text(rx+5.0,by+0.16,rw-5.2,0.55, [[("シール合計（　）こ",9,True,SKY_D)]], align=PP_ALIGN.RIGHT)
BRONZE=RGBColor(0xCD,0x7F,0x32); SILVER=RGBColor(0xA8,0xB0,0xB8); GOLD=RGBColor(0xD4,0xA0,0x17)
rwd=[(BRONZE,"シール 20こ"),(SILVER,"ゲージ 3本 クリア"),(GOLD,"ぜんぶ クリア")]
for i,(mc,c) in enumerate(rwd):
    ry=by+0.92+i*0.6
    badge(rx+0.35,ry-0.02,0.5, i+1, mc, WHITE, fs=9)
    text(rx+1.05,ry-0.05,3.4,0.5,[[(c,10,True,ORANGE)]])
    text(rx+4.5,ry-0.05,0.5,0.5,[[("→",10,True,RGBColor(0xCB,0xB3,0x6A))]])
    dot_line(rx+5.1,ry+0.28,rw-5.5)
text(rx+0.35,by+2.6,rw-0.6,0.4, [[("＊ごほうびは 子どもと いっしょに きめると やる気アップ！",8.5,True,RGBColor(0xB0,0x9A,0x4E))]])

# 今月のふりかえり
mx=17.3; mw=11.7
rrect(mx,by,mw,bh, fill=WHITE, line=PINK_L, lw=1.6, radius=0.10)
text(mx+0.35,by+0.12,mw-0.6,0.6, [[("今月の ふりかえり・がんばったね メッセージ",11.5,True,PINK)]])
dot_line(mx+0.4,by+1.3,mw-0.8, color=RGBColor(0xE3,0xB9,0xCF))
dot_line(mx+0.4,by+2.0,mw-0.8, color=RGBColor(0xE3,0xB9,0xCF))
text(mx+0.4,by+2.42,mw-0.8,0.5, [[("先生・おうちの人 ＿＿＿＿＿＿＿",10,True,GRAY)]], align=PP_ALIGN.RIGHT)

# ===== フッター =====
text(0.7,18.95,22.3,1.4,
     [[("使い方  ",9,True,SKY_D),("1ヶ月を 1まいで 記録。毎日の 終わりに ふり返り、できた 目標に シールを 1まい 貼る（注意の 代わりに「できた所を 見つけて すぐ ほめる」）。",9,False,GRAY)],
      [("月末に ゲージを 見て、得意・苦手を つぎの 月の 目標に いかす。目標①〜④は 個別支援計画の 課題から、⑤⑥は 成長・チャレンジから 書きかえてOK。",9,False,GRAY)]],
     anchor=MSO_ANCHOR.TOP, sp_after=2)
text(23.4,18.95,5.6,1.1, [[("放課後等デイサービス 夢門塾",10,True,SKY_D)],[("M U M O N J U K U",8,True,GRAY)]],
     align=PP_ALIGN.RIGHT, sp_after=0)

out=os.path.join(HERE,"ganbari-card-mumonjuku.pptx"); prs.save(out); print("saved:", out)
