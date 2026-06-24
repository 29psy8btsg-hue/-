#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
夢通信 行事予定表カレンダー生成スクリプト v2
放課後等デイサービス「夢門塾ゆうゆう日吉」用
― 令和8年7月号PDFのレイアウトを忠実に再現 ―

使い方:
    python3 yume_calendar_v2.py            # 既定（9月号）を生成
    MONTH_DATA を編集して他の月にも流用できます。
"""
import calendar as cal_mod
from docx import Document
from docx.shared import Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# 配色（PDF基準）
# ============================================================
RED   = RGBColor(0xE6, 0x00, 0x00)
BLUE  = RGBColor(0x15, 0x65, 0xC0)
DATEB = RGBColor(0x00, 0x66, 0xCC)   # おしらせの日付（青）
MAG   = RGBColor(0xEC, 0x00, 0x8C)   # 夢通信ロゴ
CYAN  = RGBColor(0x00, 0x99, 0xCC)   # 夢門塾日吉
ORANGE= RGBColor(0xE8, 0x6A, 0x00)   # ねらい本文・おしらせ見出し
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x22, 0x22, 0x22)
GREEN = RGBColor(0x4C, 0x9A, 0x2A)
GRAY  = RGBColor(0xBB, 0xBB, 0xBB)

HEAD_FILL = "FBE2EC"   # 曜日ヘッダー（薄ピンク）
SUN_FILL  = "FDEFEF"   # 日曜セル
SAT_FILL  = "EAF3FF"   # 土曜セル
HOL_FILL  = "FFF3D6"   # 祝日セル
NEWS_FILL = "FFFCF5"
AIM_FILL  = "F4FBEC"
BAR_FILL  = {"orange":"F39B3A", "magenta":"E060A8", "green":"8DC54A", "blue":"5BAEE0"}

# ============================================================
# 祝日（2025〜2028）
# ============================================================
HOLIDAYS = {
    "2026-01-01":"元日","2026-01-12":"成人の日","2026-02-11":"建国記念の日",
    "2026-02-23":"天皇誕生日","2026-03-20":"春分の日","2026-04-29":"昭和の日",
    "2026-05-03":"憲法記念日","2026-05-04":"みどりの日","2026-05-05":"こどもの日",
    "2026-05-06":"振替休日","2026-07-20":"海の日","2026-08-11":"山の日",
    "2026-09-21":"敬老の日","2026-09-22":"国民の休日","2026-09-23":"秋分の日",
    "2026-10-12":"スポーツの日","2026-11-03":"文化の日","2026-11-23":"勤労感謝の日",
}
KANJI = ["〇","一","二","三","四","五","六","七","八","九","十","十一","十二",
         "十三","十四","十五","十六","十七","十八","十九","二十"]

# ============================================================
# 月別データ（ここを編集すれば各月を作成できます）
# ============================================================
MONTH_DATA = {
 (2026, 9): {
    "facility": "夢門塾ゆうゆう日吉",
    "tel": "0898-52-8741",
    "mobile": "090-4590-0983",
    "greeting": ("朝晩は少しずつ過ごしやすくなり、秋の気配が感じられる季節になりました。"
                 "実りの秋・スポーツの秋、戸外あそびや制作を通して、"
                 "元気いっぱい楽しい思い出を沢山作りましょう！"),
    # 複数日にまたがる行事（バー表示）  week=週(0始まり) cols=曜日列(0=日)
    "bars": [
        {"week":0, "cols":[3,4,5], "name":"敬老の日カード作り", "yomi":"けいろうのひかーどづくり", "emoji":"💌", "color":"orange"},
        {"week":1, "cols":[3,4,5], "name":"運動会ごっこ",       "yomi":"うんどうかいごっこ",       "emoji":"🏃", "color":"magenta"},
        {"week":2, "cols":[3,4,5], "name":"秋の製作",           "yomi":"あきのせいさく",           "emoji":"🍁", "color":"green"},
    ],
    # 単日行事  day -> {...}
    "events": {
        1:  {"name":"防災訓練",       "yomi":"ぼうさいくんれん", "emoji":"⛑️"},
        5:  {"name":"公園",           "yomi":"こうえん",         "emoji":"🏞️", "badge":"おさんぽ"},
        7:  {"name":"リズム体操",     "yomi":"りずむたいそう",   "emoji":"🎵"},
        8:  {"name":"おはぎ作り",     "yomi":"おはぎづくり",     "emoji":"🍡"},
        12: {"name":"体育館",         "yomi":"たいいくかん",     "emoji":"🤸", "badge":"おさんぽ"},
        14: {"name":"工作",           "yomi":"こうさく",         "emoji":"✂️"},
        15: {"name":"ゆうゆうシネマ", "yomi":"ゆうゆうしねま",   "emoji":"🎬"},
        19: {"name":"芋掘り",         "yomi":"いもほり",         "emoji":"🍠", "badge":"おさんぽ", "badge2":"人気No.1"},
        21: {"name":"お楽しみ会",     "yomi":"おたのしみかい",   "emoji":"🎁"},
        22: {"name":"外食",           "yomi":"がいしょく",       "emoji":"🍔", "sub":"（マクドナルド）"},
        23: {"name":"お月見団子作り", "yomi":"おつきみだんごづくり","emoji":"🌕"},
        24: {"name":"駄菓子屋",       "yomi":"だがしや",         "emoji":"🍬"},
        25: {"name":"カラオケ",       "yomi":"からおけ",         "emoji":"🎤"},
        26: {"name":"アスレチック公園","yomi":"あすれちっくこうえん","emoji":"🧗", "badge":"おさんぽ"},
        28: {"name":"風船バレー",     "yomi":"ふうせんばれー",   "emoji":"🎈"},
        29: {"name":"ボウリング",     "yomi":"ぼうりんぐ",       "emoji":"🎳"},
        30: {"name":"クイズ大会",     "yomi":"くいずたいかい",   "emoji":"❓"},
    },
    "notices_left": [
        ("9/5(土)「公園」", "帽子、動きやすい服装・靴、お弁当、水筒、タオル、着替え", ["帽子","動きやすい服装・靴"]),
        ("9/8(火)「おはぎ作り」", "食材費200円、エプロン、マスク、三角巾、水筒、タオル、着替え", ["食材費200円","エプロン","マスク","三角巾"]),
        ("9/12(土)「体育館」", "室内シューズ、お弁当、水筒、タオル、着替え", ["室内シューズ"]),
        ("9/19(土)「芋掘り」", "軍手、汚れてもよい服装、お弁当、水筒、タオル、着替え", ["軍手","汚れてもよい服装"]),
        ("9/22(火)「マクドナルド」", "昼食代1000円まで、水筒、タオル、着替え", ["昼食代1000円まで"]),
        ("9/23(水)「お月見団子作り」", "食材費200円、エプロン、マスク、三角巾、水筒、タオル", ["食材費200円","エプロン","マスク","三角巾"]),
        ("9/24(木)「駄菓子屋」", "おやつ代100円、お弁当、水筒、タオル、着替え", ["おやつ代100円"]),
        ("9/25(金)「カラオケ」", "カラオケ代660円、お弁当、水筒、タオル、着替え", ["カラオケ代660円"]),
    ],
    "notices_right": [
        ("9/26(土)「アスレチック公園」", "帽子、動きやすい服装・靴、お弁当、水筒、タオル、着替え", ["帽子","動きやすい服装・靴"]),
        ("9/1(火)「防災訓練」", "お弁当、水筒、タオル、着替え", []),
        ("9/21(月)「敬老の日お楽しみ会」", "お弁当、水筒、タオル、着替え", []),
    ],
    "notes": [
        "※外出行事の際は、必ず動きやすい服装、運動靴でお越しください。",
        "※雨天時は中止により室内行事に切り替わることがございます。",
        "※お昼ご飯はお買い物に行くことも可能です。昼食代をご持参ください。",
        "※敬老の日・国民の休日・秋分の日も終日（10:00〜）開所いたします。",
    ],
    "aims": [
        ("敬老の日カード作り", "制作活動を通して、手先の操作性や集中力を養いながら、おじいちゃん・おばあちゃんへ感謝の気持ちを伝える経験につなげていきます。"),
        ("運動会ごっこ", "ルールを守り、友だちと協力して活動する力を育みます。また、身体を動かす楽しさを感じながら、判断力やコミュニケーション力につなげていきます。"),
        ("秋の製作", "季節の自然物に触れながら、想像力や表現力を育みます。また、繰り返し取り組むことで集中力や達成感につなげていきます。"),
    ],
 },
}

# ============================================================
# XMLヘルパー
# ============================================================
def _goc(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag); parent.append(el)
    return el

def cell_fill(cell, hex6):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),hex6)
    _goc(cell._tc,'w:tcPr').append(shd)

def cell_borders(cell, color="9AC36A", sz=6):
    tcB = OxmlElement('w:tcBorders')
    for s in ('top','left','bottom','right'):
        b=OxmlElement(f'w:{s}'); b.set(qn('w:val'),'single'); b.set(qn('w:sz'),str(sz))
        b.set(qn('w:space'),'0'); b.set(qn('w:color'),color); tcB.append(b)
    _goc(cell._tc,'w:tcPr').append(tcB)

def no_table_borders(table):
    tblB=OxmlElement('w:tblBorders')
    for s in ('top','left','bottom','right','insideH','insideV'):
        el=OxmlElement(f'w:{s}'); el.set(qn('w:val'),'none'); tblB.append(el)
    _goc(table._tbl,'w:tblPr').append(tblB)

def fixed_layout(table):
    lay=OxmlElement('w:tblLayout'); lay.set(qn('w:type'),'fixed')
    _goc(table._tbl,'w:tblPr').append(lay)

def row_h(row, mm, rule='atLeast'):
    trPr=_goc(row._tr,'w:trPr'); h=OxmlElement('w:trHeight')
    h.set(qn('w:val'),str(int(mm*56.693))); h.set(qn('w:hRule'),rule); trPr.append(h)

def col_widths(table, widths_mm):
    tblPr=_goc(table._tbl,'w:tblPr')
    w=OxmlElement('w:tblW'); w.set(qn('w:w'),str(int(sum(widths_mm)/25.4*1440))); w.set(qn('w:type'),'dxa')
    tblPr.append(w)
    grid=OxmlElement('w:tblGrid')
    for mm in widths_mm:
        gc=OxmlElement('w:gridCol'); gc.set(qn('w:w'),str(int(mm/25.4*1440))); grid.append(gc)
    tbl=table._tbl; tbl.insert(list(tbl).index(tblPr)+1, grid)
    for r in table.rows:
        for i,c in enumerate(r.cells):
            tcW=OxmlElement('w:tcW'); tcW.set(qn('w:w'),str(int(widths_mm[i]/25.4*1440))); tcW.set(qn('w:type'),'dxa')
            _goc(c._tc,'w:tcPr').append(tcW)

def valign(cell,val='top'):
    va=OxmlElement('w:vAlign'); va.set(qn('w:val'),val); _goc(cell._tc,'w:tcPr').append(va)

def spc(p,before=0,after=0,line=200):
    pPr=_goc(p._p,'w:pPr'); s=OxmlElement('w:spacing')
    s.set(qn('w:before'),str(before)); s.set(qn('w:after'),str(after))
    s.set(qn('w:line'),str(line)); s.set(qn('w:lineRule'),'auto'); pPr.append(s)

def para_fill(p,hex6):
    pPr=_goc(p._p,'w:pPr'); shd=OxmlElement('w:shd')
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),hex6); pPr.append(shd)

def run_shade(r,hex6):
    rpr=r._element.get_or_add_rPr(); shd=OxmlElement('w:shd')
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),hex6); rpr.append(shd)

def run(p,text,size,bold=False,color=None,font="BIZ UDPゴシック"):
    r=p.add_run(text); r.font.size=Pt(size); r.font.bold=bold
    if color is not None: r.font.color.rgb=color
    r.font.name=font
    rpr=r._element.get_or_add_rPr(); rf=rpr.find(qn('w:rFonts'))
    if rf is None: rf=OxmlElement('w:rFonts'); rpr.append(rf)
    rf.set(qn('w:eastAsia'),font)
    return r

def para(cell, align=WD_ALIGN_PARAGRAPH.LEFT, first=False, before=0, after=0, line=200):
    if first:
        p=cell.paragraphs[0]
        for r in list(p.runs): r._element.getparent().remove(r._element)
    else:
        p=cell.add_paragraph()
    p.alignment=align; spc(p,before,after,line); return p

# ============================================================
# 生成
# ============================================================
def generate(year, month):
    d = MONTH_DATA[(year, month)]
    reiwa = year - 2018
    rstr = KANJI[reiwa] if 0 < reiwa < len(KANJI) else str(reiwa)
    mstr = KANJI[month]

    doc = Document()
    st = doc.styles['Normal']; st.font.name='BIZ UDPゴシック'; st.font.size=Pt(10)
    st.element.rPr.rFonts.set(qn('w:eastAsia'),'BIZ UDPゴシック')
    st.paragraph_format.space_before=Pt(0); st.paragraph_format.space_after=Pt(0)

    sec=doc.sections[0]
    sec.orientation=WD_ORIENT.PORTRAIT
    sec.page_width=Mm(210); sec.page_height=Mm(297)
    sec.left_margin=Mm(8); sec.right_margin=Mm(8)
    sec.top_margin=Mm(7); sec.bottom_margin=Mm(6)

    # ---------- ① ヘッダー ----------
    ht=doc.add_table(rows=1,cols=2); ht.alignment=WD_TABLE_ALIGNMENT.CENTER
    no_table_borders(ht); fixed_layout(ht); col_widths(ht,[110,84])
    row_h(ht.rows[0],36,'atLeast')

    lc=ht.rows[0].cells[0]; valign(lc,'center')
    p=para(lc,WD_ALIGN_PARAGRAPH.CENTER,first=True)
    run(p,"夢門塾日吉",22,bold=True,color=CYAN,font="HGP創英角ﾎﾟｯﾌﾟ体")
    p=para(lc,WD_ALIGN_PARAGRAPH.CENTER)
    run(p,"夢通信",52,bold=True,color=MAG,font="HGP創英角ﾎﾟｯﾌﾟ体")
    p=para(lc,WD_ALIGN_PARAGRAPH.CENTER)
    run(p,"放課後等デイサービス",11,bold=True,color=RED,font="HGP創英角ﾎﾟｯﾌﾟ体")

    rc=ht.rows[0].cells[1]; valign(rc,'center')
    cell_fill(rc,"EAF5E1"); cell_borders(rc,"9CCB7A",12)
    p=para(rc,WD_ALIGN_PARAGRAPH.CENTER,first=True,before=40)
    run(p,f"令和{rstr}年　{mstr}月号",20,bold=True,color=BLACK)
    p=para(rc,WD_ALIGN_PARAGRAPH.CENTER,before=60)
    run(p,d["facility"],13,bold=True,color=BLACK)
    p=para(rc,WD_ALIGN_PARAGRAPH.CENTER,before=20)
    run(p,f"☎ {d['tel']}　📠",12,color=BLACK)
    p=para(rc,WD_ALIGN_PARAGRAPH.CENTER)
    run(p,f"携帯 {d['mobile']}　🌙⭐",12,color=BLACK)

    # ---------- ② あいさつ ----------
    p=doc.add_paragraph(); spc(p,40,40,200)
    run(p,d["greeting"],10.5,bold=True,color=BLACK)

    # ---------- ③ カレンダー ----------
    cal_mod.setfirstweekday(cal_mod.SUNDAY)
    weeks=cal_mod.monthcalendar(year,month)
    nw=len(weeks)
    ct=doc.add_table(rows=1+nw,cols=7); ct.alignment=WD_TABLE_ALIGNMENT.CENTER
    no_table_borders(ct); fixed_layout(ct)
    col_widths(ct,[27.7]*7)

    days=["日曜日","月曜日","火曜日","水曜日","木曜日","金曜日","土曜日"]
    hr=ct.rows[0]; row_h(hr,7,'exact')
    for i,(c,lab) in enumerate(zip(hr.cells,days)):
        cell_fill(c,HEAD_FILL); cell_borders(c,"E59ABB",6); valign(c,'center')
        p=para(c,WD_ALIGN_PARAGRAPH.CENTER,first=True)
        col=RED if i==0 else (BLUE if i==6 else BLACK)
        run(p,lab,11,bold=True,color=col)

    rh=26 if nw<=5 else 22
    bars_by_week={}
    for b in d["bars"]:
        bars_by_week.setdefault(b["week"],[]).append(b)

    for wi,week in enumerate(weeks):
        rr=ct.rows[wi+1]; row_h(rr,rh,'atLeast')
        # この週のバー対象列
        barmap={}  # col -> (bar, position)
        for b in bars_by_week.get(wi,[]):
            cols=b["cols"]
            for idx,col in enumerate(cols):
                pos='mid' if idx==len(cols)//2 else 'side'
                barmap[col]=(b,pos)
        for ci,day in enumerate(week):
            cell=rr.cells[ci]; cell_borders(cell,"CDB07A",6); valign(cell,'top')
            if day==0:
                cell_fill(cell,"F4F4F4")
                para(cell,first=True); continue

            iso=f"{year}-{month:02d}-{day:02d}"
            hol=HOLIDAYS.get(iso)
            is_sun=(ci==0); is_sat=(ci==6)

            if hol: cell_fill(cell,HOL_FILL)
            elif is_sun: cell_fill(cell,SUN_FILL)
            elif is_sat: cell_fill(cell,SAT_FILL)

            ev=d["events"].get(day)
            badge = ev.get("badge") if ev else None
            badge2= ev.get("badge2") if ev else None

            # バッジ行（おさんぽ / 人気No.1）＋ 日付
            if badge:
                pb=para(cell,WD_ALIGN_PARAGRAPH.LEFT,first=True,line=170)
                rb=run(pb," "+badge+" ",7,bold=True,color=WHITE)
                run_shade(rb,"E8553A")
                if badge2:
                    run(pb," ",6)
                    rb2=run(pb," ★"+badge2+" ",7,bold=True,color=WHITE)
                    run_shade(rb2,"F08C00")
                pnum=para(cell,WD_ALIGN_PARAGRAPH.RIGHT,line=170)
            else:
                pnum=para(cell,WD_ALIGN_PARAGRAPH.RIGHT,first=True,line=180)
            dcol = RED if (is_sun or hol) else (BLUE if is_sat else BLACK)
            run(pnum,str(day),13,bold=True,color=dcol)

            # 祝日名
            if hol:
                ph=para(cell,WD_ALIGN_PARAGRAPH.CENTER,line=160)
                run(ph,hol,7.5,bold=True,color=RED)

            # 日曜＝休
            if is_sun and not ev and not hol:
                pk=para(cell,WD_ALIGN_PARAGRAPH.CENTER,before=40)
                run(pk,"休",22,bold=True,color=GRAY)
                continue

            # バー（複数日行事）
            if ci in barmap:
                b,pos=barmap[ci]
                fill=BAR_FILL[b["color"]]
                # よみ
                pf=para(cell,WD_ALIGN_PARAGRAPH.CENTER,before=20,line=140)
                run(pf,(b["yomi"] if pos=='mid' else "　"),6.5,color=BLACK)
                # バー本体（段落背景）
                pbar=para(cell,WD_ALIGN_PARAGRAPH.CENTER,line=220)
                para_fill(pbar,fill)
                if pos=='mid':
                    run(pbar,f'{b["emoji"]} {b["name"]} {b["emoji"]}',10,bold=True,color=WHITE)
                else:
                    run(pbar,"➡" if ci>min(b["cols"]) else "　",10,bold=True,color=WHITE)
                continue

            # 単日行事
            if ev:
                pf=para(cell,WD_ALIGN_PARAGRAPH.CENTER,before=30,line=140)
                run(pf,ev["yomi"],6.5,color=BLACK)
                pn=para(cell,WD_ALIGN_PARAGRAPH.CENTER,line=170)
                run(pn,ev["name"],10.5,bold=True,color=GREEN)
                if ev.get("sub"):
                    ps=para(cell,WD_ALIGN_PARAGRAPH.CENTER,line=140)
                    run(ps,ev["sub"],7,color=BLACK)
                pe=para(cell,WD_ALIGN_PARAGRAPH.CENTER,before=10)
                run(pe,ev["emoji"],18)

    # ---------- ④ 下段：おしらせ ＋ ねらい ----------
    sp=doc.add_paragraph(); spc(sp,30,30)

    # おしらせ見出し
    pt=doc.add_paragraph(); pt.alignment=WD_ALIGN_PARAGRAPH.CENTER; spc(pt,20,20)
    para_fill(pt,"FAD7E6")
    run(pt,"お　し　ら　せ",18,bold=True,color=ORANGE,font="HGP創英角ﾎﾟｯﾌﾟ体")

    bt=doc.add_table(rows=1,cols=2); bt.alignment=WD_TABLE_ALIGNMENT.CENTER
    no_table_borders(bt); fixed_layout(bt); col_widths(bt,[97,97])
    row_h(bt.rows[0],86,'atLeast')

    def render_notice_item(cell, date, items, bolds, first=False):
        p=para(cell,WD_ALIGN_PARAGRAPH.LEFT,first=first,before=10,line=170)
        run(p,"●"+date,9,bold=True,color=DATEB)
        p2=para(cell,WD_ALIGN_PARAGRAPH.LEFT,line=170)
        run(p2,"【持参物】",8.5,bold=True,color=BLACK)
        # 太字キーワードを赤、その他は黒
        rest=items
        for kw in bolds:
            i=rest.find(kw)
            if i<0:
                continue
            if rest[:i]: run(p2,rest[:i],8.5,color=BLACK)
            run(p2,kw,8.5,bold=True,color=RED)
            rest=rest[i+len(kw):]
        if rest: run(p2,rest,8.5,color=BLACK)

    lc=bt.rows[0].cells[0]; valign(lc,'top')
    cell_fill(lc,NEWS_FILL); cell_borders(lc,"E0E0E0",6)
    first=True
    for date,items,bolds in d["notices_left"]:
        render_notice_item(lc,date,items,bolds,first=first); first=False

    rc=bt.rows[0].cells[1]; valign(rc,'top')
    cell_fill(rc,NEWS_FILL); cell_borders(rc,"E0E0E0",6)
    first=True
    for date,items,bolds in d["notices_right"]:
        render_notice_item(rc,date,items,bolds,first=first); first=False
    for note in d["notes"]:
        pn=para(rc,WD_ALIGN_PARAGRAPH.LEFT,before=10,line=170)
        run(pn,note,8.5,color=BLACK)

    # ---------- ⑤ プログラムのねらい ----------
    sp=doc.add_paragraph(); spc(sp,30,20)
    pt=doc.add_paragraph(); pt.alignment=WD_ALIGN_PARAGRAPH.CENTER; spc(pt,10,10)
    para_fill(pt,"CDEBA6")
    run(pt,"◆ プログラムのねらい ◆",15,bold=True,color=RGBColor(0x3A,0x6B,0x12),font="HGP創英角ﾎﾟｯﾌﾟ体")

    at=doc.add_table(rows=1,cols=1); at.alignment=WD_TABLE_ALIGNMENT.CENTER
    no_table_borders(at); fixed_layout(at); col_widths(at,[194])
    ac=at.rows[0].cells[0]; valign(ac,'top'); cell_fill(ac,AIM_FILL); cell_borders(ac,"9CCB7A",8)
    first=True
    for title,body in d["aims"]:
        ph=para(ac,WD_ALIGN_PARAGRAPH.LEFT,first=first,before=20,line=190); first=False
        run(ph,f"【{title}】",10.5,bold=True,color=ORANGE)
        pb=para(ac,WD_ALIGN_PARAGRAPH.LEFT,line=190)
        run(pb,body,9.5,bold=True,color=ORANGE)

    fn=f"夢通信_{year}年{month:02d}月号.docx"
    doc.save(fn)
    print("✅ 生成:", fn)
    return fn

if __name__ == "__main__":
    import sys
    if len(sys.argv)>=3:
        generate(int(sys.argv[1]), int(sys.argv[2]))
    else:
        generate(2026, 9)
