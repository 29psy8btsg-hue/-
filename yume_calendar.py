#!/usr/bin/env python3
"""
夢通信 行事予定表カレンダー生成スクリプト
放課後等デイサービス「夢門塾」用

Usage: python3 yume_calendar.py [YEAR] [MONTH]
例:   python3 yume_calendar.py 2026 7
"""
import sys, datetime, calendar as cal_mod

from docx import Document
from docx.shared import Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# 施設情報（ここを書き換えてください）
# ============================================================
FACILITY = "夢門塾〇〇"
TEL      = "012-345-6789"
FAX      = "012-345-6798"

# ============================================================
# 日本の祝日 2024〜2028
# ============================================================
HOLIDAYS = {
    "2024-01-01":"元日","2024-01-08":"成人の日","2024-02-11":"建国記念の日",
    "2024-02-12":"振替休日","2024-02-23":"天皇誕生日","2024-03-20":"春分の日",
    "2024-04-29":"昭和の日","2024-05-03":"憲法記念日","2024-05-04":"みどりの日",
    "2024-05-05":"こどもの日","2024-05-06":"振替休日","2024-07-15":"海の日",
    "2024-08-11":"山の日","2024-08-12":"振替休日","2024-09-16":"敬老の日",
    "2024-09-22":"秋分の日","2024-09-23":"振替休日","2024-10-14":"スポーツの日",
    "2024-11-03":"文化の日","2024-11-04":"振替休日","2024-11-23":"勤労感謝の日",
    "2025-01-01":"元日","2025-01-13":"成人の日","2025-02-11":"建国記念の日",
    "2025-02-23":"天皇誕生日","2025-02-24":"振替休日","2025-03-20":"春分の日",
    "2025-04-29":"昭和の日","2025-05-03":"憲法記念日","2025-05-04":"みどりの日",
    "2025-05-05":"こどもの日","2025-05-06":"振替休日","2025-07-21":"海の日",
    "2025-08-11":"山の日","2025-09-15":"敬老の日","2025-09-23":"秋分の日",
    "2025-10-13":"スポーツの日","2025-11-03":"文化の日","2025-11-23":"勤労感謝の日",
    "2025-11-24":"振替休日",
    "2026-01-01":"元日","2026-01-12":"成人の日","2026-02-11":"建国記念の日",
    "2026-02-23":"天皇誕生日","2026-03-20":"春分の日","2026-04-29":"昭和の日",
    "2026-05-03":"憲法記念日","2026-05-04":"みどりの日","2026-05-05":"こどもの日",
    "2026-05-06":"振替休日","2026-07-20":"海の日","2026-08-11":"山の日",
    "2026-09-21":"敬老の日","2026-09-22":"国民の休日","2026-09-23":"秋分の日",
    "2026-10-12":"スポーツの日","2026-11-03":"文化の日","2026-11-23":"勤労感謝の日",
    "2027-01-01":"元日","2027-01-11":"成人の日","2027-02-11":"建国記念の日",
    "2027-02-23":"天皇誕生日","2027-03-21":"春分の日","2027-03-22":"振替休日",
    "2027-04-29":"昭和の日","2027-05-03":"憲法記念日","2027-05-04":"みどりの日",
    "2027-05-05":"こどもの日","2027-07-19":"海の日","2027-08-11":"山の日",
    "2027-09-20":"敬老の日","2027-09-23":"秋分の日","2027-10-11":"スポーツの日",
    "2027-11-03":"文化の日","2027-11-23":"勤労感謝の日",
    "2028-01-01":"元日","2028-01-10":"成人の日","2028-02-11":"建国記念の日",
    "2028-02-23":"天皇誕生日","2028-03-20":"春分の日","2028-04-29":"昭和の日",
    "2028-05-03":"憲法記念日","2028-05-04":"みどりの日","2028-05-05":"こどもの日",
    "2028-07-17":"海の日","2028-08-11":"山の日","2028-09-18":"敬老の日",
    "2028-09-22":"秋分の日","2028-10-09":"スポーツの日","2028-11-03":"文化の日",
    "2028-11-23":"勤労感謝の日",
}

# ============================================================
# 長期休暇判定（地域で調整してください）
# ============================================================
def school_vacation(m, d):
    if (m == 7 and d >= 21) or m == 8:          return True, "夏休み"
    if (m == 12 and d >= 25) or (m == 1 and d <= 7): return True, "冬休み"
    if (m == 3 and d >= 25) or (m == 4 and d <= 5):  return True, "春休み"
    return False, ""

# ============================================================
# 週間行事（平日・通常期）/ 長期休暇期
# ============================================================
WEEK_NORMAL = {
    0: "《学習支援》\n宿題・漢字・計算",
    1: "《音楽リズム》\n楽器・歌・体操",
    2: "《運動・体操》\nサーキット\nストレッチ",
    3: "《料理・調理》\n季節のクッキング",
    4: "《お楽しみ》\nゲーム・工作",
    5: "《土曜活動》\n外出・地域体験",
}
WEEK_VAC = {
    0: "《学習支援》\n宿題+工作制作",
    1: "《リトミック》\n音楽+外出散歩",
    2: "《スポーツ》\n運動遊び+自由",
    3: "《調理実習》\n季節メニュー",
    4: "《社会体験》\n買い物学習",
    5: "《土曜外出》\nおでかけ活動",
}

# ============================================================
# 月ごとの特別行事
# ============================================================
def nth_weekday(weeks, n, wd):
    """月の第n(1始まり)wd曜の日付 (0=月,6=日)"""
    hits = [w[wd] for w in weeks if w[wd] != 0]
    return hits[n-1] if len(hits) >= n else hits[-1]

def special_events(y, m):
    wks  = cal_mod.monthcalendar(y, m)
    maxd = cal_mod.monthrange(y, m)[1]
    clip = lambda d: min(max(1, d), maxd)

    ev = {}
    if m == 1:
        ev = {
            6:  "お正月遊び🎍\nカルタ・福笑い\n凧揚げ",
            8:  "書き初め✏️\n今年の目標を書こう",
            15: "冬の工作\n雪だるま・飾り",
            22: "昔遊び体験\n独楽・羽根つき",
            29: "節分準備👹\n鬼のお面作り",
        }
    elif m == 2:
        ev = {
            3:  "節分👹\n豆まき！恵方巻き",
            10: "バレンタイン準備\nカード・チョコ工作🍫",
            14: "バレンタイン🍫\n手作りチョコ",
            21: "春の予感🌱\n春探し散歩",
            28: "ひな祭り準備🎎\n雛人形工作",
        }
    elif m == 3:
        ev = {
            3:  "ひな祭り🎎\nちらし寿司・お茶会",
            10: "春の工作🌸\n桜飾り・折り紙",
            14: "ホワイトデー\nお菓子交換🍭",
            21: "春の外出🌸\n公園・自然観察",
            clip(27): "お別れ会🌸\n修了おめでとう！",
        }
    elif m == 4:
        ev = {
            6:  "入学・進級お祝い会🌸",
            13: "花見おやつ🌸\n桜もち作り",
            19: "春の外出🌸\n花見・お散歩",
            26: "こいのぼり工作🎏\n兜制作",
        }
    elif m == 5:
        md = nth_weekday(wks, 2, 6)   # 母の日: 第2日曜
        ev = {
            3:  "憲法記念日\nお休み",
            4:  "みどりの日\nお休み",
            5:  "こどもの日🎏\nパーティー！",
            clip(md-1): "母の日準備🌹\nカーネーション工作",
            md: "母の日🌹\nプレゼント贈呈",
            clip(md+5): "春の遠足\n公園・外出",
        }
    elif m == 6:
        fd = nth_weekday(wks, 3, 6)   # 父の日: 第3日曜
        ev = {
            1:         "梅雨工作☔\nカタツムリ・傘",
            8:         "あじさい工作\n折り紙・切り紙",
            clip(fd-1):"父の日準備👔\nプレゼント仕上げ",
            fd:        "父の日👔\nプレゼント贈呈",
            clip(22):  "七夕準備🎋\n短冊・飾り作り",
            clip(29):  "七夕飾り完成\n願い事発表",
        }
    elif m == 7:
        ev = {
            7:  "七夕🎋\n短冊発表！",
            13: "夏祭り準備🏮\n出し物・飾り",
            clip(19): "夏祭り🏮\nお楽しみ会",
            clip(24): "水遊び💦\nプール・シャボン玉",
            clip(28): "手作りアイス🍦\nおやつ体験",
        }
    elif m == 8:
        ev = {
            4:  "夏の工作\nスイカ・ひまわり",
            10: "流しそうめん🍜\n食育体験",
            clip(18): "お泊まり体験\n宿泊学習",
            clip(22): "夏の発表会\n制作作品展",
            clip(26): "BBQ🍖\n夏の思い出！",
            clip(31): "夏休み最終日\n宿題チェック",
        }
    elif m == 9:
        kd = nth_weekday(wks, 3, 0)   # 敬老の日: 第3月曜
        ev = {
            1:         "防災訓練⛑️\n避難・防災クイズ",
            8:         "運動会練習🏃\n徒競走・ダンス",
            15:        "ミニ運動会🏃\nがんばったね！",
            clip(kd-1):"敬老の日準備\nじいじ・ばあばへ\nプレゼント工作",
            clip(27):  "お月見🌕\n月見団子作り",
        }
    elif m == 10:
        ev = {
            5:  "芋ほり🍠\n秋の外出",
            12: "秋の工作\nどんぐり・木の実",
            19: "ハロウィン準備\n飾り・仮装工作🎃",
            26: "ハロウィンパーティー🎃\n仮装・お菓子交換",
            31: "仮装コンテスト🎃\nお菓子交換！",
        }
    elif m == 11:
        ev = {
            9:  "紅葉狩り🍁\n秋の外出",
            15: "七五三\n千歳飴・着物工作",
            21: "感謝の会 準備🌟\n練習・制作",
            clip(24): "感謝の会🌟\n発表・作品展",
            clip(30): "クリスマス準備\nリース工作🎄",
        }
    elif m == 12:
        ev = {
            1:  "クリスマス工作🎄\nリース・オーナメント",
            8:  "ケーキ作り🎂\nクリスマスケーキ",
            14: "クリスマス会🎄\nプレゼント交換！",
            21: "大掃除🧹\nきれいにしよう",
            clip(24): "冬休み開始🎍\nウィンタープログラム",
            25: "クリスマス🎄\n冬休みプログラム",
            28: "年越しそば作り🍜\nお正月準備",
            31: "大みそか\nありがとう一年間！",
        }
    return {d: v for d, v in ev.items() if 1 <= d <= maxd}

# ============================================================
# おしらせ・プログラムのねらい
# ============================================================
NOTICES = {
    1:  "あけましておめでとうございます！今年もよろしくお願いいたします。\n今月はお正月遊び（カルタ・福笑い・凧揚げ）や書き初めなど、日本の伝統文化に親しむ活動を行います。\n\n▶ 連絡事項：冬休みプログラムは1月8日（火）から通常スケジュールに戻ります。",
    2:  "節分やバレンタインなど楽しいイベントが盛りだくさんの2月です。手作りチョコレートや鬼のお面作りを楽しみましょう。\n\n▶ 感染症が流行しやすい時期です。体調不良時はご連絡をお願いします。手洗い・うがいを徹底します。",
    3:  "ひな祭りで3月のスタートです！春の訪れを感じながら、桜の工作やお別れ会で素敵な思い出を作ります。\n\n▶ 春休みプログラム（3/25〜4/5）については別途ご案内します。進級・卒業おめでとうございます！",
    4:  "桜の季節！新学期が始まりました。新しい学年でもみんなでがんばりましょう🌸\n花見散歩や春の工作で、元気にスタートします。\n\n▶ 新年度の個別支援計画・連絡帳を配布します。ご確認・押印をお願いします。",
    5:  "こどもの日・母の日がある楽しい5月！こいのぼり・兜の工作、お母さんへのプレゼント作りに取り組みます。ゴールデンウィーク明けは生活リズムを整えましょう。\n\n▶ 5/3〜5/5（祝）はお休みです。",
    6:  "梅雨でもお部屋で楽しく過ごそう！あじさい・カタツムリの工作、父の日プレゼント作り、七夕の短冊準備など盛りだくさんです🎋\n\n▶ 熱中症・体調管理に気をつけましょう。水分補給を忘れずに。",
    7:  "七夕・夏祭り・水遊びと夏イベント満載の7月！7/21（火）から夏休みプログラムに切り替わります。毎日楽しいプログラムをご用意しています🌻\n\n▶ 夏休み中（7/21〜8/31）は10:00〜17:00開所。お弁当の持参をお願いします。",
    8:  "夏休み真っ最中！お泊まり体験・流しそうめん・BBQなど夏ならではの体験がいっぱいです。熱中症に気をつけながら思い出をつくりましょう🌻\n\n▶ 宿題の進捗を定期的に確認します。宿題は必ず持参してください。",
    9:  "夏の終わり、スポーツの秋！ミニ運動会・敬老の日プレゼント作り・お月見団子など季節の活動を楽しみます🌕\n\n▶ 体を動かしやすい服装でお越しください。帽子・飲み物を持参ください。",
    10: "秋の外出シーズン！芋ほりや紅葉散策、ハロウィンパーティーを楽しもう🎃\n仮装の準備はお早めに（必要に応じてご家庭でご用意ください）。\n\n▶ 外出日は動きやすい服装・帽子・飲み物を持参してください。",
    11: "感謝の気持ちを大切に、一年の成長を振り返る11月です。発表会・感謝の会に向けて練習・準備を進めます🍁\n\n▶ 感謝の会（○月○日）へのご参加をお待ちしています。詳細は後日ご案内します。",
    12: "クリスマス・大掃除・年越し準備と、一年の締めくくりのイベントが盛りだくさん！楽しく充実した師走にしましょう🎄\n\n▶ 冬休みプログラム（12/25〜1/7）は別途ご案内。年末年始（12/29〜1/3）はお休みです。",
}
AIMS = {
    1:  "・お正月の伝統文化に親しむ\n・友達と仲良く遊ぶ\n・冬の生活リズムを整える",
    2:  "・節分の文化を楽しく体験する\n・感謝の気持ちを育てる\n・体調管理の大切さを学ぶ",
    3:  "・年度末の振り返りをする\n・友達・先生への感謝を伝える\n・春の自然を感じる",
    4:  "・新しい生活への期待と適応\n・友達関係づくりを楽しむ\n・春の自然を探して楽しむ",
    5:  "・日本の伝統行事を知る\n・家族（母）への感謝を表現する\n・連休後の生活リズムを整える",
    6:  "・梅雨でも楽しく過ごす工夫をする\n・家族（父）への感謝を伝える\n・夏に向けた準備・安全を学ぶ",
    7:  "・夏の安全（熱中症・水の事故）を学ぶ\n・水遊びのルールを守る\n・夏休みの規則正しい生活をする",
    8:  "・体験活動による主体性・協調性\n・友達・スタッフへの感謝を学ぶ\n・夏の思い出を形にする",
    9:  "・スポーツで達成感を得る\n・高齢者への敬意と感謝を育てる\n・秋の自然に親しむ",
    10: "・地域・自然体験での発見と学び\n・行事参加でコミュニケーションを図る\n・安全に外出するマナーを守る",
    11: "・一年間の成長を自覚する\n・人への感謝を言葉・作品で表す\n・秋の自然・芸術に親しむ",
    12: "・一年間を振り返る\n・冬の安全な過ごし方を知る\n・年末年始の文化に親しむ",
}
DECO = {1:"🎍",2:"👹",3:"🎎",4:"🌸",5:"🎏",6:"☔",7:"🎋",8:"🌻",9:"🌕",10:"🎃",11:"🍁",12:"🎄"}
KANJI = ["〇","一","二","三","四","五","六","七","八","九","十","十一","十二",
         "十三","十四","十五","十六","十七","十八","十九","二十"]

# ============================================================
# python-docx XML ヘルパー
# ============================================================
def _get_or_create(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el

def cell_bg(cell, hex6):
    tcPr = _get_or_create(cell._tc, 'w:tcPr')
    shd  = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex6)
    tcPr.append(shd)

def cell_borders(cell, color="538135", sz=8):
    tcPr = _get_or_create(cell._tc, 'w:tcPr')
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top','left','bottom','right'):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'),   'single')
        b.set(qn('w:sz'),    str(sz))
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)

def no_borders(table):
    tbl = table._tbl
    tblPr = _get_or_create(tbl, 'w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'none')
        tblBorders.append(el)
    tblPr.append(tblBorders)

def fixed_layout(table):
    tbl = table._tbl
    tblPr = _get_or_create(tbl, 'w:tblPr')
    lay = OxmlElement('w:tblLayout')
    lay.set(qn('w:type'), 'fixed')
    tblPr.append(lay)

def row_height(row, mm, rule='exact'):
    tr = row._tr
    trPr = _get_or_create(tr, 'w:trPr')
    trH = OxmlElement('w:trHeight')
    trH.set(qn('w:val'),   str(int(mm * 56.693)))
    trH.set(qn('w:hRule'), rule)
    trPr.append(trH)

def set_col_widths_mm(table, widths):
    tbl = table._tbl
    tblPr = _get_or_create(tbl, 'w:tblPr')
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'),    str(int(sum(widths) / 25.4 * 1440)))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)
    tblGrid = OxmlElement('w:tblGrid')
    for w in widths:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(int(w / 25.4 * 1440)))
        tblGrid.append(gc)
    tbl.insert(list(tbl).index(tblPr) + 1, tblGrid)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            tc = cell._tc
            tcPr = _get_or_create(tc, 'w:tcPr')
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'),    str(int(widths[i] / 25.4 * 1440)))
            tcW.set(qn('w:type'), 'dxa')
            tcPr.append(tcW)

def para_spacing(para, before=0, after=0, line=240):
    pPr = _get_or_create(para._p, 'w:pPr')
    sp  = OxmlElement('w:spacing')
    sp.set(qn('w:before'),   str(before))
    sp.set(qn('w:after'),    str(after))
    sp.set(qn('w:line'),     str(line))
    sp.set(qn('w:lineRule'), 'auto')
    pPr.append(sp)

def cell_valign(cell, val='top'):
    tcPr = _get_or_create(cell._tc, 'w:tcPr')
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), val)
    tcPr.append(va)

def add_run(para, text, size, bold=False, color=None):
    run = para.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    return run

def new_para(cell, align=WD_ALIGN_PARAGRAPH.LEFT, first=False):
    if first:
        p = cell.paragraphs[0]
        p.clear()
    else:
        p = cell.add_paragraph()
    p.alignment = align
    para_spacing(p)
    return p

# ============================================================
# メイン生成関数
# ============================================================
def generate(y, m):
    reiwa = y - 2018
    rstr  = KANJI[reiwa] if 0 < reiwa < len(KANJI) else str(y)
    mstr  = KANJI[m]
    deco  = DECO.get(m, "🎏")
    RED   = (0xE6, 0x00, 0x00)
    BLUE  = (0x15, 0x65, 0xC0)
    GRN   = (0x53, 0x81, 0x35)
    MAG   = (0xFF, 0x00, 0xFF)
    CYN   = (0x00, 0x99, 0xCC)
    ORG   = (0xFF, 0x8A, 0x1E)

    doc = Document()

    # ---- ページ設定 (A4 横) ----
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width   = Mm(297)
    sec.page_height  = Mm(210)
    sec.left_margin  = Mm(11)
    sec.right_margin = Mm(11)
    sec.top_margin   = Mm(9)
    sec.bottom_margin= Mm(7)

    nml = doc.styles['Normal']
    nml.paragraph_format.space_before = Pt(0)
    nml.paragraph_format.space_after  = Pt(0)

    # ============================================================
    # ① ヘッダー (3列)
    # ============================================================
    ht = doc.add_table(rows=1, cols=3)
    ht.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(ht)
    fixed_layout(ht)
    set_col_widths_mm(ht, [52, 163, 60])
    row_height(ht.rows[0], 22, 'exact')

    # 左列: 令和○年○月号
    lc = ht.rows[0].cells[0]
    cell_valign(lc, 'bottom')
    p = new_para(lc, WD_ALIGN_PARAGRAPH.LEFT, first=True)
    add_run(p, f"令和{rstr}年{mstr}月号", 13, bold=True)

    # 中央列: タイトル
    cc = ht.rows[0].cells[1]
    cell_valign(cc, 'center')
    p2 = new_para(cc, WD_ALIGN_PARAGRAPH.CENTER, first=True)
    add_run(p2, FACILITY, 13, bold=True, color=CYN)
    p3 = new_para(cc, WD_ALIGN_PARAGRAPH.CENTER)
    add_run(p3, "夢", 38, bold=True, color=MAG)
    add_run(p3, "通信", 26, bold=True, color=MAG)

    # 右列: 連絡先
    rc = ht.rows[0].cells[2]
    cell_valign(rc, 'bottom')
    p4 = new_para(rc, WD_ALIGN_PARAGRAPH.RIGHT, first=True)
    add_run(p4, f"☎  {TEL}", 10)
    p5 = new_para(rc, WD_ALIGN_PARAGRAPH.RIGHT)
    add_run(p5, f"📠  {FAX}", 10)

    # ---- デコ帯 ----
    dp = doc.add_paragraph()
    dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(dp)
    add_run(dp, (deco + " ") * 26, 8)

    # ============================================================
    # ② カレンダー本体
    # ============================================================
    weeks = cal_mod.monthcalendar(y, m)
    nw    = len(weeks)
    sp    = special_events(y, m)

    # 列幅: 275mm を 6:6:6:6:6:6:5 で分割 (合計=41*6+29=275)
    wcols = [41, 41, 41, 41, 41, 41, 29]  # mm

    ct = doc.add_table(rows=1 + nw, cols=7)
    ct.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(ct)
    fixed_layout(ct)
    set_col_widths_mm(ct, wcols)

    # ヘッダー行
    hrow = ct.rows[0]
    row_height(hrow, 8, 'exact')
    days_jp = ["月曜日","火曜日","水曜日","木曜日","金曜日","土曜日","日曜日"]
    for i, (hc, label) in enumerate(zip(hrow.cells, days_jp)):
        cell_bg(hc, "99FF66")
        cell_borders(hc, "538135", 8)
        cell_valign(hc, 'center')
        hp = new_para(hc, WD_ALIGN_PARAGRAPH.CENTER, first=True)
        col = RED if i == 6 else (BLUE if i == 5 else None)
        add_run(hp, label, 11, bold=True, color=col)

    # セル高: 5週=22mm / 6週=18mm
    rh_mm = 22 if nw <= 5 else 18

    for ri, week in enumerate(weeks):
        dr = ct.rows[ri + 1]
        row_height(dr, rh_mm, 'exact')
        for ci, day in enumerate(week):
            dc = dr.cells[ci]
            cell_borders(dc, "538135", 8)
            cell_valign(dc, 'top')

            if day == 0:
                cell_bg(dc, "F0F0F0")
                dc.paragraphs[0].clear()
                continue

            iso  = f"{y}-{m:02d}-{day:02d}"
            hol  = HOLIDAYS.get(iso)
            vac, vlabel = school_vacation(m, day)
            is_sun = (ci == 6)
            is_sat = (ci == 5)
            is_off = is_sun or (hol is not None)  # 日曜・祝日は休み

            # 背景色
            if is_off:
                cell_bg(dc, "FFF0F0")
            elif is_sat:
                cell_bg(dc, "EEF4FF")
            elif vac:
                cell_bg(dc, "FFFDE7")

            # 日付番号（右上）
            p_n = new_para(dc, WD_ALIGN_PARAGRAPH.RIGHT, first=True)
            col_n = RED if (is_off or is_sun) else (BLUE if is_sat else None)
            add_run(p_n, str(day), 13, bold=True, color=col_n)

            # 祝日名
            if hol:
                p_h = new_para(dc, WD_ALIGN_PARAGRAPH.LEFT)
                add_run(p_h, hol, 7, color=RED)

            # 長期休暇ラベル（背景色で判断できるが念のため）
            if vac and not is_off:
                p_v = new_para(dc, WD_ALIGN_PARAGRAPH.CENTER)
                add_run(p_v, f"―{vlabel}―", 7, color=(0x99, 0x66, 0x00))

            # 行事内容 (特別行事 > 週間行事)
            if day in sp:
                lines = sp[day].split('\n')
            elif is_off:
                lines = []
            elif vac:
                lines = WEEK_VAC.get(ci, "").split('\n')
            else:
                lines = WEEK_NORMAL.get(ci, "").split('\n')

            for line in lines:
                if line.strip():
                    pl = new_para(dc, WD_ALIGN_PARAGRAPH.LEFT)
                    add_run(pl, line, 8)

    # ============================================================
    # ③ 下段（おしらせ＋ねらい）
    # ============================================================
    bp = doc.add_paragraph()
    para_spacing(bp)  # spacer

    bt = doc.add_table(rows=1, cols=2)
    bt.alignment = WD_TABLE_ALIGNMENT.CENTER
    no_borders(bt)
    fixed_layout(bt)
    set_col_widths_mm(bt, [175, 100])

    nc = bt.rows[0].cells[0]   # おしらせ
    ac = bt.rows[0].cells[1]   # ねらい

    # おしらせ
    cell_borders(nc, "FF9933", 6)
    cell_bg(nc, "FFFDF8")
    cell_valign(nc, 'top')
    pnt = new_para(nc, WD_ALIGN_PARAGRAPH.CENTER, first=True)
    add_run(pnt, "おしらせ", 15, bold=True, color=ORG)
    for line in NOTICES.get(m, "").split('\n'):
        pnl = new_para(nc, WD_ALIGN_PARAGRAPH.LEFT)
        add_run(pnl, line, 8.5)

    # プログラムのねらい
    cell_borders(ac, "538135", 6)
    cell_bg(ac, "FCFFFB")
    cell_valign(ac, 'top')
    pat = new_para(ac, WD_ALIGN_PARAGRAPH.CENTER, first=True)
    add_run(pat, "プログラムのねらい", 11, bold=True)
    for line in AIMS.get(m, "").split('\n'):
        pal = new_para(ac, WD_ALIGN_PARAGRAPH.LEFT)
        add_run(pal, line, 9)

    # ============================================================
    # 保存
    # ============================================================
    fname = f"夢通信_{y}年{m:02d}月号.docx"
    doc.save(fname)
    print(f"✅ 生成完了: {fname}")
    return fname

# ============================================================
# エントリポイント
# ============================================================
if __name__ == "__main__":
    today = datetime.date.today()
    y = int(sys.argv[1]) if len(sys.argv) > 1 else today.year
    m = int(sys.argv[2]) if len(sys.argv) > 2 else today.month
    if not (2000 <= y <= 2099 and 1 <= m <= 12):
        print("引数エラー: YEAR(2000-2099) MONTH(1-12)")
        sys.exit(1)
    generate(y, m)
