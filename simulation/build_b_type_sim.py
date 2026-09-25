# 就労継続支援B型（愛媛県西条市）開設シミュレーション生成スクリプト
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

FONT = 'Meiryo UI'
BLUE = Font(name=FONT, color='0000FF'); BLACK = Font(name=FONT); GREEN = Font(name=FONT, color='008000')
BOLD = Font(name=FONT, bold=True); H1 = Font(name=FONT, bold=True, size=14); WHITE_B = Font(name=FONT, bold=True, color='FFFFFF')
YEL = PatternFill('solid', fgColor='FFFF00'); HEAD = PatternFill('solid', fgColor='4A6FC7'); SUBF = PatternFill('solid', fgColor='EEF1FB')
thin = Side(style='thin', color='D5DAE3'); BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
YEN = '#,##0;(#,##0);"-"'; MAN = '#,##0.0,,"百万"'; PCT = '0.0%;(0.0%);"-"'; NUM1 = '0.0;(0.0);"-"'

wb = Workbook()

def style_all(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != FONT:
                f = c.font
                c.font = Font(name=FONT, bold=f.bold, color=f.color, size=f.size, italic=f.italic)

def header_row(ws, r, labels, c0=1):
    for i, v in enumerate(labels):
        c = ws.cell(r, c0 + i, v); c.font = WHITE_B; c.fill = HEAD
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); c.border = BOX

# ============ 0. 使い方 ============
ws0 = wb.active; ws0.title = '使い方'
lines = [
    ('就労継続支援B型 愛媛（西条市）開設シミュレーション', H1),
    ('クロスボーダーキャリアパス構想「全国にトータルサポートを。」の第一歩として、放課後等デイサービス卒業後の受け皿となるB型事業所を愛媛で開設した場合の収支を試算します。', None),
    ('', None),
    ('■ シートの構成', BOLD),
    ('前提条件：すべての入力値。青字・黄色セルを変えると全シートが再計算されます。シナリオ（慎重／標準／積極）はここで切り替えます。', None),
    ('月次シミュレーション：開設前＋36か月の利用者数・給付費・費用・営業利益・資金残高（国保連入金は2か月遅れ）。', None),
    ('年次サマリー：1〜3年目の損益、単月黒字化月、投資回収月。', None),
    ('感度分析：平均利用者数と基本報酬区分（平均工賃月額）を変えた場合の安定期の年間営業利益。', None),
    ('開設ロードマップ：開設までの準備工程と、キャレオス内の連携先。', None),
    ('', None),
    ('■ セルの色分け', BOLD),
    ('青字＋黄色塗り＝入力する前提値（ここだけ変更してください） ／ 黒字＝計算式 ／ 緑字＝他シートからの参照', None),
    ('', None),
    ('■ 制度面の前提と出典（必ず最新の告示・愛媛県／西条市の確認を）', BOLD),
    ('・基本報酬：就労継続支援B型サービス費（Ⅰ）人員配置6:1・定員20人以下。単位数は令和6年度改定の値（837/805/758/738/726/703/673/590単位）。', None),
    ('・令和8年6月の臨時応急的な見直し：区分一〜六の平均工賃月額の下限を3,000円引き上げ（例：4.5万円→4.8万円）、中間区分（A〜F）を新設。本シートは中間区分を省略し、下の区分で算定しています（＝保守的）。', None),
    ('・令和8年6月1日以降に新規指定を受ける事業所は、令和9年度改定までの間、基本報酬が所定単位数の1000分の984。令和9年度改定（2027年度）の内容は未定のため、本シートでは3年間ずっと984/1000を掛けています（保守的）。', None),
    ('・新規指定の初年度の算定区分は「平均工賃月額1万円以上1.5万円未満」（673単位）と仮定（要確認）。', None),
    ('・地域区分：西条市は「その他」＝1単位10円と仮定（要確認）。', None),
    ('・処遇改善加算は職員の賃金改善に全額充てる前提で、収入と同額を人件費に計上（損益への影響はゼロ）。', None),
    ('・人件費・家賃などの費用はすべて仮置きの値です。実際の見積もりに置き換えてください。', None),
    ('', None),
    ('出典：GLUG「就労継続支援B型の基本報酬・加算・減算を解説【2024年度報酬改定対応】」 https://glug.co.jp/column/welfare/098', None),
    ('出典：厚生労働省「就労継続支援B型・基本報酬算定区分（令和8年6月以降分）」 https://www.mhlw.go.jp/content/001683302.xlsx', None),
    ('出典：藤原行政書士事務所「令和8年6月施行『新規指定事業所の基本報酬引き下げ』」 https://fujiwaragyoseishoshi.com/kaitei20260522/', None),
]
for i, (t, f) in enumerate(lines, 1):
    c = ws0.cell(i, 1, t); c.alignment = Alignment(wrap_text=True, vertical='top')
    if f: c.font = f
ws0.column_dimensions['A'].width = 130
c = ws0['A12']; c.font = Font(name=FONT, color='0000FF'); c.fill = YEL

# ============ 1. 前提条件 ============
ws = wb.create_sheet('前提条件')
ws['A1'] = '前提条件（入力シート）'; ws['A1'].font = H1
ws['A2'] = '青字・黄色のセルが入力値です。'
header_row(ws, 3, ['項目', '値', '単位', '根拠・メモ'])
R = {}  # name -> address
rows = [
    ('sec', '■ シナリオ'),
    ('scen', '採用シナリオ', '標準', '', '「慎重」「標準」「積極」から選択（下のシナリオ表を参照）'),
    ('sec', '■ 事業所の基本'),
    ('start', '開設月', '=DATE(2027,4,1)', '', '指定申請・物件・採用に約6か月かかる想定で2027年4月開設'),
    ('cap', '利用定員', 20, '人', '定員20人以下の報酬区分'),
    ('days', '月の開所日数', 22, '日', '平日開所（原則の日数）'),
    ('yen', '1単位の単価', 10, '円', '地域区分「その他」を仮定（西条市、要確認）'),
    ('sec', '■ 報酬（1人1日あたり）'),
    ('u1', '1年目の基本報酬単位', 673, '単位', '新規指定の初年度は「1万円以上1.5万円未満」区分と仮定（要確認）'),
    ('newadj', '新規指定事業所の特例', 0.984, '倍', '令和8年6月以降の新規指定：所定単位数×984/1000'),
    ('sogei', '送迎加算（片道）', 21, '単位', '送迎加算（Ⅰ）を仮定'),
    ('sogeirate', '送迎の利用率', 0.7, '', '利用者のうち送迎を使う割合（往復）'),
    ('meal', '食事提供体制加算', 30, '単位', '令和9年3月までの経過措置の想定（要確認）'),
    ('mealrate', '食事提供の対象率', 0.5, '', '低所得等で対象となる利用者の割合'),
    ('senmon', '福祉専門職員配置等加算', 6, '単位', '（Ⅲ）を仮定'),
    ('shogu', '処遇改善加算の率', 0.093, '', '（Ⅰ）9.3%を仮定。全額を賃金改善に充当'),
    ('sec', '■ 人件費（月額）'),
    ('mgr', '管理者 兼 サービス管理責任者 給与', 320000, '円', '仮置き。本人が管理者を担う想定'),
    ('staffpay', '職業指導員・生活支援員 給与（常勤1人）', 230000, '円', '仮置き（愛媛の相場を要確認）'),
    ('ratio', '人員配置（利用者：職員）', 6, '人', 'サービス費（Ⅰ）6:1'),
    ('minfte', '支援員の最低人数（常勤換算）', 2, '人', '立ち上げ期も最低2人は配置'),
    ('bonus', '賞与（年間の月数）', 2, 'か月', '月額に按分して計上'),
    ('welfare', '法定福利費率', 0.15, '', '社会保険料の会社負担'),
    ('driver', '送迎ドライバー（パート）', 80000, '円', '仮置き'),
    ('sec', '■ その他の費用（月額）'),
    ('rent', '家賃', 120000, '円', '西条市内・約50坪を仮定'),
    ('car', '車両リース（2台）', 70000, '円', '送迎車2台'),
    ('util', '水道光熱・通信・システム', 70000, '円', '請求ソフト含む'),
    ('other', '保険・広報・雑費', 50000, '円', ''),
    ('var', '変動費（利用1人1日あたり）', 200, '円', '消耗品・食材差額など'),
    ('honbu', '本部費（給付費に対する率）', 0.03, '', '法人本部への配賦を仮定'),
    ('sec', '■ 初期投資'),
    ('naiso', '内装・設備工事', 5000000, '円', '10年で償却'),
    ('bihin', '備品・作業機器', 1500000, '円', '5年で償却'),
    ('junbi', '開設準備費（開設前の人件費・研修・広告）', 2000000, '円', '開設月の前に一括計上'),
    ('shiki', '敷金・保証金', 500000, '円', '資産（費用にしない）'),
    ('lag', '国保連からの入金の遅れ', 2, 'か月', 'サービス提供月の翌々月に入金'),
]
r = 4
for item in rows:
    if item[0] == 'sec':
        c = ws.cell(r, 1, item[1]); c.font = BOLD; c.fill = SUBF
        for k in range(2, 5): ws.cell(r, k).fill = SUBF
        r += 1; continue
    key, label, val, unit, note = item
    ws.cell(r, 1, label); v = ws.cell(r, 2, val); ws.cell(r, 3, unit); ws.cell(r, 4, note)
    v.font = BLUE; v.fill = YEL; v.border = BOX
    if isinstance(val, float) and val < 1: v.number_format = '0.0%' if key not in ('newadj',) else '0.000'
    elif key == 'start': v.number_format = 'yyyy"年"m"月"'
    elif isinstance(val, (int, float)): v.number_format = '#,##0'
    R[key] = f"前提条件!$B${r}"; r += 1
dv = DataValidation(type='list', formula1='"慎重,標準,積極"', allow_blank=False); ws.add_data_validation(dv); dv.add(R['scen'].split('!')[1].replace('$', ''))

# Scenario table
r += 1
ws.cell(r, 1, '■ シナリオ表（利用者の集まり方と工賃）').font = BOLD; r += 1
header_row(ws, r, ['項目', '慎重', '標準', '積極', '使用値（選択中）']); r += 1
scen_rows = [
    ('s_first', '開設月の平均利用者数（1日あたり）', 3, 5, 7, NUM1),
    ('s_inc', '毎月の増加数（1日あたり）', 0.8, 1.2, 1.8, NUM1),
    ('s_occ', '目標稼働率（定員比）', 0.8, 0.9, 1.0, '0%'),
    ('s_w2', '2年目の平均工賃月額', 13000, 16000, 19000, '#,##0'),
    ('s_w3', '3年目の平均工賃月額', 16000, 19000, 24000, '#,##0'),
]
scen_hdr = r - 1
for key, label, a, b, c3, fmt in scen_rows:
    ws.cell(r, 1, label)
    for j, v in enumerate((a, b, c3)):
        cc = ws.cell(r, 2 + j, v); cc.font = BLUE; cc.fill = YEL; cc.number_format = fmt; cc.border = BOX
    u = ws.cell(r, 5, f'=INDEX(B{r}:D{r},MATCH({R["scen"]},$B${scen_hdr}:$D${scen_hdr},0))'); u.number_format = fmt; u.border = BOX
    R[key] = f"前提条件!$E${r}"; r += 1

# Reward band table
r += 1
ws.cell(r, 1, '■ 基本報酬の区分表（サービス費Ⅰ・6:1・定員20人以下）').font = BOLD; r += 1
header_row(ws, r, ['区分', '平均工賃月額の下限（円）', '単位/日', 'メモ']); r += 1
bands = [('八', 0, 590), ('七', 10000, 673), ('六', 18000, 703), ('五', 23000, 726), ('四', 28000, 738), ('三', 33000, 758), ('二', 38000, 805), ('一', 48000, 837)]
b0 = r
for name, lo, u in bands:
    ws.cell(r, 1, f'区分{name}')
    c1 = ws.cell(r, 2, lo); c2 = ws.cell(r, 3, u)
    for cc in (c1, c2): cc.font = BLUE; cc.fill = YEL; cc.number_format = '#,##0'; cc.border = BOX
    r += 1
ws.cell(b0, 4, '令和8年6月以降：区分一〜六の下限を3,000円引き上げ済み。中間区分A〜Fは省略（保守的）')
R['thr'] = f"前提条件!$B${b0}:$B${r-1}"; R['units'] = f"前提条件!$C${b0}:$C${r-1}"
ws.column_dimensions['A'].width = 40; ws.column_dimensions['B'].width = 16; ws.column_dimensions['C'].width = 10
ws.column_dimensions['D'].width = 12; ws.column_dimensions['E'].width = 16
ws.column_dimensions['D'].width = 62
ws.freeze_panes = 'A4'

# ============ 2. 月次シミュレーション ============
wm = wb.create_sheet('月次シミュレーション')
wm['A1'] = '月次シミュレーション（金額：円）'; wm['A1'].font = H1
wm['A2'] = '=CONCATENATE("シナリオ：",' + R['scen'] + ')'; wm['A2'].font = GREEN
cols = ['月番号', '年月', '年次', '平均利用者数/日', '延べ利用（人日）', '平均工賃月額', '基本報酬単位', '基本報酬', '加算', '処遇改善加算', '給付費 合計',
        '支援員（常勤換算）', '人件費', '処遇改善分の賃金', '家賃・車両・光熱等', '変動費', '本部費', '減価償却', '費用 合計', '営業利益', '累積営業利益',
        '入金（国保連）', '出金（償却除く）', '月次収支', '資金残高', '黒字月', '回収月']
header_row(wm, 4, cols)
C = {n: L(i + 1) for i, n in enumerate(cols)}
first = 5
# month 0 = 開設前
r0 = first
wm.cell(r0, 1, 0); wm.cell(r0, 2, '開設前'); wm.cell(r0, 3, 0)
for n in cols[3:]: wm[f'{C[n]}{r0}'] = 0
wm[f'{C["費用 合計"]}{r0}'] = f'={R["junbi"]}'
wm[f'{C["営業利益"]}{r0}'] = f'=-{C["費用 合計"]}{r0}'
wm[f'{C["累積営業利益"]}{r0}'] = f'={C["営業利益"]}{r0}'
wm[f'{C["出金（償却除く）"]}{r0}'] = f'={R["junbi"]}+{R["naiso"]}+{R["bihin"]}+{R["shiki"]}'
wm[f'{C["月次収支"]}{r0}'] = f'=-{C["出金（償却除く）"]}{r0}'
wm[f'{C["資金残高"]}{r0}'] = f'={C["月次収支"]}{r0}'
N = 36
for m in range(1, N + 1):
    rr = first + m; p = rr - 1
    g = lambda n: f'{C[n]}{rr}'
    wm[g('月番号')] = m
    wm[g('年月')] = f'=EDATE({R["start"]},{C["月番号"]}{rr}-1)'
    wm[g('年次')] = f'=INT(({C["月番号"]}{rr}-1)/12)+1'
    wm[g('平均利用者数/日')] = f'=MIN({R["cap"]}*{R["s_occ"]},{R["s_first"]}+({C["月番号"]}{rr}-1)*{R["s_inc"]})'
    wm[g('延べ利用（人日）')] = f'={g("平均利用者数/日")}*{R["days"]}'
    wm[g('平均工賃月額')] = f'=IF({g("年次")}=1,"初年度",IF({g("年次")}=2,{R["s_w2"]},{R["s_w3"]}))'
    wm[g('基本報酬単位')] = f'=IF({g("年次")}=1,{R["u1"]},INDEX({R["units"]},MATCH({g("平均工賃月額")},{R["thr"]},1)))'
    wm[g('基本報酬')] = f'={g("延べ利用（人日）")}*{g("基本報酬単位")}*{R["newadj"]}*{R["yen"]}'
    wm[g('加算')] = f'={g("延べ利用（人日）")}*({R["sogei"]}*2*{R["sogeirate"]}+{R["meal"]}*{R["mealrate"]}+{R["senmon"]})*{R["yen"]}'
    wm[g('処遇改善加算')] = f'=({g("基本報酬")}+{g("加算")})*{R["shogu"]}'
    wm[g('給付費 合計')] = f'={g("基本報酬")}+{g("加算")}+{g("処遇改善加算")}'
    wm[g('支援員（常勤換算）')] = f'=MAX({R["minfte"]},CEILING({g("平均利用者数/日")}/{R["ratio"]},0.5))'
    wm[g('人件費')] = f'=({R["mgr"]}+{g("支援員（常勤換算）")}*{R["staffpay"]})*(12+{R["bonus"]})/12*(1+{R["welfare"]})+{R["driver"]}'
    wm[g('処遇改善分の賃金')] = f'={g("処遇改善加算")}'
    wm[g('家賃・車両・光熱等')] = f'={R["rent"]}+{R["car"]}+{R["util"]}+{R["other"]}'
    wm[g('変動費')] = f'={g("延べ利用（人日）")}*{R["var"]}'
    wm[g('本部費')] = f'={g("給付費 合計")}*{R["honbu"]}'
    wm[g('減価償却')] = f'={R["naiso"]}/120+{R["bihin"]}/60'
    wm[g('費用 合計')] = f'=SUM({C["人件費"]}{rr}:{C["減価償却"]}{rr})'
    wm[g('営業利益')] = f'={g("給付費 合計")}-{g("費用 合計")}'
    wm[g('累積営業利益')] = f'={C["累積営業利益"]}{p}+{g("営業利益")}'
    wm[g('入金（国保連）')] = f'=IF({C["月番号"]}{rr}>{R["lag"]},OFFSET({g("給付費 合計")},-{R["lag"]},0),0)'
    wm[g('出金（償却除く）')] = f'={g("費用 合計")}-{g("減価償却")}'
    wm[g('月次収支')] = f'={g("入金（国保連）")}-{g("出金（償却除く）")}'
    wm[g('資金残高')] = f'={C["資金残高"]}{p}+{g("月次収支")}'
    wm[g('黒字月')] = f'=IF({g("営業利益")}>0,1,0)'
    wm[g('回収月')] = f'=IF({g("資金残高")}>=0,1,0)'
last = first + N
for rr in range(first, last + 1):
    for i, n in enumerate(cols):
        c = wm.cell(rr, i + 1); c.border = BOX
        if n == '年月' and rr > first: c.number_format = 'yyyy"年"m"月"'
        elif n in ('平均利用者数/日', '支援員（常勤換算）'): c.number_format = NUM1
        elif n in ('月番号', '年次', '黒字月', '回収月'): c.number_format = '0'
        elif n not in ('年月',): c.number_format = YEN
    if rr > first and (rr - first) % 12 == 0:
        for i in range(len(cols)): wm.cell(rr, i + 1).fill = SUBF
red = Font(name=FONT, color='C00000')
for n in ('営業利益', '資金残高', '累積営業利益'):
    wm.conditional_formatting.add(f'{C[n]}{first}:{C[n]}{last}', CellIsRule(operator='lessThan', formula=['0'], font=red))
for i in range(len(cols)): wm.column_dimensions[L(i + 1)].width = 13
wm.column_dimensions['B'].width = 11; wm.row_dimensions[4].height = 34
wm.freeze_panes = 'C5'
M = lambda n: f"月次シミュレーション!${C[n]}${first + 1}:${C[n]}${last}"

# ============ 3. 年次サマリー ============
wy = wb.create_sheet('年次サマリー')
wy['A1'] = '年次サマリー（金額：円）'; wy['A1'].font = H1
wy['A2'] = '=CONCATENATE("シナリオ：",' + R['scen'] + ')'; wy['A2'].font = GREEN
header_row(wy, 4, ['項目', '開設前', '1年目', '2年目', '3年目', '3年間 合計'])
items = [('年度末の平均利用者数/日', 'end_users'), ('延べ利用（人日）', '延べ利用（人日）'), ('基本報酬', '基本報酬'), ('加算', '加算'), ('処遇改善加算', '処遇改善加算'),
         ('給付費 合計', '給付費 合計'), ('人件費（処遇改善分含む）', 'staff'), ('家賃・車両・光熱等', '家賃・車両・光熱等'), ('変動費', '変動費'), ('本部費', '本部費'),
         ('減価償却', '減価償却'), ('開設準備費', 'junbi'), ('費用 合計', 'cost'), ('営業利益', 'op'), ('営業利益率', 'margin'), ('年度末の資金残高', 'cash')]
YR = {}
for i, (label, key) in enumerate(items):
    rr = 5 + i; YR[key] = rr; wy.cell(rr, 1, label)
    for j, y in enumerate([0, 1, 2, 3]):
        col = L(2 + j); c = wy[f'{col}{rr}']
        s = lambda n: f'SUMIFS({M(n)},{M("年次")},{y})'
        if key == 'end_users':
            c.value = 0 if y == 0 else f"=INDEX({M('平均利用者数/日')},{12*y})"
        elif key == 'staff':
            c.value = 0 if y == 0 else f'={s("人件費")}+{s("処遇改善分の賃金")}'
        elif key == 'junbi':
            c.value = f'={R["junbi"]}' if y == 0 else 0
        elif key == 'cost':
            c.value = f'=SUM({col}{YR["staff"]}:{col}{rr-1})'
        elif key == 'op':
            c.value = f'={col}{YR["給付費 合計"]}-{col}{YR["cost"]}'
        elif key == 'margin':
            c.value = f'=IF({col}{YR["給付費 合計"]}=0,0,{col}{YR["op"]}/{col}{YR["給付費 合計"]})'
        elif key == 'cash':
            c.value = f"=月次シミュレーション!${C['資金残高']}${first}" if y == 0 else f"=INDEX({M('資金残高')},{12*y})"
        else:
            c.value = 0 if y == 0 else f'={s(key)}'
    tot = wy[f'F{rr}']
    if key in ('end_users', 'cash'): tot.value = f'=E{rr}'
    elif key == 'margin': tot.value = f'=IF(F{YR["給付費 合計"]}=0,0,F{YR["op"]}/F{YR["給付費 合計"]})'
    else: tot.value = f'=SUM(B{rr}:E{rr})'
    for j in range(6):
        c = wy.cell(rr, 1 + j); c.border = BOX
        if j: c.number_format = PCT if key == 'margin' else (NUM1 if key == 'end_users' else YEN)
    if key in ('給付費 合計', 'cost', 'op', 'cash'):
        for j in range(6): wy.cell(rr, 1 + j).font = BOLD
rr = 5 + len(items) + 1
wy.cell(rr, 1, '■ 重要指標').font = BOLD; rr += 1
kpis = [
    ('単月黒字になる月（開設から）', f'=IFERROR(MATCH(1,{M("黒字月")},0)&"か月目","36か月内は未達")'),
    ('単月黒字になる年月', f'=IFERROR(INDEX({M("年月")},MATCH(1,{M("黒字月")},0)),"-")'),
    ('初期投資を回収する月（資金残高がプラス）', f'=IFERROR(MATCH(1,{M("回収月")},0)&"か月目","36か月内は未達")'),
    ('資金残高の最低額（必要な運転資金の目安）', f"=MIN(月次シミュレーション!${C['資金残高']}${first}:${C['資金残高']}${last})"),
    ('初期投資の合計', f'={R["naiso"]}+{R["bihin"]}+{R["junbi"]}+{R["shiki"]}'),
]
for label, f in kpis:
    wy.cell(rr, 1, label); c = wy.cell(rr, 2, f); c.font = Font(name=FONT, bold=True, color='4A6FC7'); c.border = BOX
    if '年月' in label: c.number_format = 'yyyy"年"m"月"'
    elif '額' in label or '合計' in label: c.number_format = YEN
    rr += 1
wy.conditional_formatting.add(f'B{YR["op"]}:F{YR["cash"]}', CellIsRule(operator='lessThan', formula=['0'], font=red))
wy.column_dimensions['A'].width = 40
for j in range(2, 7): wy.column_dimensions[L(j)].width = 15

# ============ 4. 感度分析 ============
wsn = wb.create_sheet('感度分析')
wsn['A1'] = '感度分析：安定期の年間営業利益（円）'; wsn['A1'].font = H1
wsn['A2'] = '行＝1日の平均利用者数、列＝基本報酬の区分（平均工賃月額）。新規特例・加算・人件費などは「前提条件」の値を使用。'
wsn['A4'] = '平均利用者数/日 ＼ 基本報酬単位'; wsn['A4'].font = WHITE_B; wsn['A4'].fill = HEAD
unit_opts = [(673, '1万〜1.8万円'), (703, '1.8万〜2.3万円'), (726, '2.3万〜2.8万円'), (738, '2.8万〜3.3万円'), (758, '3.3万〜3.8万円')]
for j, (u, lab) in enumerate(unit_opts):
    c = wsn.cell(3, 2 + j, lab); c.font = BOLD; c.alignment = Alignment(horizontal='center')
    c = wsn.cell(4, 2 + j, u); c.font = Font(name=FONT, bold=True, color='0000FF'); c.fill = YEL; c.number_format = '#,##0"単位"'; c.alignment = Alignment(horizontal='center')
for i, users in enumerate(range(10, 21)):
    rr = 5 + i
    c = wsn.cell(rr, 1, users); c.font = Font(name=FONT, bold=True, color='0000FF'); c.fill = YEL; c.number_format = '0"人"'
    for j in range(len(unit_opts)):
        col = L(2 + j); u = f'{col}$4'; n = f'$A{rr}'
        rev = f'({n}*{R["days"]}*({u}*{R["newadj"]}+{R["sogei"]}*2*{R["sogeirate"]}+{R["meal"]}*{R["mealrate"]}+{R["senmon"]})*{R["yen"]})'
        staff = f'(({R["mgr"]}+MAX({R["minfte"]},CEILING({n}/{R["ratio"]},0.5))*{R["staffpay"]})*(12+{R["bonus"]})/12*(1+{R["welfare"]})+{R["driver"]})'
        cost = f'({staff}+{R["rent"]}+{R["car"]}+{R["util"]}+{R["other"]}+{n}*{R["days"]}*{R["var"]}+{rev}*(1+{R["shogu"]})*{R["honbu"]}+{R["naiso"]}/120+{R["bihin"]}/60)'
        c = wsn.cell(rr, 2 + j, f'=12*({rev}-{cost})'); c.number_format = YEN; c.border = BOX
wsn.conditional_formatting.add('B5:F15', CellIsRule(operator='lessThan', formula=['0'], font=red, fill=PatternFill('solid', fgColor='FDE8EA')))
wsn.conditional_formatting.add('B5:F15', CellIsRule(operator='greaterThan', formula=['0'], fill=PatternFill('solid', fgColor='E6F4EA')))
wsn['A17'] = '読み方：赤＝赤字、緑＝黒字。処遇改善加算は同額を賃金に充てるため、損益には含めていません。'
wsn.column_dimensions['A'].width = 32
for j in range(2, 7): wsn.column_dimensions[L(j)].width = 17

# ============ 5. 開設ロードマップ ============
wr = wb.create_sheet('開設ロードマップ')
wr['A1'] = '開設ロードマップ（2027年4月開設の場合）'; wr['A1'].font = H1
header_row(wr, 3, ['時期', 'やること', 'ポイント', 'キャレオス内の連携先（クロスボーダー）'])
steps = [
    ('2026年10〜11月', '市場と制度の確認', '西条市・今治市の障がい福祉担当課へ事前相談。障害福祉計画のB型の見込量と、総量規制の有無を確認。地域区分・新規指定の算定区分も確認。', '第三管理部、地域共創課'),
    ('2026年11〜12月', '利用ニーズの把握', '放デイ（日吉・西条）の中高生の人数と卒業時期、保護者アンケート。特別支援学校の進路担当・相談支援事業所へのヒアリング。', '相談支援、放デイ各事業所'),
    ('2026年12月', '法人内の事業計画・決裁', '本シミュレーションを実数に置き換え、投資額・回収計画を役員へ提示。', '第三管理部'),
    ('2027年1月', '物件の確保', '送迎しやすい立地、作業スペース、用途・消防の適合を確認。', '地域共創課'),
    ('2027年1〜2月', '生産活動（作業）の開拓', '地域企業の軽作業、農福連携（西条の農業）、飲食・製菓など、平均工賃を上げられる作業を確保。', 'ラトリエトゥボナペティ、ソーシャルカレッジリッツ、地域共創課'),
    ('2027年1〜3月', '人員の採用と育成', 'サービス管理責任者の要件（実務経験・研修）を確認。職業指導員・生活支援員を採用し、既存事業所で研修。', '生活介護、就労系事業所'),
    ('2027年2月', '指定申請', '愛媛県へ指定申請（締切・事前協議の時期を確認）。', '第三管理部'),
    ('2027年4月', '開設', '放デイ卒業生・特別支援学校卒業生の受け入れ。初年度は利用者の確保と工賃の底上げに集中。', '全事業部'),
    ('2028年〜', '次の展開', '平均工賃の向上で報酬区分を上げる。放課後児童クラブ（委託）や相談支援との連携で、愛媛のトータルサポートの形をつくる。', '四国支社構想へ'),
]
for i, s in enumerate(steps):
    for j, v in enumerate(s):
        c = wr.cell(4 + i, 1 + j, v); c.alignment = Alignment(wrap_text=True, vertical='top'); c.border = BOX
for col, w in zip('ABCD', (16, 22, 70, 34)): wr.column_dimensions[col].width = w

for sh in wb.worksheets: style_all(sh)
wb.save('/home/user/-/simulation/B型_愛媛_開設シミュレーション.xlsx')
print('saved')
