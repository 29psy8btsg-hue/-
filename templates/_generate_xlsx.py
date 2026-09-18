# -*- coding: utf-8 -*-
"""議事録テンプレート（Excel）を生成する。"""
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

JP = '游ゴシック'
NAVY = '1F3B63'
BAND = PatternFill('solid', fgColor=NAVY)          # セクション見出し
HEAD = PatternFill('solid', fgColor='DCE6F1')      # 表のヘッダー
LABEL = PatternFill('solid', fgColor='F2F2F2')     # 項目名
INPUT = PatternFill('solid', fgColor='FFFBE6')     # 入力欄（うすい黄色）
CALC = PatternFill('solid', fgColor='EDEDED')      # 数式（自動計算）
EX = PatternFill('solid', fgColor='EAF3EA')        # 記入例

THIN = Side(style='thin', color='BFBFBF')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

LAST = 201          # アクション管理の最終行
MEET_LAST = 101     # 会議ログの最終行


def f(size=11, bold=False, color='000000', italic=False):
    return Font(name=JP, size=size, bold=bold, color=color, italic=italic)


def put(ws, ref, value=None, font=None, fill=None, align=None, border=True, fmt=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = font or f()
    if fill is not None:
        c.fill = fill
    if align is not None:
        c.alignment = align
    if border:
        c.border = BOX
    if fmt:
        c.number_format = fmt
    return c


LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
LEFT_TOP = Alignment(horizontal='left', vertical='top', wrap_text=True)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def style_range(ws, ref, fill=None, align=None, border=True, font=None, fmt=None):
    """結合範囲も含め、範囲内の全セルに罫線・塗りを適用する。"""
    for row in ws[ref]:
        for c in row:
            c.font = font or f()
            if fill is not None:
                c.fill = fill
            if align is not None:
                c.alignment = align
            if border:
                c.border = BOX
            if fmt:
                c.number_format = fmt


def section(ws, row, text, last_col='G'):
    ws.merge_cells(f'B{row}:{last_col}{row}')
    style_range(ws, f'B{row}:{last_col}{row}', fill=BAND, align=LEFT,
                font=f(11, bold=True, color='FFFFFF'))
    ws[f'B{row}'] = text
    ws.row_dimensions[row].height = 22


def page_fit(ws, area=None, landscape=False, repeat_head=False):
    """A4・横幅1ページに収める印刷設定。"""
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    if area:
        ws.print_area = area
    if repeat_head:
        ws.print_title_rows = '1:1'

# ---------------------------------------------------------------- 議事録シート
def build_minutes(ws):
    ws.sheet_view.showGridLines = False
    widths = {'A': 2.5, 'B': 14, 'C': 30, 'D': 16, 'E': 14, 'F': 12, 'G': 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells('B2:G2')
    ws['B2'] = '議 事 録'
    style_range(ws, 'B2:G2', align=CENTER, border=False,
                font=f(20, bold=True, color=NAVY))
    ws.row_dimensions[2].height = 32

    # --- 基本情報 -------------------------------------------------
    info = [
        (4, '会議名', 'C4:D4', '文書番号', 'F4:G4'),
        (5, '開催日時', 'C5:D5', '開催場所', 'F5:G5'),
        (6, '主催・議長', 'C6:D6', '書記（作成者）', 'F6:G6'),
    ]
    for row, l1, r1, l2, r2 in info:
        put(ws, f'B{row}', l1, font=f(10, bold=True), fill=LABEL, align=LEFT)
        ws.merge_cells(r1)
        style_range(ws, r1, fill=INPUT, align=LEFT)
        put(ws, f'E{row}', l2, font=f(10, bold=True), fill=LABEL, align=LEFT)
        ws.merge_cells(r2)
        style_range(ws, r2, fill=INPUT, align=LEFT)
        ws.row_dimensions[row].height = 20

    for row, label in ((7, '出席者'), (8, '欠席者'), (9, '配布先')):
        put(ws, f'B{row}', label, font=f(10, bold=True), fill=LABEL, align=LEFT)
        ws.merge_cells(f'C{row}:G{row}')
        style_range(ws, f'C{row}:G{row}', fill=INPUT, align=LEFT)
        ws.row_dimensions[row].height = 20
    ws['C5'].number_format = '@'

    # --- 1. 目的 ---------------------------------------------------
    section(ws, 11, '1. 会議の目的・ゴール（この会議で何が決まれば成功か）')
    ws.merge_cells('B12:G13')
    style_range(ws, 'B12:G13', fill=INPUT, align=LEFT_TOP)

    # --- 2. アジェンダ ---------------------------------------------
    section(ws, 15, '2. アジェンダ（議題）')
    for col, head in zip('BCDEF', ['No', '議題', '担当', '予定時間', '資料']):
        put(ws, f'{col}16', head, font=f(10, bold=True), fill=HEAD, align=CENTER)
    ws.merge_cells('F16:G16')
    style_range(ws, 'F16:G16', fill=HEAD, align=CENTER, font=f(10, bold=True))
    ws['F16'] = '資料'
    for i, row in enumerate(range(17, 21), start=1):
        put(ws, f'B{row}', i, align=CENTER, fill=INPUT)
        for col in 'CDE':
            put(ws, f'{col}{row}', fill=INPUT, align=LEFT)
        ws.merge_cells(f'F{row}:G{row}')
        style_range(ws, f'F{row}:G{row}', fill=INPUT, align=LEFT)
        ws.row_dimensions[row].height = 20

    # --- 3. 決定事項 -----------------------------------------------
    section(ws, 22, '3. 決定事項（決まったことだけを書く）')
    put(ws, 'B23', 'No', font=f(10, bold=True), fill=HEAD, align=CENTER)
    put(ws, 'C23', '決定内容', font=f(10, bold=True), fill=HEAD, align=CENTER)
    ws.merge_cells('D23:E23')
    style_range(ws, 'D23:E23', fill=HEAD, align=CENTER, font=f(10, bold=True))
    ws['D23'] = '決定理由・背景'
    put(ws, 'F23', '決定者', font=f(10, bold=True), fill=HEAD, align=CENTER)
    put(ws, 'G23', '関連資料', font=f(10, bold=True), fill=HEAD, align=CENTER)
    for i, row in enumerate(range(24, 28), start=1):
        put(ws, f'B{row}', f'D-{i}', align=CENTER, fill=INPUT)
        put(ws, f'C{row}', fill=INPUT, align=LEFT)
        ws.merge_cells(f'D{row}:E{row}')
        style_range(ws, f'D{row}:E{row}', fill=INPUT, align=LEFT)
        put(ws, f'F{row}', fill=INPUT, align=LEFT)
        put(ws, f'G{row}', fill=INPUT, align=LEFT)
        ws.row_dimensions[row].height = 24

    # --- 4. アクションアイテム --------------------------------------
    section(ws, 29, '4. アクションアイテム（誰が・何を・いつまでに）')
    for col, head in zip('BCDEFG', ['No', '対応内容', '担当者', '期限', 'ステータス', '備考']):
        put(ws, f'{col}30', head, font=f(10, bold=True), fill=HEAD, align=CENTER)
    for i, row in enumerate(range(31, 36), start=1):
        put(ws, f'B{row}', f'A-{i}', align=CENTER, fill=INPUT)
        put(ws, f'C{row}', fill=INPUT, align=LEFT)
        put(ws, f'D{row}', fill=INPUT, align=LEFT)
        put(ws, f'E{row}', fill=INPUT, align=CENTER, fmt='yyyy/mm/dd')
        put(ws, f'F{row}', fill=INPUT, align=CENTER)
        put(ws, f'G{row}', fill=INPUT, align=LEFT)
        ws.row_dimensions[row].height = 24
    dv = DataValidation(type='list', formula1='=設定!$A$2:$A$6', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add('F31:F35')
    put(ws, 'B36', '※ 会議後、31〜35行目をコピーして「アクション管理」シートに貼り付けると進捗を追えます。',
        font=f(9, color='808080'), align=LEFT, border=False)
    ws.merge_cells('B36:G36')

    # --- 5. 議事内容 -----------------------------------------------
    section(ws, 38, '5. 議事内容（議論の要点と結論）')
    row = 39
    for i in (1, 2, 3):
        put(ws, f'B{row}', f'議題{i}', font=f(10, bold=True), fill=LABEL, align=CENTER)
        ws.merge_cells(f'C{row}:G{row}')
        style_range(ws, f'C{row}:G{row}', fill=INPUT, align=LEFT)
        put(ws, f'B{row+1}', '議論の要点', font=f(10, bold=True), fill=LABEL, align=LEFT)
        ws.merge_cells(f'C{row+1}:G{row+2}')
        style_range(ws, f'C{row+1}:G{row+2}', fill=INPUT, align=LEFT_TOP)
        ws.merge_cells(f'B{row+1}:B{row+2}')
        style_range(ws, f'B{row+1}:B{row+2}', fill=LABEL, align=LEFT, font=f(10, bold=True))
        ws[f'B{row+1}'] = '議論の要点'
        put(ws, f'B{row+3}', '結論', font=f(10, bold=True), fill=LABEL, align=LEFT)
        ws.merge_cells(f'C{row+3}:G{row+3}')
        style_range(ws, f'C{row+3}:G{row+3}', fill=INPUT, align=LEFT_TOP)
        for r in range(row, row + 4):
            ws.row_dimensions[r].height = 22
        row += 5

    # --- 6. 保留 ---------------------------------------------------
    section(ws, row, '6. 保留・継続検討事項（決まらなかったこと）')
    hr = row + 1
    put(ws, f'B{hr}', 'No', font=f(10, bold=True), fill=HEAD, align=CENTER)
    put(ws, f'C{hr}', '内容', font=f(10, bold=True), fill=HEAD, align=CENTER)
    ws.merge_cells(f'D{hr}:E{hr}')
    style_range(ws, f'D{hr}:E{hr}', fill=HEAD, align=CENTER, font=f(10, bold=True))
    ws[f'D{hr}'] = '保留の理由'
    put(ws, f'F{hr}', '再検討の時期', font=f(10, bold=True), fill=HEAD, align=CENTER)
    put(ws, f'G{hr}', '担当', font=f(10, bold=True), fill=HEAD, align=CENTER)
    for i in range(1, 4):
        r = hr + i
        put(ws, f'B{r}', f'P-{i}', align=CENTER, fill=INPUT)
        put(ws, f'C{r}', fill=INPUT, align=LEFT)
        ws.merge_cells(f'D{r}:E{r}')
        style_range(ws, f'D{r}:E{r}', fill=INPUT, align=LEFT)
        put(ws, f'F{r}', fill=INPUT, align=LEFT)
        put(ws, f'G{r}', fill=INPUT, align=LEFT)
        ws.row_dimensions[r].height = 22
    row = hr + 5

    # --- 7. 次回会議 -----------------------------------------------
    section(ws, row, '7. 次回会議の予定')
    for i, (l1, l2) in enumerate([('日時', '場所'), ('主な議題', '事前準備')]):
        r = row + 1 + i
        put(ws, f'B{r}', l1, font=f(10, bold=True), fill=LABEL, align=LEFT)
        ws.merge_cells(f'C{r}:D{r}')
        style_range(ws, f'C{r}:D{r}', fill=INPUT, align=LEFT)
        put(ws, f'E{r}', l2, font=f(10, bold=True), fill=LABEL, align=LEFT)
        ws.merge_cells(f'F{r}:G{r}')
        style_range(ws, f'F{r}:G{r}', fill=INPUT, align=LEFT)
        ws.row_dimensions[r].height = 20

    # --- 印刷設定 ---------------------------------------------------
    page_fit(ws, area=f'A1:G{row + 3}')
    ws.freeze_panes = 'A3'


# ------------------------------------------------------- アクション管理シート
def build_actions(ws):
    heads = ['会議日', '会議名', 'No', 'アクション内容', '担当者',
             '期限', 'ステータス', '残日数', '完了日', '備考']
    widths = [12, 20, 7, 40, 12, 12, 12, 9, 12, 24]
    for i, (h, w) in enumerate(zip(heads, widths), start=1):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = w
        put(ws, f'{col}1', h, font=f(10, bold=True, color='FFFFFF'), fill=BAND, align=CENTER)
    ws.row_dimensions[1].height = 24

    # 記入例（1行目）
    example = [date(2026, 9, 18), '営業定例', 'A-1', '新プランの価格案を3パターン作成する',
               '山田', date(2026, 9, 25), '対応中', None, None, '←記入例です。行ごと削除して使ってください']
    for i, v in enumerate(example, start=1):
        col = get_column_letter(i)
        cell = put(ws, f'{col}2', v, fill=EX, align=LEFT if i in (2, 4, 10) else CENTER)
        if i in (1, 6, 9):
            cell.number_format = 'yyyy/mm/dd'

    for r in range(2, LAST + 1):
        for i in range(1, 11):
            col = get_column_letter(i)
            c = ws[f'{col}{r}']
            c.font = f()
            c.border = BOX
            c.alignment = LEFT if i in (2, 4, 10) else CENTER
            if i in (1, 6, 9):
                c.number_format = 'yyyy/mm/dd'
            if r > 2:
                c.fill = CALC if i == 8 else INPUT
        # 残日数（期限までの日数。完了・中止は空欄）
        h = ws[f'H{r}']
        h.value = f'=IF(OR($F{r}="",$G{r}="完了",$G{r}="中止"),"",$F{r}-TODAY())'
        h.fill = CALC
        h.number_format = '0"日";0"日超過";"当日"'
    ws['H2'].fill = EX

    dv = DataValidation(type='list', formula1='=設定!$A$2:$A$6', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f'G2:G{LAST}')

    rng = f'A2:J{LAST}'
    # 期限超過＝行全体を赤
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND($D2<>"",$F2<>"",$F2<TODAY(),$G2<>"完了",$G2<>"中止")'],
        fill=PatternFill('solid', fgColor='FFD9D9'), stopIfTrue=False))
    # 期限まで3日以内＝黄色
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND($D2<>"",$F2<>"",$F2>=TODAY(),$F2-TODAY()<=3,$G2<>"完了",$G2<>"中止")'],
        fill=PatternFill('solid', fgColor='FFF2CC'), stopIfTrue=False))
    # 完了＝グレー
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=['$G2="完了"'],
        fill=PatternFill('solid', fgColor='EDEDED'),
        font=Font(name=JP, size=11, color='909090'), stopIfTrue=False))

    ws.auto_filter.ref = f'A1:J{LAST}'
    ws.freeze_panes = 'A2'
    page_fit(ws, area='A1:J60', landscape=True, repeat_head=True)


# ------------------------------------------------------------ 会議ログシート
def build_log(ws):
    heads = ['開催日', '会議名', '主催', '出席人数', '決定事項数',
             'ToDo件数', '未完了', '議事録の保存先']
    widths = [12, 22, 12, 10, 12, 10, 10, 34]
    for i, (h, w) in enumerate(zip(heads, widths), start=1):
        col = get_column_letter(i)
        ws.column_dimensions[col].width = w
        put(ws, f'{col}1', h, font=f(10, bold=True, color='FFFFFF'), fill=BAND, align=CENTER)
    ws.row_dimensions[1].height = 24

    ws['A2'] = date(2026, 9, 18)
    ws['B2'] = '営業定例'
    ws['C2'] = '田中'
    ws['D2'] = 6
    ws['E2'] = 3
    ws['H2'] = '20260918_営業定例_議事録.docx'

    for r in range(2, MEET_LAST + 1):
        for i in range(1, 9):
            col = get_column_letter(i)
            c = ws[f'{col}{r}']
            c.font = f()
            c.border = BOX
            c.alignment = LEFT if i in (2, 8) else CENTER
            if i == 1:
                c.number_format = 'yyyy/mm/dd'
            if r > 2:
                c.fill = CALC if i in (6, 7) else INPUT
        # ToDo件数／未完了件数（アクション管理シートから自動集計）
        ws[f'F{r}'] = (f'=IF($B{r}="","",COUNTIFS(アクション管理!$B$2:$B${LAST},$B{r}))')
        ws[f'G{r}'] = (f'=IF($B{r}="","",COUNTIFS(アクション管理!$B$2:$B${LAST},$B{r})'
                       f'-COUNTIFS(アクション管理!$B$2:$B${LAST},$B{r},アクション管理!$G$2:$G${LAST},"完了")'
                       f'-COUNTIFS(アクション管理!$B$2:$B${LAST},$B{r},アクション管理!$G$2:$G${LAST},"中止"))')
        ws[f'F{r}'].fill = CALC
        ws[f'G{r}'].fill = CALC
    for i in (1, 2, 3, 4, 5, 8):
        ws[f'{get_column_letter(i)}2'].fill = EX
    ws['F2'].fill = EX
    ws['G2'].fill = EX
    ws['I2'] = '←記入例です。行ごと削除して使ってください'
    ws['I2'].font = f(9, color='808080')
    ws.column_dimensions['I'].width = 36

    ws.auto_filter.ref = f'A1:H{MEET_LAST}'
    ws.freeze_panes = 'A2'
    page_fit(ws, area='A1:H40', landscape=True, repeat_head=True)


# -------------------------------------------------------------- 使い方シート
def build_guide(ws):
    ws.sheet_view.showGridLines = False
    for col, w in zip('ABCDEF', [2.5, 20, 46, 14, 14, 14]):
        ws.column_dimensions[col].width = w

    ws.merge_cells('B2:F2')
    ws['B2'] = '議事録テンプレート（Excel版）の使い方'
    style_range(ws, 'B2:F2', align=LEFT, border=False, font=f(18, bold=True, color=NAVY))
    ws.row_dimensions[2].height = 30

    ws.merge_cells('B3:F3')
    ws['B3'] = '「議事録」シートで1回分を記録し、ToDoだけを「アクション管理」シートに集めて進捗を追いかける構成です。'
    style_range(ws, 'B3:F3', align=LEFT, border=False, font=f(10, color='555555'))

    # --- シート構成 ---
    section(ws, 5, 'シートの構成と使い分け', last_col='F')
    put(ws, 'B6', 'シート名', font=f(10, bold=True), fill=HEAD, align=CENTER)
    ws.merge_cells('C6:F6')
    style_range(ws, 'C6:F6', fill=HEAD, align=CENTER, font=f(10, bold=True))
    ws['C6'] = '何をするシートか'
    sheets = [
        ('議事録', '1回分の議事録。会議ごとにシートを右クリック →「移動またはコピー」→「コピーを作成する」で複製し、シート名を「0918_営業定例」などに変えて使います。'),
        ('アクション管理', '全会議のToDoを1か所に集めるシート。期限切れは自動で赤、期限3日前は黄色になります。フィルターで担当者ごとに絞り込めます。'),
        ('会議ログ', '開催した会議の一覧。ToDo件数と未完了件数はアクション管理シートから自動集計されます。'),
        ('設定', 'ステータスの選択肢を管理するシート。言葉を変えたいときはここを編集します。'),
    ]
    r = 7
    for name, desc in sheets:
        put(ws, f'B{r}', name, font=f(10, bold=True), fill=LABEL, align=LEFT)
        ws.merge_cells(f'C{r}:F{r}')
        style_range(ws, f'C{r}:F{r}', align=LEFT)
        ws[f'C{r}'] = desc
        ws.row_dimensions[r].height = 34
        r += 1

    # --- 色の凡例 ---
    section(ws, 12, '色の意味', last_col='F')
    legend = [
        (INPUT, 'うすい黄色', '入力する欄です。ここだけ書き換えてください。'),
        (CALC, 'グレー', '数式が入っている自動計算欄です。上書きすると計算が壊れます。'),
        (EX, 'うすい緑', '記入例です。使い始めるときに行ごと削除してください。'),
        (PatternFill('solid', fgColor='FFD9D9'), '赤（自動）', 'アクション管理シートで期限を過ぎた未完了のToDoです。'),
        (PatternFill('solid', fgColor='FFF2CC'), '黄（自動）', 'アクション管理シートで期限まで3日以内のToDoです。'),
    ]
    r = 13
    for fill, name, desc in legend:
        put(ws, f'B{r}', name, fill=fill, align=CENTER, font=f(10, bold=True))
        ws.merge_cells(f'C{r}:F{r}')
        style_range(ws, f'C{r}:F{r}', align=LEFT)
        ws[f'C{r}'] = desc
        ws.row_dimensions[r].height = 20
        r += 1

    # --- 進捗サマリー ---
    section(ws, 19, 'ToDoの進捗（アクション管理シートから自動集計）', last_col='F')
    put(ws, 'B20', '本日', font=f(10, bold=True), fill=LABEL, align=CENTER)
    put(ws, 'C20', '=TODAY()', fill=CALC, align=LEFT, fmt='[$-411]yyyy/mm/dd"（"aaa"）"' )
    heads = ['未着手', '対応中', '完了', '期限超過']
    for i, h in enumerate(heads):
        col = get_column_letter(3 + i)  # C..F
        put(ws, f'{col}22', h, font=f(10, bold=True), fill=HEAD, align=CENTER)
    put(ws, 'B22', '件数', font=f(10, bold=True), fill=HEAD, align=CENTER)
    put(ws, 'B23', '件', font=f(10, bold=True), fill=LABEL, align=CENTER)
    for i, status in enumerate(['未着手', '対応中', '完了']):
        col = get_column_letter(3 + i)
        put(ws, f'{col}23',
            f'=COUNTIF(アクション管理!$G$2:$G${LAST},"{status}")',
            fill=CALC, align=CENTER, font=f(14, bold=True))
    put(ws, 'F23',
        f'=SUMPRODUCT((アクション管理!$D$2:$D${LAST}<>"")'
        f'*(アクション管理!$F$2:$F${LAST}<>"")'
        f'*(アクション管理!$F$2:$F${LAST}<TODAY())'
        f'*(アクション管理!$G$2:$G${LAST}<>"完了")'
        f'*(アクション管理!$G$2:$G${LAST}<>"中止"))',
        fill=CALC, align=CENTER, font=f(14, bold=True, color='C00000'))
    ws.row_dimensions[23].height = 26

    # --- 進め方 ---
    section(ws, 25, 'おすすめの進め方', last_col='F')
    steps = [
        ('会議前', '「議事録」シートをコピーし、会議名・日時・出席者とアジェンダ（2）を埋めておきます。ここまで準備できていると会議中の入力がぐっと楽になります。'),
        ('会議中', '「3. 決定事項」と「4. アクションアイテム」だけをその場で埋めます。議論の詳細は後回しで構いません。'),
        ('会議直後', '「4. アクションアイテム」の行をコピーして「アクション管理」シートに貼り付け、会議日と会議名を入れます。'),
        ('会議後', '「5. 議事内容」を清書し、「会議ログ」シートに1行追加します。共有はPDF（ファイル → エクスポート → PDF）で。'),
        ('次回まで', '「アクション管理」シートを開き、赤くなっている行から着手します。完了したらステータスを「完了」に変えます。'),
    ]
    r = 26
    for when, what in steps:
        put(ws, f'B{r}', when, font=f(10, bold=True), fill=LABEL, align=CENTER)
        ws.merge_cells(f'C{r}:F{r}')
        style_range(ws, f'C{r}:F{r}', align=LEFT)
        ws[f'C{r}'] = what
        ws.row_dimensions[r].height = 34
        r += 1

    ws.merge_cells(f'B{r+1}:F{r+1}')
    ws[f'B{r+1}'] = ('※ 行が足りなくなったら、行番号を右クリック →「挿入」で増やせます。'
                     'アクション管理シートは200行目まで数式と書式が入っています。')
    style_range(ws, f'B{r+1}:F{r+1}', align=LEFT, border=False, font=f(9, color='808080'))
    page_fit(ws, area=f'A1:F{r+1}')


# ---------------------------------------------------------------- 設定シート
def build_settings(ws):
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 46
    put(ws, 'A1', 'ステータス', font=f(10, bold=True, color='FFFFFF'), fill=BAND, align=CENTER)
    put(ws, 'B1', 'メモ', font=f(10, bold=True, color='FFFFFF'), fill=BAND, align=CENTER)
    rows = [
        ('未着手', 'まだ手をつけていない'),
        ('対応中', '着手済み'),
        ('完了', '完了（集計から除外されます）'),
        ('保留', '事情があって止めている'),
        ('中止', 'やらないことにした（集計から除外されます）'),
    ]
    for i, (s, memo) in enumerate(rows, start=2):
        put(ws, f'A{i}', s, fill=INPUT, align=CENTER)
        put(ws, f'B{i}', memo, align=LEFT)
    put(ws, 'A8', '※ この5つが「議事録」「アクション管理」シートのプルダウンの選択肢です。',
        font=f(9, color='808080'), align=LEFT, border=False)
    ws.merge_cells('A8:B8')
    page_fit(ws, area='A1:B8')


def build(path):
    wb = Workbook()
    guide = wb.active
    guide.title = '使い方'
    minutes = wb.create_sheet('議事録')
    actions = wb.create_sheet('アクション管理')
    log = wb.create_sheet('会議ログ')
    settings = wb.create_sheet('設定')

    build_minutes(minutes)
    build_actions(actions)
    build_log(log)
    build_settings(settings)
    build_guide(guide)

    wb.active = 0
    wb.save(path)
    print('saved:', path)


if __name__ == '__main__':
    build(sys.argv[1])
