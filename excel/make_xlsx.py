"""保険の利回り比較（Excel）を作るスクリプト。

使い方:
    python excel/make_xlsx.py                       # 架空のサンプル入りのひな形を作る
    python excel/make_xlsx.py data.json out.xlsx    # 自分の数字（JSON）で作る

data.json の形（金額は万円）:
    {"a": {"name": "...", "age": 54, "year": 2026, "surrender": 213.62, "dividend": 4.2,
           "premiums": 204.68, "rows": [[55, 215.1], [56, 217.1], ...]},
     "b": {"name": "", "premium": null, "rows": []}}

※ 公開リポジトリなので、本人の数字の JSON はリポジトリに置かないこと。
"""
import json
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "游ゴシック"
NAVY = "2D469C"
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
SUB_FILL = PatternFill("solid", fgColor="E6EAF6")
CALC_FILL = PatternFill("solid", fgColor="F4F6F5")
THIN = Side(style="thin", color="C3C8C5")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FIRST_ROW = 15          # 表の1行目
N_ROWS = 30             # 表の行数（空き行を含む）
LAST_ROW = FIRST_ROW + N_ROWS - 1

SAMPLE = {
    "a": {
        "name": "サンプル：今の保険（架空）", "age": 50, "year": 2026,
        "surrender": 206.0, "dividend": 2.0, "premiums": 200.0,
        "rows": [[51, 207.85], [52, 209.72], [53, 211.61], [54, 213.51], [55, 215.43],
                 [56, 217.37], [57, 219.33], [58, 221.30], [59, 223.30], [60, 225.31],
                 [65, 235.61], [70, 246.38], [75, 257.65], [80, 269.43], [85, 281.75], [90, 294.63]],
    },
    "b": {
        "name": "サンプル：比べる保険（架空）", "premium": None,
        "rows": [[51, 198.30], [52, 202.10], [53, 205.90], [54, 209.80], [55, 213.70],
                 [56, 217.70], [57, 221.70], [58, 225.80], [59, 229.90], [60, 234.10],
                 [65, 256.00], [70, 279.90], [75, 306.10], [80, 334.70], [85, 366.00], [90, 400.20]],
    },
}


def font(**kw):
    kw.setdefault("name", FONT)
    kw.setdefault("size", 10.5)
    return Font(**kw)


def put(ws, ref, value, *, bold=False, color="000000", fill=None, fmt=None, align=None, size=10.5, border=True, wrap=False):
    c = ws[ref]
    c.value = value
    c.font = font(bold=bold, color=color, size=size)
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        c.border = BOX
    return c


def label(ws, rng, text):
    ws.merge_cells(rng)
    first = rng.split(":")[0]
    put(ws, first, text, fill=SUB_FILL)
    for row in ws[rng]:
        for c in row:
            c.border = BOX
            c.fill = SUB_FILL


def build(data, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "比較"
    ws.sheet_view.showGridLines = False

    widths = {"A": 8, "B": 8, "C": 9, "D": 14, "E": 12, "F": 12, "G": 14, "H": 12, "I": 12, "J": 17, "K": 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    put(ws, "A1", "保険の利回り比較", bold=True, size=16, border=False)
    put(ws, "A2", "黄色のセルに数字を入れると、返戻率と年利回りが自動で計算されます。金額の単位は万円、税金は含みません。",
        border=False, color="4B5356")

    a, b = data["a"], data["b"]

    # ---- 今の保険（A） ----
    put(ws, "A4", "■ 今の保険（A）", bold=True, border=False, color=NAVY)
    rows_a = [
        ("名前", "D5", a["name"], None, True),
        ("今の年齢（歳）", "D6", a["age"], "0", True),
        ("今の西暦（年）", "D7", a["year"], "0", True),
        ("今の解約返戻金", "D8", a["surrender"], "#,##0.00", True),
        ("配当金など", "D9", a["dividend"], "#,##0.00", True),
        ("今解約した受取額", "D10", "=D8+D9", "#,##0.00", False),
        ("払い込んだ保険料の合計", "D11", a["premiums"], "#,##0.00", True),
        ("今の返戻率", "D12", '=IF(D11>0,D10/D11,"")', "0.0%", False),
    ]
    for i, (text, ref, val, fmt, is_input) in enumerate(rows_a):
        r = 5 + i
        label(ws, f"A{r}:C{r}", text)
        put(ws, ref, val, fmt=fmt, fill=INPUT_FILL if is_input else CALC_FILL,
            color="0000FF" if is_input else "000000", align="left" if ref == "D5" else "right")
        if ref == "D5":
            ws.merge_cells("D5:E5")
            ws["E5"].border = BOX

    # ---- 比べる保険（B） ----
    put(ws, "F4", "■ 比べる保険（B）", bold=True, border=False, color=NAVY)
    label(ws, "F5:G5", "名前")
    put(ws, "H5", b.get("name") or "", fill=INPUT_FILL, color="0000FF", align="left")
    ws.merge_cells("H5:I5")
    ws["I5"].border = BOX
    label(ws, "F6:G6", "払い込む保険料（万円）")
    prem = b.get("premium")
    put(ws, "H6", prem if prem is not None else "=D10", fill=INPUT_FILL,
        color="0000FF" if prem is not None else "000000", fmt="#,##0.00", align="right")
    ws["H6"].comment = None
    put(ws, "F7", "最初はAを今解約した受取額が入っています。比べる保険の設計書の保険料に合わせて書き換えてください。",
        border=False, color="4B5356", size=9, wrap=True)
    ws.merge_cells("F7:H9")
    ws["F7"].alignment = Alignment(vertical="top", wrap_text=True)

    # ---- 比べたい年齢 ----
    put(ws, "J4", "■ この年齢で比べる", bold=True, border=False, color=NAVY)
    summary = [
        ("比べたい年齢（歳）", 65, "0", True),
        ("Aの年利回り", f'=IFERROR(INDEX($F${FIRST_ROW}:$F${LAST_ROW},MATCH($K$5,$A${FIRST_ROW}:$A${LAST_ROW},0)),"表にない年齢")', "0.00%", False),
        ("Bの年利回り", f'=IFERROR(INDEX($I${FIRST_ROW}:$I${LAST_ROW},MATCH($K$5,$A${FIRST_ROW}:$A${LAST_ROW},0)),"表にない年齢")', "0.00%", False),
        ("Aの返戻率", f'=IFERROR(INDEX($E${FIRST_ROW}:$E${LAST_ROW},MATCH($K$5,$A${FIRST_ROW}:$A${LAST_ROW},0)),"表にない年齢")', "0.0%", False),
        ("Bの返戻率", f'=IFERROR(INDEX($H${FIRST_ROW}:$H${LAST_ROW},MATCH($K$5,$A${FIRST_ROW}:$A${LAST_ROW},0)),"表にない年齢")', "0.0%", False),
        ("有利なほう", f'=IFERROR(INDEX($J${FIRST_ROW}:$J${LAST_ROW},MATCH($K$5,$A${FIRST_ROW}:$A${LAST_ROW},0)),"表にない年齢")', None, False),
        ("差（B−A・万円）", f'=IFERROR(INDEX($K${FIRST_ROW}:$K${LAST_ROW},MATCH($K$5,$A${FIRST_ROW}:$A${LAST_ROW},0)),"表にない年齢")', "+#,##0.0;-#,##0.0;0.0", False),
    ]
    for i, (text, val, fmt, is_input) in enumerate(summary):
        r = 5 + i
        put(ws, f"J{r}", text, fill=SUB_FILL)
        put(ws, f"K{r}", val, fmt=fmt, fill=INPUT_FILL if is_input else CALC_FILL,
            color="0000FF" if is_input else "000000", align="right", bold=not is_input)

    # ---- 表の見出し ----
    ws.merge_cells("A13:C13")
    put(ws, "A13", "時期", bold=True, color="FFFFFF", fill=HEAD_FILL, align="center")
    ws.merge_cells("D13:F13")
    put(ws, "D13", "今の保険（A）", bold=True, color="FFFFFF", fill=HEAD_FILL, align="center")
    ws.merge_cells("G13:I13")
    put(ws, "G13", "比べる保険（B）", bold=True, color="FFFFFF", fill=HEAD_FILL, align="center")
    ws.merge_cells("J13:K13")
    put(ws, "J13", "比較", bold=True, color="FFFFFF", fill=HEAD_FILL, align="center")
    for rng in ("A13:C13", "D13:F13", "G13:I13", "J13:K13"):
        for row in ws[rng]:
            for c in row:
                c.border = BOX
                c.fill = HEAD_FILL
    heads = ["年齢", "西暦", "今から\n何年後", "解約返戻金\n（万円）", "返戻率\n（÷払込保険料）", "年利回り\n（今の受取額から）",
             "解約返戻金\n（万円）", "返戻率\n（÷保険料）", "年利回り", "有利なほう\n（年利回り）", "差：同じ金額を\n入れた場合（万円）"]
    for i, h in enumerate(heads):
        put(ws, f"{'ABCDEFGHIJK'[i]}14", h, bold=True, fill=SUB_FILL, align="center", wrap=True, size=9.5)
    ws.row_dimensions[14].height = 46

    # ---- 表の本体 ----
    rows_a = {int(r[0]): r[1] for r in a["rows"]}
    rows_b = {int(r[0]): r[1] for r in b.get("rows", [])}
    ages = sorted(set(rows_a) | set(rows_b))
    for i in range(N_ROWS):
        r = FIRST_ROW + i
        age = ages[i] if i < len(ages) else None
        put(ws, f"A{r}", age, fill=INPUT_FILL, color="0000FF", fmt="0", align="center")
        put(ws, f"B{r}", f'=IF(A{r}="","",$D$7+A{r}-$D$6)', fmt="0", align="center")
        put(ws, f"C{r}", f'=IF(A{r}="","",A{r}-$D$6)', fmt='0"年"', align="center")
        put(ws, f"D{r}", rows_a.get(age) if age is not None else None, fill=INPUT_FILL, color="0000FF", fmt="#,##0.00", align="right")
        put(ws, f"E{r}", f'=IF(OR(D{r}="",$D$11<=0),"",ROUNDDOWN(D{r}/$D$11,3))', fmt="0.0%", align="right")
        put(ws, f"F{r}", f'=IF(OR(D{r}="",C{r}="",$D$10<=0),"",IF(C{r}<=0,"",((D{r}+$D$9)/$D$10)^(1/C{r})-1))', fmt="0.00%", align="right")
        put(ws, f"G{r}", rows_b.get(age) if age is not None else None, fill=INPUT_FILL, color="0000FF", fmt="#,##0.00", align="right")
        put(ws, f"H{r}", f'=IF(OR(G{r}="",$H$6<=0),"",ROUNDDOWN(G{r}/$H$6,3))', fmt="0.0%", align="right")
        put(ws, f"I{r}", f'=IF(OR(G{r}="",C{r}="",$H$6<=0),"",IF(C{r}<=0,"",(G{r}/$H$6)^(1/C{r})-1))', fmt="0.00%", align="right")
        put(ws, f"J{r}", f'=IF(OR(F{r}="",I{r}=""),"",IF(I{r}>F{r},"B（比べる保険）",IF(I{r}<F{r},"A（今の保険）","同じ")))', align="center")
        put(ws, f"K{r}", f'=IF(OR(D{r}="",G{r}="",$H$6<=0),"",G{r}/$H$6*$D$10-(D{r}+$D$9))', fmt="+#,##0.0;-#,##0.0;0.0", align="right")

    # 有利なほうに色
    rng = f"J{FIRST_ROW}:J{LAST_ROW}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(J{FIRST_ROW},1)="B"'], fill=PatternFill("solid", fgColor="D9F2D9"), font=Font(name=FONT, color="0A6B0A", bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(J{FIRST_ROW},1)="A"'], fill=PatternFill("solid", fgColor="E6EAF6"), font=Font(name=FONT, color=NAVY, bold=True)))
    ws.conditional_formatting.add("K10", FormulaRule(formula=['LEFT(K10,1)="B"'], fill=PatternFill("solid", fgColor="D9F2D9"), font=Font(name=FONT, color="0A6B0A", bold=True)))
    # 比べたい年齢の行に色
    ws.conditional_formatting.add(f"A{FIRST_ROW}:K{LAST_ROW}", FormulaRule(formula=[f'AND($A{FIRST_ROW}<>"",$A{FIRST_ROW}=$K$5)'], fill=PatternFill("solid", fgColor="FFF4CC")))

    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                        error="0以上の数字を入れてください", errorTitle="入力の確認")
    ws.add_data_validation(dv)
    dv.add(f"D{FIRST_ROW}:D{LAST_ROW}")
    dv.add(f"G{FIRST_ROW}:G{LAST_ROW}")

    end = LAST_ROW + 2
    put(ws, f"A{end}", "※ 年利回り：今の受取額（Bは払い込む保険料）が、その年齢までに毎年何%ずつ増えたのと同じかを表します。Aは配当金なども足しています。",
        border=False, color="4B5356", size=9)
    put(ws, f"A{end + 1}", "※ 差：Bに「Aを今解約した受取額」と同じ金額を入れた場合の受取額から、Aを続けた場合の受取額を引いた額です。",
        border=False, color="4B5356", size=9)
    put(ws, f"A{end + 2}", "※ 税金・手数料・為替は含みません。将来の金額は各社の試算書（予想額）で、約束された額ではありません。",
        border=False, color="4B5356", size=9)
    ws.freeze_panes = f"A{FIRST_ROW}"
    ws.print_title_rows = "13:14"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ---- 使い方 ----
    hs = wb.create_sheet("使い方")
    hs.sheet_view.showGridLines = False
    hs.column_dimensions["A"].width = 4
    hs.column_dimensions["B"].width = 100
    lines = [
        ("使い方", True),
        ("1. 「比較」シートの黄色のセルが入力するところです（青い文字）。白いセルは自動で計算されます。", False),
        ("2. 今の保険（A）：試算書の「今の解約返戻金」「配当金など」「払い込んだ保険料の合計」を入れます。", False),
        ("3. 比べる保険（B）：設計書の保険料を「払い込む保険料」に入れます（最初はAを解約した額が入っています）。", False),
        ("4. 表の「年齢」をそろえて、AとBの解約返戻金を入れると、返戻率と年利回りが出ます。空いている行に年齢を足せます。", False),
        ("5. 右上の「比べたい年齢」を変えると、その年齢での比較がまとめて表示されます。表の行も黄色く光ります。", False),
        ("", False),
        ("計算のしかた", True),
        ("・返戻率 ＝ 解約返戻金 ÷ 払い込んだ保険料（試算書の「単純返戻率」と同じく、小数第2位を切り捨て）", False),
        ("・年利回り ＝ （受取額 ÷ 今のお金）を年数で割り戻した、1年あたりの増え方", False),
        ("　 A：（解約返戻金＋配当金など）÷ 今解約した受取額　／　B：解約返戻金 ÷ 払い込む保険料", False),
        ("・差 ＝ Bに同じ金額を入れた場合の受取額 − Aを続けた場合の受取額（プラスならBが多い）", False),
        ("", False),
        ("注意", True),
        ("・税金（一時所得など）、為替、手数料は含みません。くわしい比較は「保険の乗り換え判定」アプリを使ってください。", False),
        ("・Bは保険料を一度に払う形（一時払）を前提にしています。毎月払う保険の年利回りは正しく出ません。", False),
        ("・将来の解約返戻金は各社の試算（予想額）です。解約すると死亡保障もなくなる点に注意してください。", False),
    ]
    for i, (t, bold) in enumerate(lines):
        put(hs, f"B{i + 2}", t, bold=bold, border=False, size=12 if bold else 10.5, color=NAVY if bold else "000000")

    wb.save(out_path)


if __name__ == "__main__":
    if len(sys.argv) >= 3:
        with open(sys.argv[1], encoding="utf-8") as f:
            data = json.load(f)
        build(data, sys.argv[2])
    else:
        build(SAMPLE, "excel/hoken-rimawari-hikaku-sample.xlsx")
    print("ok")
