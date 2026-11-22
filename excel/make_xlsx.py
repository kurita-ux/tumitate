"""保険の利回り比較（Excel）を作るスクリプト。

  A：今の保険（円）
  B：比べる保険（円・一時払）
  C：比べる保険（ドル建て・一時払。為替で円に直して比べる）

使い方:
    python excel/make_xlsx.py                       # 架空のサンプル入りのひな形を作る
    python excel/make_xlsx.py data.json out.xlsx    # 自分の数字（JSON）で作る

data.json の形（円は万円、ドルはドル）:
    {"a": {"name": "...", "age": 54, "year": 2026, "surrender": 213.6, "dividend": 4.2,
           "premiums": 204.7, "rows": [[55, 215.1], ...]},
     "b": {"name": "", "premium": null, "rows": []},
     "c": {"name": "", "premium_usd": null, "fx_now": 159, "fx_future": 159, "fee": 0.5, "rows": []}}

※ 公開リポジトリなので、本人の数字の JSON や出力はリポジトリに置かないこと。
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
GRAY = "4B5356"

FIRST_ROW = 16          # 表の1行目
MIN_ROWS = 30           # 表の最低行数（空き行を含む）
SPARE_ROWS = 10         # 年齢の数より多めにとる空き行
N_ROWS = MIN_ROWS       # build() の中で、入れる年齢の数に合わせて決め直す
LAST_ROW = FIRST_ROW + N_ROWS - 1
DIFF_FMT = "+#,##0.0;-#,##0.0;0.0"


def _sample_c_rows():
    prem = round(208 * 10000 / (159 + 0.5))
    rows = []
    for age in list(range(51, 61)) + [65, 70, 75, 80, 85, 90]:
        t = age - 50
        ratio = 1.035 ** t * (1 - max(0.0, 0.10 - 0.01 * t))
        rows.append([age, round(prem * ratio, 2)])
    return rows


SAMPLE = {
    "a": {
        "name": "サンプル：今の保険（架空）", "age": 50, "year": 2026,
        "surrender": 206.0, "dividend": 2.0, "premiums": 200.0,
        "rows": [[51, 207.85], [52, 209.72], [53, 211.61], [54, 213.51], [55, 215.43],
                 [56, 217.37], [57, 219.33], [58, 221.30], [59, 223.30], [60, 225.31],
                 [65, 235.61], [70, 246.38], [75, 257.65], [80, 269.43], [85, 281.75], [90, 294.63]],
    },
    "b": {
        "name": "サンプル：円建て保険（架空）", "premium": None,
        "rows": [[51, 198.30], [52, 202.10], [53, 205.90], [54, 209.80], [55, 213.70],
                 [56, 217.70], [57, 221.70], [58, 225.80], [59, 229.90], [60, 234.10],
                 [65, 256.00], [70, 279.90], [75, 306.10], [80, 334.70], [85, 366.00], [90, 400.20]],
    },
    "c": {
        "name": "サンプル：ドル建て保険（架空）", "premium_usd": None,
        "fx_now": 159, "fx_future": 159, "fee": 0.5, "rows": _sample_c_rows(),
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
    put(ws, rng.split(":")[0], text, fill=SUB_FILL)
    for row in ws[rng]:
        for c in row:
            c.border = BOX
            c.fill = SUB_FILL


def block(ws, top_ref, title, items):
    """items: (ラベル範囲, ラベル, 値セル, 値, 書式, 入力か, 式の初期値か, 結合先)"""
    put(ws, top_ref, title, bold=True, border=False, color=NAVY)
    for lab_rng, text, ref, val, fmt, is_input, fdef, merge_to in items:
        label(ws, lab_rng, text)
        put(ws, ref, val, fmt=fmt, fill=INPUT_FILL if is_input else CALC_FILL,
            color=("000000" if fdef else "0000FF") if is_input else "000000",
            align="left" if merge_to else "right")
        if merge_to:
            ws.merge_cells(f"{ref}:{merge_to}")
            ws[merge_to].border = BOX


def build(data, out_path):
    global N_ROWS, LAST_ROW
    a, b = data["a"], data["b"]
    c = data.get("c") or {}
    rows_a = {int(r[0]): r[1] for r in a["rows"]}
    rows_b = {int(r[0]): r[1] for r in b.get("rows", [])}
    rows_c = {int(r[0]): r[1] for r in c.get("rows", [])}
    ages = sorted(set(rows_a) | set(rows_b) | set(rows_c))
    N_ROWS = max(MIN_ROWS, len(ages) + SPARE_ROWS)
    LAST_ROW = FIRST_ROW + N_ROWS - 1

    wb = Workbook()
    ws = wb.active
    ws.title = "比較"
    ws.sheet_view.showGridLines = False

    widths = {"A": 7, "B": 7, "C": 8, "D": 12, "E": 11, "F": 11, "G": 12, "H": 11, "I": 11,
              "J": 13, "K": 11, "L": 12, "M": 11, "N": 13, "O": 18, "P": 13, "Q": 13}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    put(ws, "A1", "保険の利回り比較", bold=True, size=16, border=False)
    put(ws, "A2", "黄色のセルに数字を入れると、返戻率と年利回りが自動で計算されます。円の金額は万円、ドルはドル、為替は1ドルあたりの円。税金は含みません。",
        border=False, color=GRAY)

    # ---- A：今の保険 ----
    block(ws, "A4", "■ A：今の保険（円）", [
        ("A5:C5", "名前", "D5", a["name"], None, True, False, "E5"),
        ("A6:C6", "今の年齢（歳）", "D6", a["age"], "0", True, False, None),
        ("A7:C7", "今の西暦（年）", "D7", a["year"], "0", True, False, None),
        ("A8:C8", "今の解約返戻金", "D8", a["surrender"], "#,##0.00", True, False, None),
        ("A9:C9", "配当金など", "D9", a["dividend"], "#,##0.00", True, False, None),
        ("A10:C10", "今解約した受取額", "D10", "=D8+D9", "#,##0.00", False, False, None),
        ("A11:C11", "払い込んだ保険料の合計", "D11", a["premiums"], "#,##0.00", True, False, None),
        ("A12:C12", "今の返戻率（解約返戻金÷払込）", "D12", '=IF(D11>0,ROUNDDOWN(D8/D11,3),"")', "0.0%", False, False, None),
    ])

    # ---- B：円建て・一時払 ----
    prem_b = b.get("premium")
    block(ws, "F4", "■ B：比べる保険（円・一時払）", [
        ("F5:G5", "名前", "H5", b.get("name") or "", None, True, False, "I5"),
        ("F6:G6", "払い込む保険料（万円）", "H6", prem_b if prem_b is not None else "=D10", "#,##0.00", True, prem_b is None, None),
    ])
    put(ws, "F7", "最初はAを今解約した受取額が入っています。設計書の保険料に合わせて書き換えてください。",
        border=False, color=GRAY, size=9, wrap=True)
    ws.merge_cells("F7:I9")
    ws["F7"].alignment = Alignment(vertical="top", wrap_text=True)

    # ---- C：ドル建て・一時払 ----
    prem_c = c.get("premium_usd")
    block(ws, "J4", "■ C：比べる保険（ドル建て・一時払）", [
        ("J5:K5", "名前", "L5", c.get("name") or "", None, True, False, "M5"),
        ("J6:K6", "払い込む保険料（ドル）", "L6", prem_c if prem_c is not None else "=IF(L7+L9>0,ROUND(D10*10000/(L7+L9),0),0)", "#,##0", True, prem_c is None, None),
        ("J7:K7", "今の為替（円）", "L7", c.get("fx_now", 159), "#,##0.00", True, False, None),
        ("J8:K8", "受け取る時の為替（円）", "L8", c.get("fx_future", c.get("fx_now", 159)), "#,##0.00", True, False, None),
        ("J9:K9", "両替手数料（片道・円）", "L9", c.get("fee", 0.5), "#,##0.00", True, False, None),
        ("J10:K10", "払い込む保険料（円・万円）", "L10", "=L6*(L7+L9)/10000", "#,##0.00", False, False, None),
    ])
    put(ws, "J11", "払い込む保険料は、最初はAの受取額をドルに替えた額です。為替は1ドルあたりの円。",
        border=False, color=GRAY, size=9, wrap=True)
    ws.merge_cells("J11:M12")
    ws["J11"].alignment = Alignment(vertical="top", wrap_text=True)

    # ---- この年齢で比べる ----
    put(ws, "O4", "■ この年齢で比べる", bold=True, border=False, color=NAVY)
    rng = lambda col: f"${col}${FIRST_ROW}:${col}${LAST_ROW}"
    look = lambda col: f'=IFERROR(INDEX({rng(col)},MATCH($P$5,{rng("A")},0)),"表にない年齢")'
    summary = [
        ("比べたい年齢（歳）", 65, "0", True),
        ("Aの年利回り", look("F"), "0.00%", False),
        ("Bの年利回り", look("I"), "0.00%", False),
        ("Cの年利回り（円）", look("M"), "0.00%", False),
        ("いちばん有利", look("O"), None, False),
        ("Bの差（万円）", look("P"), DIFF_FMT, False),
        ("Cの差（万円）", look("Q"), DIFF_FMT, False),
        ("Cの分かれ目の為替", look("N"), '#,##0"円"', False),
    ]
    for i, (text, val, fmt, is_input) in enumerate(summary):
        r = 5 + i
        put(ws, f"O{r}", text, fill=SUB_FILL)
        put(ws, f"P{r}", val, fmt=fmt, fill=INPUT_FILL if is_input else CALC_FILL,
            color="0000FF" if is_input else "000000", align="right", bold=not is_input)

    # ---- 表の見出し ----
    groups = [("A14:C14", "時期"), ("D14:F14", "A：今の保険"), ("G14:I14", "B：円建て"),
              ("J14:N14", "C：ドル建て"), ("O14:Q14", "比較（同じ金額を入れた場合）")]
    for g, text in groups:
        ws.merge_cells(g)
        put(ws, g.split(":")[0], text, bold=True, color="FFFFFF", fill=HEAD_FILL, align="center")
        for row in ws[g]:
            for cell in row:
                cell.border = BOX
                cell.fill = HEAD_FILL
    heads = ["年齢", "西暦", "今から\n何年後",
             "解約返戻金\n（万円）", "返戻率\n（÷払込保険料）", "年利回り",
             "解約返戻金\n（万円）", "返戻率", "年利回り",
             "解約返戻金\n（ドル）", "返戻率\n（ドルで）", "円での受取額\n（万円）", "年利回り\n（円で）", "分かれ目の為替\n（これより円安でC有利）",
             "いちばん有利\n（年利回り）", "Bの差\n（B−A・万円）", "Cの差\n（C−A・万円）"]
    cols = "ABCDEFGHIJKLMNOPQ"
    for i, h in enumerate(heads):
        put(ws, f"{cols[i]}15", h, bold=True, fill=SUB_FILL, align="center", wrap=True, size=9)
    ws.row_dimensions[15].height = 48

    # ---- 表の本体 ----
    for i in range(N_ROWS):
        r = FIRST_ROW + i
        age = ages[i] if i < len(ages) else None
        put(ws, f"A{r}", age, fill=INPUT_FILL, color="0000FF", fmt="0", align="center")
        put(ws, f"B{r}", f'=IF(A{r}="","",$D$7+A{r}-$D$6)', fmt="0", align="center")
        put(ws, f"C{r}", f'=IF(A{r}="","",A{r}-$D$6)', fmt='0"年"', align="center")
        # A
        put(ws, f"D{r}", rows_a.get(age), fill=INPUT_FILL, color="0000FF", fmt="#,##0.00", align="right")
        put(ws, f"E{r}", f'=IF(OR(D{r}="",$D$11<=0),"",ROUNDDOWN(D{r}/$D$11,3))', fmt="0.0%", align="right")
        put(ws, f"F{r}", f'=IF(OR(D{r}="",C{r}="",$D$10<=0),"",IF(C{r}<=0,"",((D{r}+$D$9)/$D$10)^(1/C{r})-1))', fmt="0.00%", align="right")
        # B
        put(ws, f"G{r}", rows_b.get(age), fill=INPUT_FILL, color="0000FF", fmt="#,##0.00", align="right")
        put(ws, f"H{r}", f'=IF(OR(G{r}="",$H$6<=0),"",ROUNDDOWN(G{r}/$H$6,3))', fmt="0.0%", align="right")
        put(ws, f"I{r}", f'=IF(OR(G{r}="",C{r}="",$H$6<=0),"",IF(C{r}<=0,"",(G{r}/$H$6)^(1/C{r})-1))', fmt="0.00%", align="right")
        # C
        put(ws, f"J{r}", rows_c.get(age), fill=INPUT_FILL, color="0000FF", fmt="#,##0.00", align="right")
        put(ws, f"K{r}", f'=IF(OR(J{r}="",$L$6<=0),"",ROUNDDOWN(J{r}/$L$6,3))', fmt="0.0%", align="right")
        put(ws, f"L{r}", f'=IF(J{r}="","",J{r}*MAX(0,$L$8-$L$9)/10000)', fmt="#,##0.00", align="right")
        put(ws, f"M{r}", f'=IF(OR(L{r}="",C{r}="",$L$10<=0),"",IF(OR(C{r}<=0,L{r}<=0),"",(L{r}/$L$10)^(1/C{r})-1))', fmt="0.00%", align="right")
        put(ws, f"N{r}", f'=IF(OR(D{r}="",J{r}="",$L$6<=0,$D$10<=0),"",IF(J{r}<=0,"",$L$9+($L$7+$L$9)*((D{r}+$D$9)/$D$10)/(J{r}/$L$6)))', fmt='#,##0"円"', align="right")
        # 比較
        put(ws, f"O{r}", f'=IF(COUNT(F{r},I{r},M{r})<2,"",IF(MAX(F{r},I{r},M{r})=F{r},"A（今の保険）",IF(MAX(F{r},I{r},M{r})=I{r},"B（円建て）","C（ドル建て）")))', align="center")
        put(ws, f"P{r}", f'=IF(OR(D{r}="",G{r}="",$H$6<=0),"",G{r}/$H$6*$D$10-(D{r}+$D$9))', fmt=DIFF_FMT, align="right")
        put(ws, f"Q{r}", f'=IF(OR(D{r}="",L{r}="",$L$10<=0),"",L{r}/$L$10*$D$10-(D{r}+$D$9))', fmt=DIFF_FMT, align="right")

    # 色分け：いちばん有利
    def mark(rng_, ref, letter, bg, fg):
        ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'LEFT({ref},1)="{letter}"'], fill=PatternFill("solid", fgColor=bg), font=Font(name=FONT, color=fg, bold=True)))
    o_rng = f"O{FIRST_ROW}:O{LAST_ROW}"
    for letter, bg, fg in (("A", "E6EAF6", NAVY), ("B", "D9F2D9", "0A6B0A"), ("C", "FCE8D9", "B4541E")):
        mark(o_rng, f"O{FIRST_ROW}", letter, bg, fg)
        mark("P9", "P9", letter, bg, fg)
    # 比べたい年齢の行
    ws.conditional_formatting.add(f"A{FIRST_ROW}:Q{LAST_ROW}", FormulaRule(formula=[f'AND($A{FIRST_ROW}<>"",$A{FIRST_ROW}=$P$5)'], fill=PatternFill("solid", fgColor="FFF4CC")))

    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                        error="0以上の数字を入れてください", errorTitle="入力の確認")
    ws.add_data_validation(dv)
    for col in "DGJ":
        dv.add(f"{col}{FIRST_ROW}:{col}{LAST_ROW}")

    end = LAST_ROW + 2
    notes = [
        "※ 年利回り：今のお金（Aは今解約した受取額、B・Cは払い込む保険料）が、その年齢までに毎年何%ずつ増えたのと同じかを表します。Aは配当金なども足しています。",
        "※ C：ドルの解約返戻金を「受け取る時の為替 − 両替手数料」で円に直し、払い込みは「今の為替 ＋ 両替手数料」で計算しています。",
        "※ 分かれ目の為替：受け取る時の為替がこれより円安ならCがAより多く、円高ならAが多くなります（同じ金額を入れた場合）。",
        "※ 差：B・Cに「Aを今解約した受取額」と同じ金額を入れた場合の受取額から、Aを続けた場合の受取額を引いた額です。",
        "※ A は保険料を払い終わった保険が前提です。これから保険料を払う保険を入れると、A の年利回りが実際より高く出ます。",
        "※ 税金・手数料（両替以外）は含みません。将来の金額は各社の試算書（予想額）で、約束された額ではありません。",
    ]
    for i, t in enumerate(notes):
        put(ws, f"A{end + i}", t, border=False, color=GRAY, size=9)

    ws.freeze_panes = f"D{FIRST_ROW}"
    ws.print_title_rows = "14:15"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # ---- 使い方 ----
    hs = wb.create_sheet("使い方")
    hs.sheet_view.showGridLines = False
    hs.column_dimensions["A"].width = 4
    hs.column_dimensions["B"].width = 110
    lines = [
        ("使い方", True),
        ("1. 「比較」シートの黄色のセルが入力するところです（青い文字）。白いセルは自動で計算されます。", False),
        ("2. A（今の保険）：試算書の「今の解約返戻金」「配当金など」「払い込んだ保険料の合計」と、年齢ごとの解約返戻金を入れます。", False),
        ("3. B（円建て・一時払）：設計書の保険料を「払い込む保険料」に入れ、年齢ごとの解約返戻金（万円）を入れます。", False),
        ("4. C（ドル建て・一時払）：設計書の保険料（ドル）と、年齢ごとの解約返戻金（ドル）を入れます。", False),
        ("　 今の為替・受け取る時の為替・両替手数料を入れると、円での受取額と年利回りが出ます。", False),
        ("　 「受け取る時の為替」を 130円、190円 などに変えると、円高・円安のときの結果を試せます。", False),
        ("5. 表の「年齢」をそろえて入れます。空いている行に年齢を足せます。", False),
        ("6. 右上の「比べたい年齢」を変えると、その年齢での比較がまとめて表示され、表の行も黄色く光ります。", False),
        ("", False),
        ("計算のしかた", True),
        ("・返戻率 ＝ 解約返戻金 ÷ 払い込んだ保険料（試算書の「単純返戻率」と同じく、小数第2位を切り捨て）。Cはドルどうしで計算", False),
        ("・年利回り ＝ （受取額 ÷ 今のお金）を年数で割り戻した、1年あたりの増え方。Cは円に直してから計算", False),
        ("・C の円での受取額 ＝ 解約返戻金（ドル）×（受け取る時の為替 − 両替手数料）", False),
        ("・C の払い込む保険料（円）＝ 払い込む保険料（ドル）×（今の為替 ＋ 両替手数料）", False),
        ("・分かれ目の為替 ＝ Cの円での受取額が、Aを続けた場合と同じになる為替（同じ金額を入れた場合）", False),
        ("・差 ＝ B・Cに同じ金額を入れた場合の受取額 − Aを続けた場合の受取額（プラスならB・Cが多い）", False),
        ("", False),
        ("注意", True),
        ("・税金（一時所得など）は含みません。くわしい比較は「保険の乗り換え判定」アプリを使ってください。", False),
        ("・Aは保険料を払い終わった保険を前提にしています（払込中の保険だと、Aの年利回りが高めに出ます）。", False),
        ("・B・Cは保険料を一度に払う形（一時払）を前提にしています。", False),
        ("・為替は1つの値で計算しています。実際の為替は毎日変わり、大きく動くこともあります。", False),
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
