"""
本文FMT 追加手順(15〜19) 生成スクリプト

①〜⑤の場合分けに対応する追加手順を、元の「本文FMT」シートと同じ
12行1ステップの書式（手順No./手順/急所/写真）で作成する。
コピー＆ペーストで元ファイルの手順14の後ろ(203行目以降)に追加できる。

使用方法:
    python create_additional_steps.py
    → 本文FMT_追加手順15-19.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as XLImage

OUTPUT_FILE = "本文FMT_追加手順15-19.xlsx"

FONT_NO     = Font(name="Arial Black", size=24)
FONT_PROC   = Font(name="HGP創英角ｺﾞｼｯｸUB", size=20)
FONT_SUPP_L = Font(name="HGP創英角ｺﾞｼｯｸUB", size=16)
FONT_SUPP   = Font(name="HGP創英角ｺﾞｼｯｸUB", size=14)
FONT_KEY    = Font(name="HGP創英角ｺﾞｼｯｸUB", size=14)

THIN  = Side(style="thin", color="000000")
HAIR  = Side(style="hair", color="808080")

BORDER_OUTER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_INNER = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)

ALIGN_NO   = Alignment(horizontal="right", vertical="top")
ALIGN_PROC = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_KEY_LABEL = Alignment(horizontal="right", vertical="top", wrap_text=True)
ALIGN_KEY_TEXT  = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

FILL_CASE = {
    "12": "DDEBF7",  # ①YES・②共通 → OLPN統合
    "13": "DDEBF7",  # ①NO → 国内梱包
    "14": "E4DFEC",  # ③ → 分割ラベル
    "15": "FCE4D6",  # ④ → 親部材集約
    "16": "D9D9D9",  # ⑤ → MAラベル再印刷
}

ROW_HEIGHT = 24


def set_cell(ws, coord, value=None, font=None, align=None, border=None, fill=None):
    c = ws[coord]
    if value is not None:
        c.value = value
    if font:
        c.font = font
    if align:
        c.alignment = align
    if border:
        c.border = border
    if fill:
        c.fill = PatternFill(start_color=fill, end_color=fill, fill_type="solid")
    return c


def build_step(ws, start_row, step_no, proc_text, supp_text,
                s_text="", s_reason="", q_text="", q_reason="",
                fill=None, image_path=None):
    """1ステップ(12行)を作成する"""
    r0 = start_row  # 上段ブロック開始行
    r6 = start_row + 6  # 下段ブロック開始行

    for r in range(start_row, start_row + 12):
        ws.row_dimensions[r].height = ROW_HEIGHT

    # ── 手順No. ───────────────────────────────
    ws.merge_cells(start_row=r0, start_column=2, end_row=r0 + 5, end_column=2)
    set_cell(ws, f"B{r0}", step_no, font=FONT_NO, align=ALIGN_NO,
             border=BORDER_OUTER, fill=fill)

    # ── 手順 ─────────────────────────────────
    ws.merge_cells(start_row=r0, start_column=3, end_row=r0 + 5, end_column=8)
    set_cell(ws, f"C{r0}", proc_text, font=FONT_PROC, align=ALIGN_PROC,
             border=BORDER_OUTER, fill=fill)

    # ── 補足 ─────────────────────────────────
    ws.merge_cells(start_row=r6, start_column=2, end_row=r6 + 5, end_column=2)
    set_cell(ws, f"B{r6}", "補足:", font=FONT_SUPP_L, align=ALIGN_NO,
             border=BORDER_OUTER, fill=fill)

    ws.merge_cells(start_row=r6, start_column=3, end_row=r6 + 5, end_column=8)
    set_cell(ws, f"C{r6}", supp_text, font=FONT_SUPP, align=ALIGN_PROC,
             border=BORDER_OUTER, fill=fill)

    # ── スペーサー列 I ────────────────────────
    ws.merge_cells(start_row=r0, start_column=9, end_row=r0 + 11, end_column=9)
    set_cell(ws, f"I{r0}", "", border=BORDER_INNER)

    # ── 急所(S/理由/Q/理由) ─────────────────────
    # S:
    ws.merge_cells(start_row=r0, start_column=10, end_row=r0 + 2, end_column=10)
    set_cell(ws, f"J{r0}", "S:", font=FONT_KEY, align=ALIGN_KEY_LABEL, border=BORDER_INNER)
    ws.merge_cells(start_row=r0, start_column=11, end_row=r0 + 2, end_column=13)
    set_cell(ws, f"K{r0}", s_text, font=FONT_KEY, align=ALIGN_KEY_TEXT, border=BORDER_INNER)

    # 理由(S):
    ws.merge_cells(start_row=r0 + 3, start_column=10, end_row=r0 + 5, end_column=10)
    set_cell(ws, f"J{r0+3}", "理由:", font=FONT_KEY, align=ALIGN_KEY_LABEL, border=BORDER_INNER)
    ws.merge_cells(start_row=r0 + 3, start_column=11, end_row=r0 + 5, end_column=13)
    set_cell(ws, f"K{r0+3}", s_reason, font=FONT_KEY, align=ALIGN_KEY_TEXT, border=BORDER_INNER)

    # Q:
    ws.merge_cells(start_row=r6, start_column=10, end_row=r6 + 2, end_column=10)
    set_cell(ws, f"J{r6}", "Q:", font=FONT_KEY, align=ALIGN_KEY_LABEL, border=BORDER_INNER)
    ws.merge_cells(start_row=r6, start_column=11, end_row=r6 + 2, end_column=13)
    set_cell(ws, f"K{r6}", q_text, font=FONT_KEY, align=ALIGN_KEY_TEXT, border=BORDER_INNER)

    # 理由(Q):
    ws.merge_cells(start_row=r6 + 3, start_column=10, end_row=r6 + 5, end_column=10)
    set_cell(ws, f"J{r6+3}", "理由:", font=FONT_KEY, align=ALIGN_KEY_LABEL, border=BORDER_INNER)
    ws.merge_cells(start_row=r6 + 3, start_column=11, end_row=r6 + 5, end_column=13)
    set_cell(ws, f"K{r6+3}", q_reason, font=FONT_KEY, align=ALIGN_KEY_TEXT, border=BORDER_INNER)

    # ── 写真 ─────────────────────────────────
    ws.merge_cells(start_row=r0, start_column=14, end_row=r0 + 11, end_column=20)
    set_cell(ws, f"N{r0}", "", border=BORDER_OUTER, align=ALIGN_CENTER)

    if image_path:
        img = XLImage(image_path)
        scale = 380 / img.width
        img.width = int(img.width * scale)
        img.height = int(img.height * scale)
        ws.add_image(img, f"N{r0}")


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "本文FMT (追加手順)"

    # ── 列幅(本文FMTに合わせる) ────────────────
    widths = {"A": 10.33, "B": 10.16, "C": 8.5, "D": 8.5, "E": 8.5, "F": 8.5,
              "G": 8.5, "H": 8.5, "I": 10.33, "J": 6.33, "K": 8.33, "L": 8.5,
              "M": 8.5, "N": 9.33}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # 案内行
    ws["B1"] = "※元ファイル「本文FMT」シートの手順14の下(203行目以降)に、この行ごとコピー＆貼り付けしてください"
    ws["B1"].font = Font(name="游ゴシック", size=10, bold=True, color="C00000")
    ws.row_dimensions[1].height = 20

    row = 3  # 手順15開始行

    # ── 手順15: OLPN統合（①YES・②共通） ─────────
    build_step(
        ws, row, "15",
        "OLPN統合を実施する\n"
        "（①即出荷後にLPN分割した場合のYES分岐／②LPN分割で数量を誤った場合に対応）",
        "【場合分け対応表】\n"
        "①即出荷後にLPN分割を忘れた → 手順15・16\n"
        "②LPN分割で数量を誤った　　　 → 手順15\n"
        "③LPN分割を過剰に行った　　　 → 手順17\n"
        "④複数部材のLPN分割を中断した → 手順18\n"
        "⑤プリンタ未設定のままLPN分割 → 手順19",
        s_text="統合対象のLPNを誤らないよう、現品票で再確認する",
        s_reason="統合先LPNの誤りによる出荷誤り防止",
        q_text="統合後、数量が分割前の総数と一致しているか確認する",
        q_reason="数量誤りの再発防止",
        fill=FILL_CASE["12"],
        image_path="手順1_即出荷後フロー図.png",
    )
    row += 12

    # ── 手順16: 国内梱包（①NO分岐） ──────────────
    build_step(
        ws, row, "16",
        "対象がワンレックか確認し、梱包を実施する\n"
        "（①即出荷後にLPN分割していない場合のNO分岐）\n\n"
        "・ワンレック　→ 国内梱包（ワンレック）を実施する\n"
        "・複数レック　→ 国内梱包（複数レック）を実施し、"
        "イレギュラー置場へ移動する",
        "複数レックの場合はイレギュラー置場への移動が必要。"
        "対応漏れがないよう関係者へ必ず周知すること。",
        s_text="複数レックの場合、すべてのレックを移動したか確認する",
        s_reason="一部レックの移動漏れによる出荷遅延防止",
        q_text="梱包指示書のレック数と現物のレック数が一致しているか確認する",
        q_reason="梱包誤り防止",
        fill=FILL_CASE["13"],
    )
    row += 12

    # ── 手順17: ③分割ラベル使用 ───────────────────
    build_step(
        ws, row, "17",
        "【③対応】分割ラベルを使用する\n\n"
        "LPN分割を過剰に行った場合、余分に発行されたラベルは"
        "「分割ラベル」として保管し、該当LPNに使用する。",
        "余分に発行したラベルは捨てずに保管箱へ入れること。",
        s_text="誤って別LPNに分割ラベルを貼付しないよう、"
               "ラベルのLPN番号を必ず確認する",
        s_reason="誤出荷・誤集約防止",
        fill=FILL_CASE["14"],
    )
    row += 12

    # ── 手順18: ④親部材集約 ───────────────────────
    build_step(
        ws, row, "18",
        "【④対応】親部材集約を実施する\n\n"
        "複数部材のLPN分割作業を中断した場合、未処理の部材を"
        "親部材へ集約する。",
        "中断したLPN分割作業のうち、未処理の部材のみを対象とする。",
        q_text="集約後、親部材の数量が分割前の総数と一致しているか確認する",
        q_reason="数量誤り防止",
        fill=FILL_CASE["15"],
    )
    row += 12

    # ── 手順19: ⑤MAラベル再印刷 ───────────────────
    build_step(
        ws, row, "19",
        "【⑤対応】MAラベルを再印刷する\n\n"
        "プリンタを設定しないままLPN分割を行った場合、正しい"
        "プリンタを設定し、MAラベルを再印刷する。",
        "再印刷前に旧ラベルを必ず破棄すること。",
        s_text="再印刷前に旧ラベルを破棄し、二重貼付を防止する",
        s_reason="二重貼付による誤出荷防止",
        fill=FILL_CASE["16"],
    )

    wb.save(OUTPUT_FILE)
    print(f"作成完了: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
