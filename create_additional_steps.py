"""
本文FMT 追加手順(15〜19) 生成スクリプト - 書き方ルール準拠版

書き方1の分岐ルール: 「○○の場合→手順【N】へ」と文章で明記
フォント: 本文24pt / 補足16pt / 急所14pt（すべてHGP創英角ゴシックUB）
色: 原則白黒（書き方1: 「原則白黒印刷」）
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

OUTPUT_FILE = "本文FMT_追加手順15-19.xlsx"

# ── フォント（書き方1: 文字→HGP創英角ゴシックUB, 数字→Arial Black）
FONT_NO   = Font(name="Arial Black",       size=24)
FONT_PROC = Font(name="HGP創英角ｺﾞｼｯｸUB", size=24)
FONT_SUPP = Font(name="HGP創英角ｺﾞｼｯｸUB", size=16)
FONT_KEY  = Font(name="HGP創英角ｺﾞｼｯｸUB", size=14)

# ── 罫線（元シートに合わせてthin/hair）
THIN = Side(style="thin")
HAIR = Side(style="hair")
B_OUTER = Border(left=THIN, right=THIN, top=THIN,  bottom=THIN)
B_INNER = Border(left=HAIR, right=HAIR, top=HAIR,  bottom=HAIR)
B_MIX   = Border(left=THIN, right=THIN, top=HAIR,  bottom=HAIR)

ALIGN_R_TOP  = Alignment(horizontal="right",  vertical="top",    wrap_text=True)
ALIGN_L_TOP  = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
ALIGN_C_MID  = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_H = 24


def s(ws, coord, val=None, font=None, align=None, border=None):
    c = ws[coord]
    if val   is not None: c.value = val
    if font  is not None: c.font  = font
    if align is not None: c.alignment = align
    if border is not None: c.border = border
    return c


def merge(ws, r1, c1, r2, c2, val=None, font=None, align=None, border=B_OUTER):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    from openpyxl.utils import get_column_letter as gcl
    coord = f"{gcl(c1)}{r1}"
    s(ws, coord, val, font, align, border)
    # 結合範囲全体に外枠を引く
    for row in range(r1, r2 + 1):
        for col in range(c1, c2 + 1):
            c = ws.cell(row=row, column=col)
            if c.border == Border():
                c.border = B_MIX


def build_step(ws, start, no, proc, supp,
               s_text="", s_reason="", q_text="", q_reason=""):
    """
    12行1ステップブロック（元フォーマット準拠）
    行レイアウト（相対オフセット）:
      0-5  : 手順No(B) / 手順(C:H) / 急所S(J:M) / 写真(N:T)
      6-11 : 補足(B:H) / 急所Q(J:M)
    """
    r, r6 = start, start + 6
    for i in range(12):
        ws.row_dimensions[r + i].height = ROW_H

    # ── 手順No.
    merge(ws, r, 2, r+5, 2, no, FONT_NO, ALIGN_R_TOP)

    # ── 手順（本文 24pt）
    merge(ws, r, 3, r+5, 8, proc, FONT_PROC, ALIGN_L_TOP)

    # ── 補足: ラベル
    merge(ws, r6, 2, r6+5, 2, "補足:", FONT_SUPP, ALIGN_R_TOP)

    # ── 補足（16pt）
    merge(ws, r6, 3, r6+5, 8, supp, FONT_SUPP, ALIGN_L_TOP)

    # ── スペーサー I列
    merge(ws, r, 9, r+11, 9, "", None, None, B_INNER)

    # ── 急所 S:
    merge(ws, r,   10, r+2,  10, "S:",    FONT_KEY, ALIGN_R_TOP, B_INNER)
    merge(ws, r,   11, r+2,  13, s_text,  FONT_KEY, ALIGN_L_TOP, B_INNER)
    merge(ws, r+3, 10, r+5,  10, "理由:", FONT_KEY, ALIGN_R_TOP, B_INNER)
    merge(ws, r+3, 11, r+5,  13, s_reason,FONT_KEY, ALIGN_L_TOP, B_INNER)

    # ── 急所 Q:
    merge(ws, r6,   10, r6+2, 10, "Q:",    FONT_KEY, ALIGN_R_TOP, B_INNER)
    merge(ws, r6,   11, r6+2, 13, q_text,  FONT_KEY, ALIGN_L_TOP, B_INNER)
    merge(ws, r6+3, 10, r6+5, 10, "理由:", FONT_KEY, ALIGN_R_TOP, B_INNER)
    merge(ws, r6+3, 11, r6+5, 13, q_reason,FONT_KEY, ALIGN_L_TOP, B_INNER)

    # ── 写真欄
    merge(ws, r, 14, r+11, 20, "", None, ALIGN_C_MID)


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "本文FMT (例外処理手順)"

    # ── 列幅（元シートに合わせる）
    widths = {"A":10.33,"B":10.16,"C":8.5,"D":8.5,"E":8.5,"F":8.5,
              "G":8.5,  "H":8.5,  "I":10.33,"J":6.33,"K":8.33,"L":8.5,
              "M":8.5,  "N":9.33}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ── 案内メモ（赤テキスト：元ファイルには含めない）
    ws["B1"] = (
        "【元ファイル 手順1 の「補足」欄(C14:H19)に追記する文章案】\n"
        "・LPN分割を実施した場合　　　　　→手順【15】へ\n"
        "・LPN分割を実施していない場合　　→手順【16】へ\n"
        "・LPN分割で数量を誤った場合　　　→手順【15】へ\n"
        "・LPN分割を過剰に行った場合　　　→手順【17】へ\n"
        "・複数部材のLPN分割を途中で中断した場合　→手順【18】へ\n"
        "・プリンタを設定しないままLPN分割した場合→手順【19】へ"
    )
    ws["B1"].font = Font(name="HGP創英角ｺﾞｼｯｸUB", size=14, color="C00000")
    ws["B1"].alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[1].height = 130
    ws.merge_cells("B1:T1")

    row = 3  # 手順15 開始

    # ── 手順15: OLPN統合（①実施済み / ②数量誤り）────────────────
    build_step(
        ws, row, "15",
        "OLPNを統合する",
        (
            "対象：\n"
            "①即出荷後にLPN分割を実施した場合\n"
            "②LPN分割で数量を誤ってしまった場合"
        ),
        s_text="統合対象のLPNを現品票で再確認する",
        s_reason="統合先LPN誤りによる出荷誤り防止",
        q_text="統合後、数量が分割前の総数と一致しているか確認する",
        q_reason="数量誤りの再発防止",
    )
    row += 12

    # ── 手順16: 国内梱包（①LPN分割未実施のNO分岐）───────────────
    build_step(
        ws, row, "16",
        "対象がワンレックか確認し、梱包を実施する",
        (
            "対象：①即出荷後にLPN分割を実施しなかった場合\n\n"
            "・ワンレックの場合\n"
            "　→国内梱包（ワンレック）を実施する\n"
            "・複数レックの場合\n"
            "　→国内梱包（複数レック）を実施し、\n"
            "　　イレギュラー置場へ移動する"
        ),
        s_text="複数レックの場合、全レック分を移動したか確認する",
        s_reason="一部レックの移動漏れによる出荷遅延防止",
        q_text="梱包指示書のレック数と現物のレック数が一致しているか確認する",
        q_reason="梱包誤り防止",
    )
    row += 12

    # ── 手順17: 分割ラベル使用（③過剰分割）──────────────────────
    build_step(
        ws, row, "17",
        "分割ラベルを使用する",
        (
            "対象：③LPN分割を過剰に行った場合\n\n"
            "余分に発行されたラベルは「分割ラベル」として\n"
            "保管し、該当LPNに使用する。\n"
            "余分なラベルは捨てずに保管箱へ入れること。"
        ),
        s_text="別LPNへの誤貼付防止のため、ラベルのLPN番号を必ず確認する",
        s_reason="誤出荷・誤集約防止",
    )
    row += 12

    # ── 手順18: 親部材集約（④中断）──────────────────────────────
    build_step(
        ws, row, "18",
        "親部材集約を実施する",
        (
            "対象：④複数部材をLPN分割する途中で中断した場合\n\n"
            "未処理の部材を親部材へ集約する。\n"
            "中断したLPN分割作業のうち、\n"
            "未処理の部材のみを対象とすること。"
        ),
        q_text="集約後、親部材の数量が分割前の総数と一致しているか確認する",
        q_reason="数量誤り防止",
    )
    row += 12

    # ── 手順19: MAラベル再印刷（⑤プリンタ未設定）───────────────
    build_step(
        ws, row, "19",
        "MAラベルを再印刷する",
        (
            "対象：⑤プリンタを設定しないままLPN分割した場合\n\n"
            "正しいプリンタを設定し、MAラベルを再印刷する。\n"
            "再印刷前に旧ラベルを必ず破棄すること。"
        ),
        s_text="再印刷前に旧ラベルを破棄し、二重貼付を防止する",
        s_reason="二重貼付による誤出荷防止",
    )

    wb.save(OUTPUT_FILE)
    print(f"作成完了: {OUTPUT_FILE}")
    print("書き方ルール:")
    print("  フォント: HGP創英角ゴシックUB (本文24pt/補足16pt/急所14pt)")
    print("  分岐: 「○○の場合→手順【N】へ」形式で文章記述")
    print("  色: 白黒（色付けなし）")


if __name__ == "__main__":
    main()
