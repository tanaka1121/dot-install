"""
LPN分割 誤り対応フロー図 生成スクリプト

手書きメモの内容をもとに、5パターンの場合分けを1枚のフローチャート(PNG)
として作成し、Excel(xlsx)にも貼り付けて出力する。

使用方法:
    python create_lpn_flowchart.py
    → LPN分割_誤り対応フロー図.png
    → LPN分割_誤り対応フロー図.xlsx
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

# ── 日本語フォント設定 ──────────────────────────
JP_FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
jp_font = fm.FontProperties(fname=JP_FONT_PATH)

PNG_FILE = "LPN分割_誤り対応フロー図.png"
XLSX_FILE = "LPN分割_誤り対応フロー図.xlsx"

# ── 色定義 ──────────────────────────────────────
COLOR_START    = "#FFF2CC"  # 黄: 発生事象（開始）
COLOR_DECISION = "#FCE4D6"  # 橙: 判定
COLOR_ACTION   = "#DDEBF7"  # 青: 対応アクション
COLOR_WARN     = "#F8CBAD"  # 赤橙: 注記


def box(ax, xy, w, h, text, color, fontsize=11, shape="rect"):
    x, y = xy
    if shape == "diamond":
        boxstyle = "round,pad=0.02,rounding_size=0.15"
    else:
        boxstyle = "round,pad=0.02,rounding_size=0.08"
    patch = FancyBboxPatch(
        (x, y), w, h,
        boxstyle=boxstyle,
        linewidth=1.3,
        edgecolor="#404040",
        facecolor=color,
    )
    ax.add_patch(patch)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
            fontproperties=jp_font, fontsize=fontsize, wrap=True)
    return (x, y, w, h)


def arrow(ax, p1, p2, label=None, label_pos=0.5, color="#404040", rad=0.0):
    connectionstyle = f"arc3,rad={rad}" if rad else "arc3"
    a = FancyArrowPatch(p1, p2, arrowstyle="-|>", mutation_scale=14,
                         linewidth=1.3, color=color,
                         connectionstyle=connectionstyle)
    ax.add_patch(a)
    if label:
        lx = p1[0] + (p2[0] - p1[0]) * label_pos
        ly = p1[1] + (p2[1] - p1[1]) * label_pos
        if rad:
            ly += rad * 1.5
        ax.text(lx, ly, label, ha="center", va="center",
                fontproperties=jp_font, fontsize=10, color="#C00000",
                bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none"))


def main():
    fig, ax = plt.subplots(figsize=(15, 11))
    ax.set_xlim(0, 15)
    ax.set_ylim(0, 11)
    ax.axis("off")

    fig.suptitle("LPN分割 誤り対応フロー（重長品検数作業 例外処理）",
                  fontproperties=jp_font, fontsize=16, y=0.985)

    # ── ① の流れ ─────────────────────────────────
    b1 = box(ax, (0.3, 8.6), 3.0, 1.2,
             "①LPN分割を忘れて\n即出荷してしまった", COLOR_START)
    b2 = box(ax, (3.9, 8.6), 3.2, 1.2,
             "即出荷後に\nLPN分割したか？", COLOR_DECISION, shape="diamond")
    b3 = box(ax, (8.2, 9.3), 2.6, 1.0, "OLPN統合", COLOR_ACTION)
    b4 = box(ax, (8.2, 6.5), 3.2, 1.2,
             "ワンレックか？\n複数レックか？", COLOR_DECISION, shape="diamond")
    b5 = box(ax, (12.0, 7.4), 2.7, 1.0, "国内梱包\n（ワンレック）", COLOR_ACTION)
    b6 = box(ax, (12.0, 5.6), 2.7, 1.6,
             "国内梱包（複数レック）\n＋イレギュラー置場\n（相当分の周知が必要）", COLOR_WARN, fontsize=9.5)

    arrow(ax, (3.3, 9.2), (3.9, 9.2))
    arrow(ax, (7.1, 9.4), (8.2, 9.7), label="YES", label_pos=0.5)
    arrow(ax, (5.5, 8.6), (5.5, 7.7))
    arrow(ax, (5.5, 7.7), (8.2, 7.3), label="NO")
    arrow(ax, (11.4, 7.5), (12.0, 7.8), label="ワンレック")
    arrow(ax, (9.8, 6.5), (12.0, 6.3), label="複数レック")

    # ── ② の流れ ─────────────────────────────────
    b7 = box(ax, (0.3, 6.6), 3.0, 1.2,
             "②LPN分割で数量を\n誤ってしまった", COLOR_START)
    arrow(ax, (1.8, 8.6), (1.8, 7.8))
    arrow(ax, (3.3, 6.9), (8.2, 9.55), label="OLPN統合へ", label_pos=0.5, rad=-0.35)

    # ── ③ の流れ ─────────────────────────────────
    b8 = box(ax, (0.3, 4.4), 3.0, 1.2,
             "③LPN分割を\n過剰に行った", COLOR_START)
    b9 = box(ax, (3.9, 4.4), 3.0, 1.2, "分割ラベルを\n使用する", COLOR_ACTION)
    arrow(ax, (3.3, 5.0), (3.9, 5.0))

    # ── ④ の流れ ─────────────────────────────────
    b10 = box(ax, (0.3, 2.2), 3.0, 1.4,
              "④複数部材をLPN分割する\n途中で中断した", COLOR_START)
    b11 = box(ax, (3.9, 2.4), 3.0, 1.0, "親部材集約", COLOR_ACTION)
    arrow(ax, (3.3, 2.9), (3.9, 2.9))

    # ── ⑤ の流れ ─────────────────────────────────
    b12 = box(ax, (0.3, 0.3), 3.4, 1.4,
              "⑤プリンタを設定しない\nままLPN分割した", COLOR_START)
    b13 = box(ax, (3.9, 0.5), 3.0, 1.0, "MAラベルを\n再印刷する", COLOR_ACTION)
    arrow(ax, (3.7, 1.0), (3.9, 1.0))

    # ── 凡例 ─────────────────────────────────────
    legend_y = 0.0
    box(ax, (8.2, 3.4), 1.5, 0.6, "発生事象", COLOR_START, fontsize=9)
    box(ax, (9.9, 3.4), 1.5, 0.6, "判定", COLOR_DECISION, fontsize=9)
    box(ax, (11.6, 3.4), 1.5, 0.6, "対応", COLOR_ACTION, fontsize=9)
    box(ax, (13.3, 3.4), 1.5, 0.6, "注記", COLOR_WARN, fontsize=9)

    plt.tight_layout()
    fig.savefig(PNG_FILE, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"作成完了: {PNG_FILE}")

    # ── Excelに貼り付け ───────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = "誤り対応フロー図"
    img = XLImage(PNG_FILE)
    # 画像が大きすぎる場合は縮小
    scale = 1400 / img.width
    img.width = int(img.width * scale)
    img.height = int(img.height * scale)
    ws.add_image(img, "A1")
    wb.save(XLSX_FILE)
    print(f"作成完了: {XLSX_FILE}")


if __name__ == "__main__":
    main()
