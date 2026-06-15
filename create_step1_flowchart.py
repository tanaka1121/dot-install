"""
本文FMT 手順1「即出荷後に」用 縦長フロー図 生成

N8:T19(写真欄)に収まる縦長サイズで、①の判定分岐を表現する。
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

JP_FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
jp_font = fm.FontProperties(fname=JP_FONT_PATH)

PNG_FILE = "手順1_即出荷後フロー図.png"

COLOR_DECISION = "#FCE4D6"
COLOR_ACTION   = "#DDEBF7"
COLOR_WARN     = "#F8CBAD"


def box(ax, xy, w, h, text, color, fontsize=12):
    x, y = xy
    patch = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.1",
                            linewidth=1.4, edgecolor="#404040", facecolor=color)
    ax.add_patch(patch)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
            fontproperties=jp_font, fontsize=fontsize, wrap=True)


def arrow(ax, p1, p2, label=None, label_offset=(0.15, 0)):
    a = FancyArrowPatch(p1, p2, arrowstyle="-|>", mutation_scale=14,
                         linewidth=1.4, color="#404040")
    ax.add_patch(a)
    if label:
        mx = (p1[0] + p2[0]) / 2 + label_offset[0]
        my = (p1[1] + p2[1]) / 2 + label_offset[1]
        ax.text(mx, my, label, ha="center", va="center",
                fontproperties=jp_font, fontsize=11, color="#C00000",
                bbox=dict(boxstyle="round,pad=0.1", fc="white", ec="none"))


def main():
    fig, ax = plt.subplots(figsize=(4.2, 5.0))
    ax.set_xlim(0, 4.2)
    ax.set_ylim(0, 5.0)
    ax.axis("off")

    box(ax, (0.3, 4.0), 3.6, 0.8,
        "即出荷後にLPN分割を\n実施したか？", COLOR_DECISION)

    box(ax, (0.3, 2.9), 3.6, 0.7, "OLPN統合を実施する\n（手順15）", COLOR_ACTION)
    arrow(ax, (2.1, 4.0), (2.1, 3.6), label="YES", label_offset=(0.5, 0))

    box(ax, (0.3, 1.8), 3.6, 0.7, "対象はワンレックか？", COLOR_DECISION)
    arrow(ax, (2.1, 2.9), (2.1, 2.5), label="NO", label_offset=(0.5, 0))

    box(ax, (0.3, 0.85), 3.6, 0.7, "国内梱包（ワンレック）\nを実施する（手順16）", COLOR_ACTION)
    arrow(ax, (1.1, 1.8), (1.1, 1.55), label="YES", label_offset=(-0.45, 0))

    box(ax, (0.3, -0.25), 3.6, 0.9,
        "国内梱包（複数レック）\n＋イレギュラー置場へ移動\n（手順16・要周知）", COLOR_WARN, fontsize=10.5)
    arrow(ax, (3.1, 1.8), (3.1, 0.65), label="NO", label_offset=(0.45, 0))

    ax.set_ylim(-0.35, 5.0)

    plt.tight_layout()
    fig.savefig(PNG_FILE, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"作成完了: {PNG_FILE}")


if __name__ == "__main__":
    main()
