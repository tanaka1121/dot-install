"""
LPN分割 誤り対応 フローチャート（標準記号版）
  ○ : 開始 / 終了
  ◇ : 分岐（判断）
  □ : 処理（アクション）
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
import numpy as np

JP_FONT_PATH = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
FP = fm.FontProperties(fname=JP_FONT_PATH)

PNG_FILE  = "LPN分割_誤り対応フロー図.png"
XLSX_FILE = "LPN分割_誤り対応フロー図.xlsx"

# ── 色 ──────────────────────────────────────────
C_START  = "white"
C_CASE   = "white"
C_DEC    = "#FFF2CC"
C_ACT    = "#DDEBF7"
C_EDGE   = "black"
C_ARROW  = "black"
C_YES    = "#007700"
C_NO     = "#CC0000"


def txt(ax, x, y, s, size=11, bold=False, color="black", ha="center", va="center"):
    weight = "bold" if bold else "normal"
    ax.text(x, y, s, ha=ha, va=va,
            fontproperties=FP, fontsize=size, fontweight=weight, color=color)


# ── 図形描画 ─────────────────────────────────────
def oval(ax, cx, cy, rw, rh, text, fsize=10.5, fc=C_START):
    """開始／終了: 楕円○"""
    e = mpatches.Ellipse((cx, cy), rw * 2, rh * 2,
                          facecolor=fc, edgecolor=C_EDGE, linewidth=1.5, zorder=3)
    ax.add_patch(e)
    txt(ax, cx, cy, text, size=fsize, bold=True)
    return dict(cx=cx, cy=cy, rw=rw, rh=rh, kind="oval")


def rect(ax, cx, cy, hw, hh, text, fsize=10.5, fc=C_ACT):
    """処理: 四角□"""
    p = mpatches.FancyBboxPatch((cx - hw, cy - hh), hw * 2, hh * 2,
                                  boxstyle="round,pad=0.02,rounding_size=0.06",
                                  facecolor=fc, edgecolor=C_EDGE, linewidth=1.5, zorder=3)
    ax.add_patch(p)
    txt(ax, cx, cy, text, size=fsize)
    return dict(cx=cx, cy=cy, hw=hw, hh=hh, kind="rect")


def case_box(ax, cx, cy, hw, hh, text, fsize=10.0, fc=C_CASE):
    """ケース入口: 角丸四角（やや丸め）"""
    p = mpatches.FancyBboxPatch((cx - hw, cy - hh), hw * 2, hh * 2,
                                  boxstyle="round,pad=0.04,rounding_size=0.18",
                                  facecolor=fc, edgecolor=C_EDGE, linewidth=1.5,
                                  linestyle="dashed", zorder=3)
    ax.add_patch(p)
    txt(ax, cx, cy, text, size=fsize)
    return dict(cx=cx, cy=cy, hw=hw, hh=hh, kind="rect")


def diamond(ax, cx, cy, hw, hh, text, fsize=10.5, fc=C_DEC):
    """分岐: ひし形◇"""
    pts = np.array([[cx, cy + hh], [cx + hw, cy],
                    [cx, cy - hh], [cx - hw, cy]])
    poly = mpatches.Polygon(pts, closed=True,
                             facecolor=fc, edgecolor=C_EDGE, linewidth=1.5, zorder=3)
    ax.add_patch(poly)
    txt(ax, cx, cy, text, size=fsize)
    return dict(cx=cx, cy=cy, hw=hw, hh=hh, kind="diamond")


def edge_pt(shape, direction):
    """図形のエッジ座標（接続点）を返す"""
    cx, cy = shape["cx"], shape["cy"]
    if shape["kind"] == "oval":
        rw, rh = shape["rw"], shape["rh"]
        return {"top":(cx,cy+rh),"bottom":(cx,cy-rh),
                "left":(cx-rw,cy),"right":(cx+rw,cy)}[direction]
    elif shape["kind"] == "rect":
        hw, hh = shape["hw"], shape["hh"]
        return {"top":(cx,cy+hh),"bottom":(cx,cy-hh),
                "left":(cx-hw,cy),"right":(cx+hw,cy)}[direction]
    elif shape["kind"] == "diamond":
        hw, hh = shape["hw"], shape["hh"]
        return {"top":(cx,cy+hh),"bottom":(cx,cy-hh),
                "left":(cx-hw,cy),"right":(cx+hw,cy)}[direction]


def arrow(ax, p1, p2, label=None, lc=C_ARROW, via=None):
    """矢印描画。via=(x,y) で折れ線"""
    style = dict(arrowstyle="-|>", mutation_scale=14, color=lc, linewidth=1.4)
    if via:
        ax.annotate("", xy=via, xytext=p1,
                    arrowprops=dict(arrowstyle="-", color=lc, linewidth=1.4))
        ax.annotate("", xy=p2,  xytext=via,
                    arrowprops=dict(**style))
    else:
        ax.annotate("", xy=p2, xytext=p1,
                    arrowprops=dict(**style))
    if label:
        mx = (p1[0] + (via[0] if via else p2[0])) / 2
        my = (p1[1] + (via[1] if via else p2[1])) / 2
        ax.text(mx + 0.08, my, label, ha="left", va="center",
                fontproperties=FP, fontsize=10.5, color=lc,
                fontweight="bold",
                bbox=dict(fc="white", ec="none", pad=1))


# ── メイン ──────────────────────────────────────
def main():
    fig, ax = plt.subplots(figsize=(14, 11))
    ax.set_xlim(0, 14)
    ax.set_ylim(-0.6, 11)
    ax.axis("off")
    fig.suptitle("LPN分割 誤り対応フロー", fontproperties=FP, fontsize=15,
                  fontweight="bold", y=0.97)

    # ════════════════════════════════════════
    # ① LPN分割を忘れて即出荷
    # ════════════════════════════════════════
    s1  = oval(ax,  1.1, 10.2, 0.35, 0.28, "開始", fsize=10)
    c1  = case_box(ax, 1.7, 9.5, 1.3, 0.38,
                   "①LPN分割を忘れて\n即出荷してしまった", fsize=9.5)
    d1  = diamond(ax, 4.5, 9.5, 1.5, 0.45,
                  "即出荷後に\nLPN分割したか？", fsize=10)
    act_olpn = rect(ax, 8.5, 10.2, 1.6, 0.38, "OLPNを統合する", fc=C_ACT)
    e1  = oval(ax, 11.0, 10.2, 0.35, 0.28, "終了", fsize=10)

    d2  = diamond(ax, 6.8, 8.5, 1.4, 0.45,
                  "ワンレックか？\n複数レックか？", fsize=10)
    act_w1 = rect(ax, 10.2, 9.2, 1.7, 0.38, "国内梱包\n（ワンレック）", fc=C_ACT)
    e_w1   = oval(ax, 12.8, 9.2, 0.35, 0.28, "終了", fsize=10)
    act_w2 = rect(ax, 10.2, 7.9, 1.7, 0.50,
                  "国内梱包（複数レック）\n＋イレギュラー置場へ", fc=C_ACT, fsize=9.5)
    e_w2   = oval(ax, 12.8, 7.9, 0.35, 0.28, "終了", fsize=10)

    arrow(ax, edge_pt(s1,"bottom"), edge_pt(c1,"top"))
    arrow(ax, edge_pt(c1,"right"),  edge_pt(d1,"left"))
    arrow(ax, edge_pt(d1,"top"),    edge_pt(act_olpn,"left"),
          label="YES", lc=C_YES, via=(4.5, 10.2))
    arrow(ax, edge_pt(act_olpn,"right"), edge_pt(e1,"left"))

    arrow(ax, edge_pt(d1,"right"),  edge_pt(d2,"left"),  label="NO",  lc=C_NO)
    arrow(ax, edge_pt(d2,"top"),    edge_pt(act_w1,"left"),
          label="ワンレック", lc=C_YES, via=(6.8, 9.2))
    arrow(ax, edge_pt(act_w1,"right"), edge_pt(e_w1,"left"))
    arrow(ax, edge_pt(d2,"bottom"), edge_pt(act_w2,"left"),
          label="複数レック", lc=C_NO, via=(6.8, 7.9))
    arrow(ax, edge_pt(act_w2,"right"), edge_pt(e_w2,"left"))

    # ════════════════════════════════════════
    # ② LPN分割で数量を誤った → OLPN統合
    # ════════════════════════════════════════
    s2     = oval(ax, 1.1, 7.0, 0.35, 0.28, "開始", fsize=10)
    c2     = case_box(ax, 1.7, 7.0, 1.3, 0.38,
                      "②LPN分割で\n数量を誤った", fsize=9.5)
    act2   = rect(ax, 5.5, 7.0, 1.7, 0.38, "OLPNを統合する", fc=C_ACT)
    e2     = oval(ax, 8.2, 7.0, 0.35, 0.28, "終了", fsize=10)

    arrow(ax, edge_pt(s2,"bottom"), edge_pt(c2,"top"))
    arrow(ax, edge_pt(c2,"right"),  edge_pt(act2,"left"))
    arrow(ax, edge_pt(act2,"right"),edge_pt(e2,"left"))

    # ════════════════════════════════════════
    # ③ LPN分割を過剰に行った
    # ════════════════════════════════════════
    s3  = oval(ax, 1.1, 5.4, 0.35, 0.28, "開始", fsize=10)
    c3  = case_box(ax, 1.7, 5.4, 1.3, 0.38,
                   "③LPN分割を\n過剰に行った", fsize=9.5)
    act3= rect(ax, 5.5, 5.4, 1.7, 0.38, "分割ラベルを\n使用する", fc=C_ACT)
    e3  = oval(ax, 8.2, 5.4, 0.35, 0.28, "終了", fsize=10)

    arrow(ax, edge_pt(s3,"bottom"), edge_pt(c3,"top"))
    arrow(ax, edge_pt(c3,"right"),  edge_pt(act3,"left"))
    arrow(ax, edge_pt(act3,"right"),edge_pt(e3,"left"))

    # ════════════════════════════════════════
    # ④ 複数部材LPN分割を中断した
    # ════════════════════════════════════════
    s4  = oval(ax, 1.1, 3.8, 0.35, 0.28, "開始", fsize=10)
    c4  = case_box(ax, 1.7, 3.8, 1.3, 0.50,
                   "④複数部材のLPN分割を\n途中で中断した", fsize=9.0)
    act4= rect(ax, 5.5, 3.8, 1.7, 0.38, "親部材集約を\n実施する", fc=C_ACT)
    e4  = oval(ax, 8.2, 3.8, 0.35, 0.28, "終了", fsize=10)

    arrow(ax, edge_pt(s4,"bottom"), edge_pt(c4,"top"))
    arrow(ax, edge_pt(c4,"right"),  edge_pt(act4,"left"))
    arrow(ax, edge_pt(act4,"right"),edge_pt(e4,"left"))

    # ════════════════════════════════════════
    # ⑤ プリンタ未設定のままLPN分割
    # ════════════════════════════════════════
    s5  = oval(ax, 1.1, 2.2, 0.35, 0.28, "開始", fsize=10)
    c5  = case_box(ax, 1.7, 2.2, 1.3, 0.50,
                   "⑤プリンタを設定しない\nままLPN分割した", fsize=9.0)
    act5= rect(ax, 5.5, 2.2, 1.7, 0.38, "MAラベルを\n再印刷する", fc=C_ACT)
    e5  = oval(ax, 8.2, 2.2, 0.35, 0.28, "終了", fsize=10)

    arrow(ax, edge_pt(s5,"bottom"), edge_pt(c5,"top"))
    arrow(ax, edge_pt(c5,"right"),  edge_pt(act5,"left"))
    arrow(ax, edge_pt(act5,"right"),edge_pt(e5,"left"))

    # ── 凡例 ────────────────────────────────
    lx, ly = 0.3, 0.8
    oval(ax,  lx+0.4,  ly, 0.28, 0.22, "開始/終了", fsize=9)
    rect(ax,  lx+2.2,  ly, 0.7,  0.22, "処理 □", fsize=9, fc=C_ACT)
    diamond(ax,lx+4.0, ly, 0.8,  0.30, "分岐 ◇", fsize=9, fc=C_DEC)
    case_box(ax,lx+6.0,ly, 0.8,  0.22, "発生事象", fsize=9, fc=C_CASE)

    ax.axhline(1.35, color="#aaaaaa", linewidth=0.8, linestyle="--")

    plt.tight_layout(rect=[0, 0.02, 1, 0.97])
    fig.savefig(PNG_FILE, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"作成完了: {PNG_FILE}")

    # ── Excelにも貼付 ───────────────────────
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    wb = Workbook()
    ws = wb.active
    ws.title = "LPN分割誤り対応フロー"
    img = XLImage(PNG_FILE)
    scale = 1400 / img.width
    img.width  = int(img.width  * scale)
    img.height = int(img.height * scale)
    ws.add_image(img, "A1")
    wb.save(XLSX_FILE)
    print(f"作成完了: {XLSX_FILE}")


if __name__ == "__main__":
    main()
