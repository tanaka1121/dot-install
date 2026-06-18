"""
LPN分割 誤り対応フロー - PowerPoint版（バランス調整済み）

ユーザーの既存スライド（場合分けLPN分割 誤り対応フロー）のラベルを踏襲し、
重なり・配置崩れを解消したレイアウトで再構築する。
手順書に貼り付けて使う想定のため、1スライドに収め、図形はすべて
PowerPoint上でそのまま編集できるネイティブ図形（pptxの標準Shape）のみで構成。
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

OUTPUT = "LPN分割_誤り対応フロー.pptx"

JP_FONT = "メイリオ"

COLOR_ENTRY = RGBColor(0xE7, 0xE6, 0xE6)   # 薄灰（発生事象）
COLOR_DECISION = RGBColor(0xFF, 0xE6, 0x99)  # 黄（判定）
COLOR_ACTION = RGBColor(0xDD, 0xEB, 0xF7)   # 青（対応）
COLOR_LINE = RGBColor(0x40, 0x40, 0x40)
COLOR_YES = RGBColor(0x37, 0x56, 0x23)
COLOR_NO = RGBColor(0x9C, 0x00, 0x06)


def add_box(slide, shape_type, x, y, w, h, text, fill, font_size=14, bold=False):
    sp = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    sp.line.color.rgb = COLOR_LINE
    sp.line.width = Pt(1.25)
    sp.shadow.inherit = False
    tf = sp.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name = JP_FONT
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return sp


def connect(slide, sp1, sp2, from_pos="r", to_pos="l", label=None, label_color=None, bend=None):
    """sp1からsp2へ矢印を引く。from_pos/to_posは l/r/t/b。bend指定でクランク接続。"""
    def point(sp, pos):
        x, y, w, h = sp.left, sp.top, sp.width, sp.height
        if pos == "r":
            return x + w, y + h // 2
        if pos == "l":
            return x, y + h // 2
        if pos == "t":
            return x + w // 2, y
        if pos == "b":
            return x + w // 2, y + h

    x1, y1 = point(sp1, from_pos)
    x2, y2 = point(sp2, to_pos)

    connector = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, x1, y1, x2, y2)
    connector.line.color.rgb = COLOR_LINE
    connector.line.width = Pt(1.5)
    line_elem = connector.line._get_or_add_ln()
    tail = line_elem.makeelement(qn("a:tailEnd"), {"type": "triangle"})
    line_elem.append(tail)

    if label:
        lx = (x1 + x2) / 2
        ly = (y1 + y2) / 2
        box = slide.shapes.add_textbox(lx - Inches(0.4), ly - Inches(0.18), Inches(0.8), Inches(0.3))
        tf = box.text_frame
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = label
        run.font.name = JP_FONT
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = label_color or COLOR_LINE
    return connector


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.1), Inches(10), Inches(0.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "場合分け：LPN分割　誤り対応フロー"
    run.font.name = JP_FONT
    run.font.size = Pt(20)
    run.font.bold = True

    # レイアウト規則:
    #   ・本流（まだ条件確認が続く経路）は同じ行を左→右に直進させる
    #   ・確定した対応（YES/結果）は必ず1行下に分岐させる（向きを統一）
    #   ・各ケースは専用の行（レーン）に置き、本流と交差させない
    # ── 行0：本流の判定チェーン ────────────────────────
    e1 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.3, 0.7, 2.3, 1.0,
                 "①LPN分割を忘れて\n即出荷してしまった", COLOR_ENTRY, font_size=13)
    d1 = add_box(slide, MSO_SHAPE.DIAMOND, 3.3, 0.6, 1.7, 1.2,
                 "即出荷後に\nLPN分割したか？", COLOR_DECISION, font_size=13)
    d2 = add_box(slide, MSO_SHAPE.DIAMOND, 7.3, 0.6, 1.6, 1.2,
                 "1RECか？", COLOR_DECISION, font_size=14)

    # ── 行1：①②の対応（行0の直下に整列）────────────────
    a1 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 3.3, 2.2, 1.7, 0.8,
                 "oLPN統合", COLOR_ACTION, font_size=14, bold=True)
    e2 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.3, 2.2, 2.3, 0.8,
                 "②LPN分割で\n数量を誤った", COLOR_ENTRY, font_size=13)
    a2 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 6.2, 2.2, 1.8, 0.8,
                 "国内梱包_1REC", COLOR_ACTION, font_size=13, bold=True)
    a3 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 8.5, 2.2, 2.6, 0.8,
                 "国内梱包_複数REC\n＋イレギュラー置き場へ", COLOR_ACTION, font_size=12, bold=True)

    # ── 行2〜4：③④⑤（行0/1とは独立した単純レーン）──────
    e3 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.3, 3.5, 2.3, 0.9,
                 "③LPN分割を\n過剰に行った", COLOR_ENTRY, font_size=13)
    a4 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 3.3, 3.5, 2.3, 0.9,
                 "分割ラベルを使用する", COLOR_ACTION, font_size=14, bold=True)

    e4 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.3, 4.7, 2.3, 0.9,
                 "④複数部材の分割を\n途中で中断した", COLOR_ENTRY, font_size=13)
    a5 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 3.3, 4.7, 2.3, 0.9,
                 "親部材集約", COLOR_ACTION, font_size=14, bold=True)

    e5 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.3, 5.9, 2.3, 0.9,
                 "⑤プリンタの設定を\nしないままLPN分割した", COLOR_ENTRY, font_size=13)
    a6 = add_box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 3.3, 5.9, 2.6, 0.9,
                 "MAからラベル再印刷", COLOR_ACTION, font_size=14, bold=True)

    # ── 接続線 ───────────────────────────────────────
    # 本流（行0）：YESは1行下へ、NOは同じ行を直進
    connect(slide, e1, d1, "r", "l")
    connect(slide, d1, a1, "b", "t", label="YES", label_color=COLOR_YES)
    connect(slide, d1, d2, "r", "l", label="NO", label_color=COLOR_NO)
    connect(slide, d2, a2, "b", "t", label="YES", label_color=COLOR_YES)
    connect(slide, d2, a3, "b", "t", label="NO", label_color=COLOR_NO)

    # 行1内（②はそのまま同じ行でoLPN統合へ）
    connect(slide, e2, a1, "r", "l")

    # 行2〜4（各ケース単独・本流と非交差）
    connect(slide, e3, a4, "r", "l")
    connect(slide, e4, a5, "r", "l")
    connect(slide, e5, a6, "r", "l")

    prs.save(OUTPUT)
    print(f"作成完了: {OUTPUT}")


if __name__ == "__main__":
    main()
