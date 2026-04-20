"""
Excel セクション複製スクリプト

機能:
1. N列1〜32の文字列を A3, A16, A29, A42, ... (13行間隔) のタイトルセルに数式で連動
2. 既存の4セクション構造を16セクションへ複製（書式・行高・結合セルをコピー）

使用方法:
    python excel_dynamic_cell_copy.py <Excelファイル> [シート名]

例:
    python excel_dynamic_cell_copy.py DPC_template.xlsx
    python excel_dynamic_cell_copy.py DPC_template.xlsx Sheet1
"""

import sys
import copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


SECTION_HEIGHT = 13   # 1セクションの行数（3→16→29→42 の間隔）
TITLE_ROW_OFFSET = 2  # セクション開始行からタイトル行までのオフセット（行1→タイトル行3）
TOTAL_SECTIONS = 16   # 作成するセクション総数
TEMPLATE_SECTIONS = 4 # 既存テンプレートのセクション数（A3, A16, A29, A42）


def copy_cell_style(src, dst):
    """セルの書式をコピーする"""
    if src.has_style:
        dst.font = copy.copy(src.font)
        dst.border = copy.copy(src.border)
        dst.fill = copy.copy(src.fill)
        dst.number_format = src.number_format
        dst.protection = copy.copy(src.protection)
        dst.alignment = copy.copy(src.alignment)


def section_start_row(section_index):
    """セクションインデックス(0始まり)から開始行を返す"""
    return 1 + section_index * SECTION_HEIGHT


def title_row(section_index):
    """セクションインデックス(0始まり)からタイトルセルの行番号を返す"""
    return section_start_row(section_index) + TITLE_ROW_OFFSET


def duplicate_section(ws, src_start, dst_start):
    """
    src_start から SECTION_HEIGHT 行分の書式・値を dst_start へ複製する。
    タイトルセル（オフセット行）の値は後で数式に上書きするためそのままコピー。
    """
    src_end = src_start + SECTION_HEIGHT - 1

    # 行の高さをコピー
    for offset in range(SECTION_HEIGHT):
        src_row = src_start + offset
        dst_row = dst_start + offset
        if src_row in ws.row_dimensions:
            ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height

    # セルの値と書式をコピー
    for row in ws.iter_rows(min_row=src_start, max_row=src_end):
        for cell in row:
            offset = cell.row - src_start
            dst_cell = ws.cell(row=dst_start + offset, column=cell.column)
            dst_cell.value = cell.value
            copy_cell_style(cell, dst_cell)

    # 結合セルをコピー
    for merge_range in list(ws.merged_cells.ranges):
        if merge_range.min_row >= src_start and merge_range.max_row <= src_end:
            row_offset = dst_start - src_start
            new_min_row = merge_range.min_row + row_offset
            new_max_row = merge_range.max_row + row_offset
            min_col = get_column_letter(merge_range.min_col)
            max_col = get_column_letter(merge_range.max_col)
            ws.merge_cells(f"{min_col}{new_min_row}:{max_col}{new_max_row}")


def set_title_formulas(ws, total_sections):
    """全セクションのタイトルセル (A列) に N列参照数式をセット"""
    for i in range(total_sections):
        row = title_row(i)
        n_row = i + 1  # N1, N2, N3, ...
        ws.cell(row=row, column=1).value = f"=N{n_row}"
        print(f"  A{row} → =N{n_row}")


def main(filename, sheet_name=None):
    print(f"ファイルを開いています: {filename}")
    wb = load_workbook(filename)
    ws = wb.active if sheet_name is None else wb[sheet_name]
    print(f"シート: {ws.title}")

    # セクション 5〜16 をテンプレートから複製
    print(f"\n--- セクション複製 ({TEMPLATE_SECTIONS + 1}〜{TOTAL_SECTIONS}) ---")
    for i in range(TEMPLATE_SECTIONS, TOTAL_SECTIONS):
        # テンプレートセクションを循環して使用 (0, 1, 2, 3, 0, 1, ...)
        tmpl_idx = i % TEMPLATE_SECTIONS
        src_start = section_start_row(tmpl_idx)
        dst_start = section_start_row(i)
        print(f"  セクション{i + 1}: 行{src_start}〜{src_start + SECTION_HEIGHT - 1} → 行{dst_start}〜{dst_start + SECTION_HEIGHT - 1}")
        duplicate_section(ws, src_start, dst_start)

    # 全16セクションのタイトルセルに N列数式をセット
    print(f"\n--- タイトルセルに N列数式をセット ---")
    set_title_formulas(ws, TOTAL_SECTIONS)

    wb.save(filename)
    print(f"\n完了: {filename} を保存しました")
    print(f"  ・{TOTAL_SECTIONS} セクション分の構造を作成")
    print(f"  ・A列タイトルセルを N1〜N{TOTAL_SECTIONS} に連動する数式に設定")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("使用方法: python excel_dynamic_cell_copy.py <Excelファイル> [シート名]")
        print("")
        print("例:")
        print("  python excel_dynamic_cell_copy.py DPC_template.xlsx")
        print("  python excel_dynamic_cell_copy.py DPC_template.xlsx Sheet1")
        sys.exit(1)

    excel_file = sys.argv[1]
    sheet = sys.argv[2] if len(sys.argv) > 2 else None
    main(excel_file, sheet)
