"""Excel差分検出アルゴリズム

行単位の内容ベース比較を実装。
行の挿入・削除があっても、実際の変更のみを正確に抽出します。
"""

from typing import List, Dict, Set, Tuple, Optional
from dataclasses import dataclass
from enum import Enum
from .excel_reader import ExcelRow, ExcelSheet, ExcelWorkbook


class ChangeType(Enum):
    """変更タイプの列挙"""
    ADDED = "added"          # 追加
    DELETED = "deleted"      # 削除
    MODIFIED = "modified"    # 変更
    SHEET_ADDED = "sheet_added"      # シート追加
    SHEET_DELETED = "sheet_deleted"  # シート削除


@dataclass
class CellChange:
    """セルの変更を表すクラス"""
    column_index: int           # 列インデックス（0始まり）
    column_letter: str          # 列文字（A, B, C...）
    old_value: Optional[str]    # 変更前の値
    new_value: Optional[str]    # 変更後の値


@dataclass
class RowChange:
    """行の変更を表すクラス"""
    change_type: ChangeType              # 変更タイプ
    sheet_name: str                      # シート名
    old_row_number: Optional[int]        # 変更前の行番号
    new_row_number: Optional[int]        # 変更後の行番号
    old_row: Optional[ExcelRow]          # 変更前の行データ
    new_row: Optional[ExcelRow]          # 変更後の行データ
    cell_changes: List[CellChange]       # セル単位の変更リスト

    def __repr__(self) -> str:
        row_num = self.new_row_number or self.old_row_number
        return f"RowChange({self.change_type.value}, {self.sheet_name}, row={row_num})"


@dataclass
class SheetChange:
    """シートの変更を表すクラス"""
    change_type: ChangeType           # 変更タイプ
    sheet_name: str                   # シート名
    row_changes: List[RowChange]      # 行単位の変更リスト


@dataclass
class WorkbookDiff:
    """ブック全体の差分を表すクラス"""
    old_file: str                     # 変更前のファイル名
    new_file: str                     # 変更後のファイル名
    sheet_changes: List[SheetChange]  # シート単位の変更リスト

    def get_summary(self) -> Dict[str, int]:
        """変更の統計情報を取得"""
        summary = {
            "sheets_added": 0,
            "sheets_deleted": 0,
            "sheets_modified": 0,
            "rows_added": 0,
            "rows_deleted": 0,
            "rows_modified": 0,
        }

        for sheet_change in self.sheet_changes:
            if sheet_change.change_type == ChangeType.SHEET_ADDED:
                summary["sheets_added"] += 1
            elif sheet_change.change_type == ChangeType.SHEET_DELETED:
                summary["sheets_deleted"] += 1
            else:
                if sheet_change.row_changes:
                    summary["sheets_modified"] += 1

            for row_change in sheet_change.row_changes:
                if row_change.change_type == ChangeType.ADDED:
                    summary["rows_added"] += 1
                elif row_change.change_type == ChangeType.DELETED:
                    summary["rows_deleted"] += 1
                elif row_change.change_type == ChangeType.MODIFIED:
                    summary["rows_modified"] += 1

        return summary


def column_index_to_letter(index: int) -> str:
    """列インデックス（0始まり）をExcel列文字（A, B, ...）に変換"""
    letter = ""
    index += 1  # 1始まりに変換
    while index > 0:
        index -= 1
        letter = chr(index % 26 + ord('A')) + letter
        index //= 26
    return letter


def find_row_matches(old_rows: List[ExcelRow], new_rows: List[ExcelRow]) -> Tuple[
    Dict[int, int],  # old_idx -> new_idx (完全一致)
    Set[int],        # 一致しなかった旧行のインデックス
    Set[int],        # 一致しなかった新行のインデックス
]:
    """
    内容ベースの比較で行のマッチングを検出

    Returns:
        (マッチング辞書, 未一致の旧行, 未一致の新行) のタプル
    """
    # 内容のハッシュからインデックスへのマッピングを作成
    old_content_map: Dict[str, List[int]] = {}
    new_content_map: Dict[str, List[int]] = {}

    for idx, row in enumerate(old_rows):
        content = row.to_string()
        if content not in old_content_map:
            old_content_map[content] = []
        old_content_map[content].append(idx)

    for idx, row in enumerate(new_rows):
        content = row.to_string()
        if content not in new_content_map:
            new_content_map[content] = []
        new_content_map[content].append(idx)

    # 完全一致する行を検出
    matches = {}
    matched_old = set()
    matched_new = set()

    for content, old_indices in old_content_map.items():
        if content in new_content_map:
            new_indices = new_content_map[content]
            # 行を1対1でマッチング
            for old_idx, new_idx in zip(old_indices, new_indices):
                matches[old_idx] = new_idx
                matched_old.add(old_idx)
                matched_new.add(new_idx)

    unmatched_old = set(range(len(old_rows))) - matched_old
    unmatched_new = set(range(len(new_rows))) - matched_new

    return matches, unmatched_old, unmatched_new


def find_similar_rows(old_rows: List[ExcelRow], new_rows: List[ExcelRow],
                     unmatched_old: Set[int], unmatched_new: Set[int]) -> List[Tuple[int, int]]:
    """
    類似した行（変更の可能性）を検出

    50%以上のセルが一致する行を「変更」とみなします。

    Returns:
        (旧行インデックス, 新行インデックス) のペアのリスト
    """
    similar_pairs = []

    for old_idx in unmatched_old:
        old_row = old_rows[old_idx]
        best_match = None
        best_score = 0

        for new_idx in unmatched_new:
            new_row = new_rows[new_idx]

            # 類似度計算: 一致するセルの数をカウント
            max_len = max(len(old_row.cells), len(new_row.cells))
            if max_len == 0:
                continue

            matching_cells = sum(
                1 for i in range(min(len(old_row.cells), len(new_row.cells)))
                if old_row.cells[i] == new_row.cells[i]
            )

            score = matching_cells / max_len

            # 50%以上の類似度で「変更」とみなす
            if score >= 0.5 and score > best_score:
                best_score = score
                best_match = new_idx

        if best_match is not None:
            similar_pairs.append((old_idx, best_match))

    return similar_pairs


def detect_cell_changes(old_row: ExcelRow, new_row: ExcelRow) -> List[CellChange]:
    """2つの行間でどのセルが変更されたかを検出"""
    changes = []
    max_len = max(len(old_row.cells), len(new_row.cells))

    for i in range(max_len):
        old_val = old_row.cells[i] if i < len(old_row.cells) else None
        new_val = new_row.cells[i] if i < len(new_row.cells) else None

        if old_val != new_val:
            changes.append(CellChange(
                column_index=i,
                column_letter=column_index_to_letter(i),
                old_value=str(old_val) if old_val is not None else None,
                new_value=str(new_val) if new_val is not None else None,
            ))

    return changes


def diff_sheets(old_sheet: ExcelSheet, new_sheet: ExcelSheet) -> List[RowChange]:
    """
    2つのシートを比較して行レベルの変更を検出

    Returns:
        RowChangeオブジェクトのリスト
    """
    row_changes = []

    # 完全一致する行を検出
    matches, unmatched_old, unmatched_new = find_row_matches(old_sheet.rows, new_sheet.rows)

    # 類似した行（変更の可能性）を検出
    similar_pairs = find_similar_rows(old_sheet.rows, new_sheet.rows, unmatched_old, unmatched_new)

    # 変更された行を処理
    for old_idx, new_idx in similar_pairs:
        old_row = old_sheet.rows[old_idx]
        new_row = new_sheet.rows[new_idx]
        cell_changes = detect_cell_changes(old_row, new_row)

        row_changes.append(RowChange(
            change_type=ChangeType.MODIFIED,
            sheet_name=old_sheet.name,
            old_row_number=old_row.row_number,
            new_row_number=new_row.row_number,
            old_row=old_row,
            new_row=new_row,
            cell_changes=cell_changes,
        ))

        unmatched_old.discard(old_idx)
        unmatched_new.discard(new_idx)

    # 削除された行を処理
    for old_idx in sorted(unmatched_old):
        old_row = old_sheet.rows[old_idx]
        row_changes.append(RowChange(
            change_type=ChangeType.DELETED,
            sheet_name=old_sheet.name,
            old_row_number=old_row.row_number,
            new_row_number=None,
            old_row=old_row,
            new_row=None,
            cell_changes=[],
        ))

    # 追加された行を処理
    for new_idx in sorted(unmatched_new):
        new_row = new_sheet.rows[new_idx]
        row_changes.append(RowChange(
            change_type=ChangeType.ADDED,
            sheet_name=new_sheet.name,
            old_row_number=None,
            new_row_number=new_row.row_number,
            old_row=None,
            new_row=new_row,
            cell_changes=[],
        ))

    return row_changes


def diff_workbooks(old_wb: ExcelWorkbook, new_wb: ExcelWorkbook) -> WorkbookDiff:
    """
    2つのExcelブックを比較してすべての変更を検出

    Args:
        old_wb: 変更前のワークブック
        new_wb: 変更後のワークブック

    Returns:
        検出されたすべての変更を含むWorkbookDiffオブジェクト
    """
    sheet_changes = []

    old_sheet_names = set(old_wb.sheets.keys())
    new_sheet_names = set(new_wb.sheets.keys())

    # 削除されたシート
    for sheet_name in sorted(old_sheet_names - new_sheet_names):
        sheet_changes.append(SheetChange(
            change_type=ChangeType.SHEET_DELETED,
            sheet_name=sheet_name,
            row_changes=[],
        ))

    # 追加されたシート
    for sheet_name in sorted(new_sheet_names - old_sheet_names):
        sheet_changes.append(SheetChange(
            change_type=ChangeType.SHEET_ADDED,
            sheet_name=sheet_name,
            row_changes=[],
        ))

    # 変更されたシート（両方に存在するシート）
    for sheet_name in sorted(old_sheet_names & new_sheet_names):
        old_sheet = old_wb.sheets[sheet_name]
        new_sheet = new_wb.sheets[sheet_name]

        row_changes = diff_sheets(old_sheet, new_sheet)

        if row_changes:  # 変更があるシートのみ含める
            sheet_changes.append(SheetChange(
                change_type=ChangeType.MODIFIED,
                sheet_name=sheet_name,
                row_changes=row_changes,
            ))

    return WorkbookDiff(
        old_file=str(old_wb.filepath.name),
        new_file=str(new_wb.filepath.name),
        sheet_changes=sheet_changes,
    )
