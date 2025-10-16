"""Excel diff detection algorithm."""

from typing import List, Dict, Set, Tuple, Optional
from dataclasses import dataclass
from enum import Enum
from .excel_reader import ExcelRow, ExcelSheet, ExcelWorkbook


class ChangeType(Enum):
    """Types of changes detected."""
    ADDED = "added"
    DELETED = "deleted"
    MODIFIED = "modified"
    SHEET_ADDED = "sheet_added"
    SHEET_DELETED = "sheet_deleted"


@dataclass
class CellChange:
    """Represents a change in a single cell."""
    column_index: int
    column_letter: str
    old_value: Optional[str]
    new_value: Optional[str]


@dataclass
class RowChange:
    """Represents a change in a row."""
    change_type: ChangeType
    sheet_name: str
    old_row_number: Optional[int]
    new_row_number: Optional[int]
    old_row: Optional[ExcelRow]
    new_row: Optional[ExcelRow]
    cell_changes: List[CellChange]

    def __repr__(self) -> str:
        row_num = self.new_row_number or self.old_row_number
        return f"RowChange({self.change_type.value}, {self.sheet_name}, row={row_num})"


@dataclass
class SheetChange:
    """Represents a change in a sheet."""
    change_type: ChangeType
    sheet_name: str
    row_changes: List[RowChange]


@dataclass
class WorkbookDiff:
    """Represents the complete diff between two workbooks."""
    old_file: str
    new_file: str
    sheet_changes: List[SheetChange]

    def get_summary(self) -> Dict[str, int]:
        """Get summary statistics of changes."""
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
    """Convert column index (0-based) to Excel column letter."""
    letter = ""
    index += 1  # Convert to 1-based
    while index > 0:
        index -= 1
        letter = chr(index % 26 + ord('A')) + letter
        index //= 26
    return letter


def find_row_matches(old_rows: List[ExcelRow], new_rows: List[ExcelRow]) -> Tuple[
    Dict[int, int],  # old_idx -> new_idx (exact matches)
    Set[int],  # unmatched old indices
    Set[int],  # unmatched new indices
]:
    """
    Find matching rows between old and new sheets using content-based comparison.

    Returns:
        Tuple of (matches_dict, unmatched_old_indices, unmatched_new_indices)
    """
    # Create content hash to indices mapping
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

    # Find exact matches
    matches = {}
    matched_old = set()
    matched_new = set()

    for content, old_indices in old_content_map.items():
        if content in new_content_map:
            new_indices = new_content_map[content]
            # Match rows one-to-one
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
    Find similar rows that might be modifications.
    Uses simple similarity heuristic: at least one cell matches.

    Returns:
        List of (old_idx, new_idx) pairs
    """
    similar_pairs = []

    for old_idx in unmatched_old:
        old_row = old_rows[old_idx]
        best_match = None
        best_score = 0

        for new_idx in unmatched_new:
            new_row = new_rows[new_idx]

            # Simple similarity: count matching cells
            max_len = max(len(old_row.cells), len(new_row.cells))
            if max_len == 0:
                continue

            matching_cells = sum(
                1 for i in range(min(len(old_row.cells), len(new_row.cells)))
                if old_row.cells[i] == new_row.cells[i]
            )

            score = matching_cells / max_len

            # Require at least 50% similarity to consider it a modification
            if score >= 0.5 and score > best_score:
                best_score = score
                best_match = new_idx

        if best_match is not None:
            similar_pairs.append((old_idx, best_match))

    return similar_pairs


def detect_cell_changes(old_row: ExcelRow, new_row: ExcelRow) -> List[CellChange]:
    """Detect which cells changed between two rows."""
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
    Compare two sheets and detect row-level changes.

    Returns:
        List of RowChange objects
    """
    row_changes = []

    # Find exact matches
    matches, unmatched_old, unmatched_new = find_row_matches(old_sheet.rows, new_sheet.rows)

    # Find similar rows (potential modifications)
    similar_pairs = find_similar_rows(old_sheet.rows, new_sheet.rows, unmatched_old, unmatched_new)

    # Process modifications
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

    # Process deletions
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

    # Process additions
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
    Compare two Excel workbooks and detect all changes.

    Args:
        old_wb: Old version of the workbook
        new_wb: New version of the workbook

    Returns:
        WorkbookDiff object containing all detected changes
    """
    sheet_changes = []

    old_sheet_names = set(old_wb.sheets.keys())
    new_sheet_names = set(new_wb.sheets.keys())

    # Deleted sheets
    for sheet_name in sorted(old_sheet_names - new_sheet_names):
        sheet_changes.append(SheetChange(
            change_type=ChangeType.SHEET_DELETED,
            sheet_name=sheet_name,
            row_changes=[],
        ))

    # Added sheets
    for sheet_name in sorted(new_sheet_names - old_sheet_names):
        sheet_changes.append(SheetChange(
            change_type=ChangeType.SHEET_ADDED,
            sheet_name=sheet_name,
            row_changes=[],
        ))

    # Modified sheets (sheets present in both)
    for sheet_name in sorted(old_sheet_names & new_sheet_names):
        old_sheet = old_wb.sheets[sheet_name]
        new_sheet = new_wb.sheets[sheet_name]

        row_changes = diff_sheets(old_sheet, new_sheet)

        if row_changes:  # Only include sheets with changes
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
