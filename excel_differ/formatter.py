"""Output formatters for Excel diff results."""

import csv
import sys
from typing import TextIO
from io import StringIO
from .differ import WorkbookDiff, SheetChange, RowChange, ChangeType


class TextFormatter:
    """Format diff output as human-readable text."""

    @staticmethod
    def format(diff: WorkbookDiff, output: TextIO = sys.stdout) -> None:
        """
        Format diff as text output.

        Args:
            diff: WorkbookDiff object to format
            output: Output stream (default: stdout)
        """
        output.write("=" * 70 + "\n")
        output.write(f"Excel Diff: {diff.new_file}\n")
        output.write(f"Comparing: {diff.old_file} <-> {diff.new_file}\n")
        output.write("=" * 70 + "\n\n")

        if not diff.sheet_changes:
            output.write("No changes detected.\n")
            return

        for sheet_change in diff.sheet_changes:
            TextFormatter._format_sheet(sheet_change, output)

        # Summary
        summary = diff.get_summary()
        output.write("\n" + "=" * 70 + "\n")
        output.write("Summary\n")
        output.write("=" * 70 + "\n")

        if summary["sheets_added"]:
            output.write(f"Sheets added: {summary['sheets_added']}\n")
        if summary["sheets_deleted"]:
            output.write(f"Sheets deleted: {summary['sheets_deleted']}\n")
        if summary["sheets_modified"]:
            output.write(f"Sheets modified: {summary['sheets_modified']}\n")

        output.write(f"Rows added: {summary['rows_added']}\n")
        output.write(f"Rows deleted: {summary['rows_deleted']}\n")
        output.write(f"Rows modified: {summary['rows_modified']}\n")

    @staticmethod
    def _format_sheet(sheet_change: SheetChange, output: TextIO) -> None:
        """Format a single sheet change."""
        if sheet_change.change_type == ChangeType.SHEET_ADDED:
            output.write(f"[Sheet ADDED] {sheet_change.sheet_name}\n\n")
            return

        if sheet_change.change_type == ChangeType.SHEET_DELETED:
            output.write(f"[Sheet DELETED] {sheet_change.sheet_name}\n\n")
            return

        # Modified sheet
        output.write(f"[Sheet: {sheet_change.sheet_name}]\n")

        if not sheet_change.row_changes:
            output.write("  No changes\n\n")
            return

        for row_change in sheet_change.row_changes:
            TextFormatter._format_row(row_change, output)

        output.write("\n")

    @staticmethod
    def _format_row(row_change: RowChange, output: TextIO) -> None:
        """Format a single row change."""
        if row_change.change_type == ChangeType.ADDED:
            row_num = row_change.new_row_number
            output.write(f"  Row {row_num} ADDED\n")
            output.write(f"    + {row_change.new_row.to_string()}\n")

        elif row_change.change_type == ChangeType.DELETED:
            row_num = row_change.old_row_number
            output.write(f"  Row {row_num} DELETED\n")
            output.write(f"    - {row_change.old_row.to_string()}\n")

        elif row_change.change_type == ChangeType.MODIFIED:
            old_num = row_change.old_row_number
            new_num = row_change.new_row_number
            output.write(f"  Row {new_num} MODIFIED")
            if old_num != new_num:
                output.write(f" (was row {old_num})")
            output.write("\n")

            for cell_change in row_change.cell_changes:
                output.write(f"    Column {cell_change.column_letter}: ")
                output.write(f'"{cell_change.old_value}" -> "{cell_change.new_value}"\n')


class CSVFormatter:
    """Format diff output as CSV."""

    @staticmethod
    def format(diff: WorkbookDiff, output: TextIO = sys.stdout) -> None:
        """
        Format diff as CSV output.

        Args:
            diff: WorkbookDiff object to format
            output: Output stream (default: stdout)
        """
        writer = csv.writer(output)

        # Header
        writer.writerow([
            "type",
            "sheet",
            "old_row",
            "new_row",
            "column",
            "cell",
            "old_value",
            "new_value",
            "description"
        ])

        for sheet_change in diff.sheet_changes:
            CSVFormatter._format_sheet(sheet_change, writer)

    @staticmethod
    def _format_sheet(sheet_change: SheetChange, writer: csv.writer) -> None:
        """Format a single sheet change."""
        if sheet_change.change_type == ChangeType.SHEET_ADDED:
            writer.writerow([
                "sheet_added",
                sheet_change.sheet_name,
                "", "", "", "", "", "",
                "Sheet added"
            ])
            return

        if sheet_change.change_type == ChangeType.SHEET_DELETED:
            writer.writerow([
                "sheet_deleted",
                sheet_change.sheet_name,
                "", "", "", "", "", "",
                "Sheet deleted"
            ])
            return

        # Modified sheet
        for row_change in sheet_change.row_changes:
            CSVFormatter._format_row(row_change, writer)

    @staticmethod
    def _format_row(row_change: RowChange, writer: csv.writer) -> None:
        """Format a single row change."""
        sheet = row_change.sheet_name
        old_row = row_change.old_row_number or ""
        new_row = row_change.new_row_number or ""

        if row_change.change_type == ChangeType.ADDED:
            writer.writerow([
                "row_added",
                sheet,
                old_row,
                new_row,
                "",
                "",
                "",
                row_change.new_row.to_string(),
                "Row added"
            ])

        elif row_change.change_type == ChangeType.DELETED:
            writer.writerow([
                "row_deleted",
                sheet,
                old_row,
                new_row,
                "",
                "",
                row_change.old_row.to_string(),
                "",
                "Row deleted"
            ])

        elif row_change.change_type == ChangeType.MODIFIED:
            for cell_change in row_change.cell_changes:
                cell_addr = f"{cell_change.column_letter}{new_row}"
                writer.writerow([
                    "cell_modified",
                    sheet,
                    old_row,
                    new_row,
                    cell_change.column_letter,
                    cell_addr,
                    cell_change.old_value or "",
                    cell_change.new_value or "",
                    "Cell modified"
                ])


def format_diff(diff: WorkbookDiff, format_type: str = "text", output: TextIO = sys.stdout) -> None:
    """
    Format diff output.

    Args:
        diff: WorkbookDiff object to format
        format_type: Output format ("text" or "csv")
        output: Output stream (default: stdout)
    """
    if format_type == "csv":
        CSVFormatter.format(diff, output)
    else:
        TextFormatter.format(diff, output)
