"""Command-line interface for Excel Differ."""

import sys
from pathlib import Path
from typing import Optional
import click
from .excel_reader import read_excel_file
from .differ import diff_workbooks
from .formatter import format_diff
from .git_handler import GitHandler


@click.command()
@click.argument("file", type=click.Path(exists=True, path_type=Path))
@click.option(
    "--old",
    type=click.Path(exists=True, path_type=Path),
    help="Old Excel file for direct comparison (cannot be used with Git options)"
)
@click.option(
    "--new",
    type=click.Path(exists=True, path_type=Path),
    help="New Excel file for direct comparison (cannot be used with Git options)"
)
@click.option(
    "--from",
    "from_commit",
    type=str,
    help="Git commit/ref to compare from (e.g., HEAD~1, commit hash)"
)
@click.option(
    "--to",
    "to_commit",
    type=str,
    default="HEAD",
    help="Git commit/ref to compare to (default: HEAD)"
)
@click.option(
    "--format",
    "output_format",
    type=click.Choice(["text", "csv"], case_sensitive=False),
    default="text",
    help="Output format (default: text)"
)
@click.option(
    "--output", "-o",
    type=click.Path(path_type=Path),
    help="Output file (default: stdout)"
)
@click.option(
    "--working-tree",
    is_flag=True,
    help="Compare commit with working tree (unstaged changes)"
)
def main(
    file: Path,
    old: Optional[Path],
    new: Optional[Path],
    from_commit: Optional[str],
    to_commit: str,
    output_format: str,
    output: Optional[Path],
    working_tree: bool,
) -> None:
    """
    Excel Differ - Git-integrated Excel diff tool.

    Compare Excel files to see changes in sheets, rows, and cells.

    Examples:

    \b
    # Compare latest commit with previous commit
    excel-diff myfile.xlsx

    \b
    # Compare specific commits
    excel-diff --from HEAD~2 --to HEAD myfile.xlsx

    \b
    # Compare with working tree (unstaged changes)
    excel-diff --working-tree myfile.xlsx

    \b
    # Direct file comparison
    excel-diff --old old.xlsx --new new.xlsx old.xlsx

    \b
    # Output as CSV
    excel-diff --format csv myfile.xlsx > diff.csv
    """
    try:
        # Validate mutually exclusive options
        direct_comparison = old is not None or new is not None
        git_comparison = from_commit is not None or working_tree

        if direct_comparison and git_comparison:
            click.echo("Error: Cannot use --old/--new with Git options (--from/--to/--working-tree)", err=True)
            sys.exit(1)

        # Direct file comparison
        if old and new:
            old_wb = read_excel_file(old)
            new_wb = read_excel_file(new)
            diff = diff_workbooks(old_wb, new_wb)

        # Git comparison
        else:
            git_handler = GitHandler()

            if working_tree:
                # Compare commit with working tree
                commit = from_commit or "HEAD"
                old_wb, new_wb = git_handler.compare_with_working_tree(file, commit)
                diff = diff_workbooks(old_wb, new_wb)

            else:
                # Compare commits
                old_wb, new_wb = git_handler.compare_commits(file, from_commit, to_commit)
                diff = diff_workbooks(old_wb, new_wb)

                # Add commit info to output
                from_info = git_handler.get_commit_info(from_commit or git_handler.get_previous_commit(to_commit))
                to_info = git_handler.get_commit_info(to_commit)
                diff.old_file = f"{file.name} ({from_info['hash']})"
                diff.new_file = f"{file.name} ({to_info['hash']})"

        # Output results
        if output:
            with open(output, "w", encoding="utf-8") as f:
                format_diff(diff, output_format, f)
            click.echo(f"Diff written to: {output}")
        else:
            format_diff(diff, output_format, sys.stdout)

    except ValueError as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)
    except FileNotFoundError as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)
    except Exception as e:
        click.echo(f"Unexpected error: {e}", err=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
