# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel Differ Git is a Git-integrated Excel diff tool that detects and displays changes in Excel files at the row and cell level. It uses content-based row matching to accurately detect insertions, deletions, and modifications even when row numbers shift.

**Key Features:**
- Git integration: Compare Excel files across commits or with working tree
- Direct file comparison: Compare two local files without Git
- Content-based row matching: Tracks actual changes despite row number shifts
- Multiple output formats: Text (human-readable) and CSV

## Development Commands

### Running the Tool

The tool can be run as a Python module:

```bash
# Direct file comparison (most common during development)
python -m excel_differ.cli --old test/old.xlsx --new test/new.xlsx test/old.xlsx

# Git comparison mode
python -m excel_differ.cli --from HEAD~1 --to HEAD file.xlsx

# Working tree comparison
python -m excel_differ.cli --working-tree file.xlsx

# CSV output
python -m excel_differ.cli --old v1.xlsx --new v2.xlsx --format csv --output diff.csv v1.xlsx
```

### Installation

**IMPORTANT**: This project supports complete offline installation. All dependencies are pre-downloaded in `vendor/`.

```bash
# Complete offline installation (recommended)
py -m pip install --no-index --find-links=vendor -r requirements.txt
py -m pip install --no-index --find-links=vendor -e .

# Or use the batch script on Windows
install_offline.bat

# After installation, the command is available globally
excel-diff --old old.xlsx --new new.xlsx old.xlsx
```

**Why `--no-index --find-links=vendor` is required for BOTH commands:**
- First command: Installs dependencies (openpyxl, GitPython, click, etc.)
- Second command: Installs the app itself, which requires `setuptools` and `wheel` for building
- Without these flags on the second command, pip will try to download setuptools from PyPI and timeout

**What's in vendor/:**
- Application dependencies: openpyxl, GitPython, click
- Transitive dependencies: et_xmlfile, gitdb, smmap, colorama
- Build tools: setuptools, wheel (required for `pip install -e .`)

### Testing

Test files are located in `test/` directory. Create test Excel files:

```bash
# Generate test Excel files
python test/create_test_excel.py
```

## Architecture

### Module Structure

```
excel_differ/
├── cli.py           # Entry point - argument parsing, mode selection
├── git_handler.py   # Git operations - extract files from commits/working tree
├── excel_reader.py  # Excel I/O - convert Excel files to internal representation
├── differ.py        # Core algorithm - content-based row matching and diff detection
└── formatter.py     # Output - text and CSV formatters
```

### Data Flow

1. **Input**: CLI receives file paths and comparison mode (Git or direct)
2. **Reading**: `excel_reader.py` loads Excel files into structured objects:
   - `ExcelWorkbook` → `ExcelSheet` → `ExcelRow` → cells (list)
3. **Git Integration** (if applicable): `git_handler.py` extracts file bytes from Git commits
4. **Diff Algorithm**: `differ.py` performs three-phase matching:
   - Phase 1: Find exact matches by row content hash
   - Phase 2: Find similar rows (≥50% cell match) as "modified"
   - Phase 3: Classify remaining rows as "added" or "deleted"
5. **Output**: `formatter.py` formats results as text or CSV

### Core Algorithm (differ.py)

The diff algorithm is the heart of this tool:

**Key Functions:**
- `find_row_matches()`: Creates content hash maps to find exact matches
- `find_similar_rows()`: Uses similarity scoring (≥50% threshold) to detect modifications
- `detect_cell_changes()`: Identifies which specific cells changed in modified rows
- `diff_sheets()`: Orchestrates the three-phase matching for a single sheet
- `diff_workbooks()`: Compares all sheets and detects sheet-level changes

**Important Details:**
- Row matching is content-based, not position-based
- Similarity threshold is 50% (hardcoded in `find_similar_rows()`)
- The algorithm handles row reordering correctly
- Empty cells are treated as `None`

### Git Integration (git_handler.py)

Uses GitPython library to:
- Extract file contents from specific commits as bytes
- Compare commits or working tree
- Handle relative paths correctly (converts to repo-root-relative)
- Supports all Git refs (HEAD, HEAD~1, commit hashes, branch names)

**Critical Implementation Notes:**
- File paths must be converted from Windows backslashes to forward slashes for Git
- Working tree comparison reads current file directly with `read_excel_file()`
- Commit comparison uses `read_excel_from_bytes()` for both sides

### Output Formats (formatter.py)

**Text Format:**
- Human-readable with sections for each sheet
- Shows row number changes (e.g., "Row 5 MODIFIED (was row 3)")
- Displays cell-by-cell changes with column letters

**CSV Format:**
- Machine-readable for Excel import
- Columns: type, sheet, old_row, new_row, column, cell, old_value, new_value, description
- Each cell change is a separate row

## Important Context

### Windows-Specific Considerations

This project is developed primarily for Windows:
- Uses backslashes in file paths (converted to forward slashes for Git)
- Command examples use `py` launcher (Python on Windows)
- Batch scripts for installation (`install_offline.bat`)

### Offline Installation Support

The `vendor/` directory contains all dependencies as wheel files to enable complete offline installation:

**Application dependencies:**
- openpyxl (Excel reading)
- GitPython (Git operations)
- click (CLI framework)

**Transitive dependencies:**
- et_xmlfile (required by openpyxl)
- gitdb, smmap (required by GitPython)
- colorama (required by click on Windows)

**Build tools:**
- setuptools (required for `pip install -e .`)
- wheel (required for building packages)

**Critical installation detail:**
Both installation commands MUST use `--no-index --find-links=vendor`:
```bash
py -m pip install --no-index --find-links=vendor -r requirements.txt
py -m pip install --no-index --find-links=vendor -e .
```

If the second command omits these flags, pip will attempt to download setuptools from PyPI (defined in `pyproject.toml` as `build-system.requires`), causing a timeout in offline environments.

**Why this matters:**
- Designed for secure/restricted environments without internet access
- Common use case: Corporate networks with strict firewall policies
- All dependencies must be pre-downloaded to `vendor/` before transferring to the target machine

### Testing Status

According to TEST_REPORT.md:
- Small-scale tests: ✅ Verified (basic functionality)
- Medium-scale tests: ✅ Verified (100×100×10 sheets, 2.2s processing)
- Large-scale tests: ⏳ File generation only
- Git integration: Not yet tested

### Known Limitations

- Large files (>1000×1000 cells) may be slow
- Similarity threshold is hardcoded at 50%
- No support for detecting formatting changes (colors, fonts)
- No support for formula change detection (only values)
- CSV output may need proper encoding for Japanese text

### File Path Handling

When working with file paths in this codebase:
- CLI accepts both absolute and relative paths
- Git handler converts paths to repo-root-relative
- Excel reader uses `pathlib.Path` objects
- Git operations require forward slashes (Windows backslashes are converted)

## Common Development Tasks

### Adding a New Output Format

1. Create a new formatter class in `formatter.py` (e.g., `JSONFormatter`)
2. Implement a `format(diff: WorkbookDiff, output: TextIO)` method
3. Add format choice to `format_diff()` function
4. Update CLI `--format` option in `cli.py`

### Modifying the Similarity Threshold

The 50% threshold is in `differ.py:184`:
```python
if score >= 0.5 and score > best_score:
```

Change `0.5` to desired threshold. Consider making it a CLI parameter.

### Supporting Additional Excel Formats

The tool uses openpyxl, which supports:
- `.xlsx` (Excel 2007+)
- `.xlsm` (macro-enabled)

For `.xls` (Excel 97-2003), would need to add `xlrd` library.

### Debugging Tips

- Use `--format csv` to get structured output for analysis
- Row matching happens in `find_row_matches()` - add debug prints there
- The `ExcelRow.to_string()` method shows how rows are compared (pipe-separated)
- Git path issues: Check `git_handler.py:48` for path conversion logic

## Dependencies

Core dependencies (see `pyproject.toml`):
- `openpyxl>=3.1.0` - Excel file reading (data_only=True for formulas)
- `GitPython>=3.1.0` - Git repository operations
- `click>=8.1.0` - CLI framework with decorators

Dev dependencies (optional):
- `pytest>=7.0.0`
- `black>=23.0.0`
- `mypy>=1.0.0`

## Entry Points

The main entry point is defined in `pyproject.toml`:
```toml
[project.scripts]
excel-diff = "excel_differ.cli:main"
```

This creates the `excel-diff` command after `pip install`.
