"""Git integration for Excel diff."""

from pathlib import Path
from typing import Optional, Tuple
import tempfile
import git
from .excel_reader import ExcelWorkbook, read_excel_file, read_excel_from_bytes


class GitHandler:
    """Handle Git operations for Excel files."""

    def __init__(self, repo_path: Optional[Path] = None):
        """
        Initialize Git handler.

        Args:
            repo_path: Path to git repository. If None, uses current directory.
        """
        try:
            if repo_path:
                self.repo = git.Repo(repo_path, search_parent_directories=True)
            else:
                self.repo = git.Repo(".", search_parent_directories=True)
        except git.InvalidGitRepositoryError:
            raise ValueError("Not a git repository")

    def get_file_at_commit(self, filepath: Path, commit: str = "HEAD") -> bytes:
        """
        Get file content at a specific commit.

        Args:
            filepath: Path to file relative to repository root
            commit: Git commit reference (hash, HEAD, HEAD~1, etc.)

        Returns:
            File content as bytes
        """
        # Get relative path from repo root
        repo_root = Path(self.repo.working_dir)
        try:
            rel_path = filepath.relative_to(repo_root)
        except ValueError:
            # If path is already relative
            rel_path = filepath

        # Convert to forward slashes for git
        git_path = str(rel_path).replace("\\", "/")

        try:
            commit_obj = self.repo.commit(commit)
            blob = commit_obj.tree / git_path
            return blob.data_stream.read()
        except (KeyError, AttributeError):
            raise FileNotFoundError(f"File '{git_path}' not found in commit '{commit}'")

    def get_workbook_at_commit(self, filepath: Path, commit: str = "HEAD") -> ExcelWorkbook:
        """
        Get Excel workbook at a specific commit.

        Args:
            filepath: Path to Excel file
            commit: Git commit reference

        Returns:
            ExcelWorkbook object
        """
        file_bytes = self.get_file_at_commit(filepath, commit)
        filename = f"{filepath.name} ({commit})"
        return read_excel_from_bytes(file_bytes, filename)

    def get_previous_commit(self, commit: str = "HEAD") -> str:
        """
        Get the parent commit of the specified commit.

        Args:
            commit: Git commit reference

        Returns:
            Parent commit hash
        """
        commit_obj = self.repo.commit(commit)
        if not commit_obj.parents:
            raise ValueError(f"Commit '{commit}' has no parent (initial commit)")
        return commit_obj.parents[0].hexsha

    def get_commit_info(self, commit: str = "HEAD") -> dict:
        """
        Get information about a commit.

        Args:
            commit: Git commit reference

        Returns:
            Dictionary with commit info (hash, message, author, date)
        """
        commit_obj = self.repo.commit(commit)
        return {
            "hash": commit_obj.hexsha[:7],
            "full_hash": commit_obj.hexsha,
            "message": commit_obj.message.strip(),
            "author": str(commit_obj.author),
            "date": commit_obj.committed_datetime.isoformat(),
        }

    def compare_commits(
        self, filepath: Path, from_commit: str = None, to_commit: str = "HEAD"
    ) -> Tuple[ExcelWorkbook, ExcelWorkbook]:
        """
        Get workbooks from two commits for comparison.

        Args:
            filepath: Path to Excel file
            from_commit: Old commit reference (if None, uses parent of to_commit)
            to_commit: New commit reference

        Returns:
            Tuple of (old_workbook, new_workbook)
        """
        if from_commit is None:
            from_commit = self.get_previous_commit(to_commit)

        old_wb = self.get_workbook_at_commit(filepath, from_commit)
        new_wb = self.get_workbook_at_commit(filepath, to_commit)

        return old_wb, new_wb

    def compare_with_working_tree(
        self, filepath: Path, commit: str = "HEAD"
    ) -> Tuple[ExcelWorkbook, ExcelWorkbook]:
        """
        Compare a commit with the current working tree.

        Args:
            filepath: Path to Excel file
            commit: Commit to compare against

        Returns:
            Tuple of (old_workbook, new_workbook)
        """
        old_wb = self.get_workbook_at_commit(filepath, commit)
        new_wb = read_excel_file(filepath)

        return old_wb, new_wb
