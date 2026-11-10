"""End-to-end pipeline: load -> clean -> de-duplicate -> report."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import pandas as pd

from .cleaning import FRAME_COLUMNS, Issue, clean_frame, remove_duplicates
from .loader import load_tables
from .report import build_workbook


class NoDataError(Exception):
    """Raised when no valid rows were found in the input."""


@dataclass
class RunStats:
    files: int
    tables: int
    rows_read: int
    rows_kept: int
    rejected: int
    duplicates: int
    output: Path


def run(
    input_path: Path,
    output_path: Path,
    dayfirst: bool = True,
    dedupe: bool = True,
    log: Optional[Callable[[str], None]] = None,
) -> RunStats:
    log = log or (lambda message: None)
    frames: list[pd.DataFrame] = []
    issues: list[Issue] = []
    rows_read = 0
    tables = 0
    source_files: set[str] = set()

    for label, table, error in load_tables(Path(input_path)):
        source_files.add(label.split(" [")[0])
        if error:
            issues.append(Issue(label, None, f"File could not be read: {error}"))
            log(f"  ! {label}: could not be read")
            continue
        tables += 1
        result = clean_frame(table, label, dayfirst=dayfirst)
        rows_read += result.rows_read
        issues.extend(result.issues)
        if not result.frame.empty:
            frames.append(result.frame)
        log(f"  - {label}: {len(result.frame)} valid rows, {len(result.issues)} issue(s)")

    if not frames:
        raise NoDataError("No valid rows found. Check the 'Issues' messages above and your input files.")

    combined = pd.concat(frames, ignore_index=True)[FRAME_COLUMNS]
    before = len(combined)
    duplicate_issues: list[Issue] = []
    if dedupe:
        combined, duplicate_issues = remove_duplicates(combined)
    combined = combined.sort_values(["date", "source", "source_row"], kind="stable").reset_index(drop=True)
    issues.extend(duplicate_issues)

    workbook = build_workbook(combined, issues, files=len(source_files))
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)

    rejected = len(issues) - len(duplicate_issues)
    return RunStats(
        files=len(source_files),
        tables=tables,
        rows_read=rows_read,
        rows_kept=len(combined),
        rejected=rejected,
        duplicates=before - len(combined),
        output=output_path,
    )
