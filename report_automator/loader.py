"""Find and read the input files (CSV and Excel) into raw tables."""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterator, Optional

import pandas as pd

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xlsm"}


def discover_files(path: Path) -> list[Path]:
    """Return the supported files for a file path or a folder (non-recursive, sorted)."""
    path = Path(path)
    if path.is_file():
        return [path] if path.suffix.lower() in SUPPORTED_EXTENSIONS else []
    if not path.is_dir():
        raise FileNotFoundError(f"Input path not found: {path}")
    return sorted(
        p
        for p in path.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS and not p.name.startswith(("~$", "."))
    )


def _read_csv(path: Path) -> pd.DataFrame:
    last_error: Optional[Exception] = None
    for encoding in ("utf-8-sig", "cp1252"):
        for sniff in (True, False):
            try:
                kwargs = {"sep": None, "engine": "python"} if sniff else {"sep": ","}
                return pd.read_csv(path, dtype=object, encoding=encoding, keep_default_na=False, **kwargs)
            except pd.errors.EmptyDataError:
                return pd.DataFrame()
            except UnicodeDecodeError as exc:
                last_error = exc
                break  # try the next encoding
            except (pd.errors.ParserError, csv.Error) as exc:
                last_error = exc  # delimiter sniffing failed, retry with a comma
    raise ValueError(f"Could not read CSV file {path.name}: {last_error}")


def load_tables(path: Path) -> Iterator[tuple[str, Optional[pd.DataFrame], Optional[str]]]:
    """Yield (label, raw_table, error_message) for every CSV file and every Excel sheet."""
    for file in discover_files(path):
        try:
            if file.suffix.lower() == ".csv":
                yield file.name, _read_csv(file), None
            else:
                sheets = pd.read_excel(file, sheet_name=None, dtype=object)
                for sheet_name, table in sheets.items():
                    yield f"{file.name} [{sheet_name}]", table, None
        except Exception as exc:  # one bad file must not stop the whole run
            yield file.name, None, f"{type(exc).__name__}: {exc}"
