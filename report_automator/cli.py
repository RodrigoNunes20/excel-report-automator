"""Command line interface."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional, Sequence

from . import __version__
from .pipeline import NoDataError, run


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="report_automator",
        description="Merge messy CSV/Excel sales files into one clean Excel report.",
    )
    parser.add_argument("-i", "--input", required=True, type=Path, help="A file or a folder with .csv/.xlsx files")
    parser.add_argument("-o", "--output", type=Path, default=Path("sales_report.xlsx"), help="Report path (default: sales_report.xlsx)")
    parser.add_argument(
        "--month-first",
        action="store_true",
        help="Read ambiguous dates like 03/04/2026 as month/day (US). Default is day/month.",
    )
    parser.add_argument(
        "--keep-duplicates",
        action="store_true",
        help="Do not remove rows that are identical (same date, customer, product, quantity and amount).",
    )
    parser.add_argument("-q", "--quiet", action="store_true", help="Only print errors")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    log = (lambda message: None) if args.quiet else print
    try:
        log(f"Reading {args.input} ...")
        stats = run(args.input, args.output, dayfirst=not args.month_first, dedupe=not args.keep_duplicates, log=log)
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    except NoDataError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except PermissionError:
        print(f"Error: cannot write {args.output}. Is the file open in Excel?", file=sys.stderr)
        return 1

    log(
        f"\nDone. {stats.rows_kept} rows kept from {stats.files} file(s) "
        f"({stats.rejected} rejected, {stats.duplicates} duplicates removed)."
    )
    log(f"Report saved to {stats.output}")
    return 0
