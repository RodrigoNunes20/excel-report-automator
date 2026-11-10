"""Build the Excel report. Summary figures are live formulas that point at the Data sheet."""
from __future__ import annotations

from datetime import datetime
from typing import Optional, Sequence

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .cleaning import Issue

FONT = "Arial"
HEADER_FILL = PatternFill("solid", start_color="1F3864")
HEADER_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
BASE_FONT = Font(name=FONT, size=10)
BOLD_FONT = Font(name=FONT, size=10, bold=True)
TITLE_FONT = Font(name=FONT, size=16, bold=True)

# Column letters on the Data sheet.
DATE, MONTH, CUSTOMER, CATEGORY, PRODUCT, QTY, AMOUNT, SOURCE, SRC_ROW = "ABCDEFGHI"
DATA_HEADERS = ["Date", "Month", "Customer", "Category", "Product", "Quantity", "Amount", "Source file", "Source row"]


def _style_header(ws, row: int, first_col: int, last_col: int) -> None:
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _data_range(column: str, last_row: int) -> str:
    return f"Data!${column}$2:${column}${last_row}"


def _write_data_sheet(ws, frame: pd.DataFrame) -> int:
    ws.append(DATA_HEADERS)
    _style_header(ws, 1, 1, len(DATA_HEADERS))
    for row in frame.itertuples(index=False):
        quantity = None if pd.isna(row.quantity) else float(row.quantity)
        ws.append(
            [
                row.date,
                row.date.strftime("%Y-%m"),
                row.customer,
                row.category,
                row.product,
                quantity,
                float(row.amount),
                row.source,
                int(row.source_row),
            ]
        )
    last_row = len(frame) + 1
    for r in range(2, last_row + 1):
        ws[f"{DATE}{r}"].number_format = "yyyy-mm-dd"
        ws[f"{QTY}{r}"].number_format = "#,##0.##"
        ws[f"{AMOUNT}{r}"].number_format = "#,##0.00"
        for c in range(1, len(DATA_HEADERS) + 1):
            ws.cell(row=r, column=c).font = BASE_FONT
    for letter, width in zip("ABCDEFGHI", (12, 10, 26, 18, 28, 10, 14, 30, 11)):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{SRC_ROW}{last_row}"
    return last_row


def _write_issues_sheet(ws, issues: Sequence[Issue]) -> None:
    ws.append(["Source", "Row", "Reason", "Raw data"])
    _style_header(ws, 1, 1, 4)
    if not issues:
        ws.append(["-", None, "No issues found", ""])
    for issue in issues:
        ws.append([issue.source, issue.row, issue.reason, issue.raw])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = BASE_FONT
    for letter, width in zip("ABCD", (30, 8, 60, 70)):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"


def _write_summary_sheet(ws, frame: pd.DataFrame, last_row: int, generated_at: datetime, files: int) -> None:
    amount, date_col = _data_range(AMOUNT, last_row), _data_range(DATE, last_row)

    ws["A1"] = "Sales Report"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated {generated_at:%Y-%m-%d %H:%M} from {files} source file(s). Figures below are formulas based on the Data sheet."
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")

    # Category and month tables (labels are written by Python, numbers are formulas).
    categories = frame.groupby("category")["amount"].sum().sort_values(ascending=False).index.tolist()
    months = sorted(frame["date"].map(lambda d: d.strftime("%Y-%m")).unique())
    cat_first, month_first = 14, 14
    cat_last = cat_first + len(categories) - 1
    month_last = month_first + len(months) - 1

    kpis = [
        ("Total revenue", f"=SUM({amount})", "#,##0.00"),
        ("Transactions", f"=COUNT({amount})", "#,##0"),
        ("Average ticket", "=IFERROR(B5/B6,0)", "#,##0.00"),
        ("First sale date", f"=MIN({date_col})", "yyyy-mm-dd"),
        ("Last sale date", f"=MAX({date_col})", "yyyy-mm-dd"),
        ("Top category", f'=IFERROR(INDEX($A${cat_first}:$A${cat_last},MATCH(MAX($B${cat_first}:$B${cat_last}),$B${cat_first}:$B${cat_last},0)),"-")', "General"),
    ]
    for offset, (label, formula, fmt) in enumerate(kpis):
        row = 5 + offset
        ws[f"A{row}"] = label
        ws[f"A{row}"].font = BOLD_FONT
        ws[f"B{row}"] = formula
        ws[f"B{row}"].font = BASE_FONT
        ws[f"B{row}"].number_format = fmt
        ws[f"B{row}"].alignment = Alignment(horizontal="right")

    ws["A12"], ws["E12"] = "Revenue by category", "Revenue by month"
    ws["A12"].font = ws["E12"].font = Font(name=FONT, size=12, bold=True)
    for col, text in zip("ABC", ("Category", "Revenue", "Share")):
        ws[f"{col}13"] = text
    for col, text in zip("EFG", ("Month", "Revenue", "Transactions")):
        ws[f"{col}13"] = text
    _style_header(ws, 13, 1, 3)
    _style_header(ws, 13, 5, 7)

    cat_range = _data_range(CATEGORY, last_row)
    month_range = _data_range(MONTH, last_row)
    for i, category in enumerate(categories):
        r = cat_first + i
        ws[f"A{r}"] = category
        ws[f"B{r}"] = f"=SUMIFS({amount},{cat_range},$A{r})"
        ws[f"C{r}"] = f"=IFERROR(B{r}/$B$5,0)"
        ws[f"B{r}"].number_format = "#,##0.00"
        ws[f"C{r}"].number_format = "0.0%"
    for i, month in enumerate(months):
        r = month_first + i
        ws[f"E{r}"] = month
        ws[f"F{r}"] = f"=SUMIFS({amount},{month_range},$E{r})"
        ws[f"G{r}"] = f"=COUNTIFS({month_range},$E{r})"
        ws[f"F{r}"].number_format = "#,##0.00"
        ws[f"G{r}"].number_format = "#,##0"
    for row in ws.iter_rows(min_row=13, max_row=max(cat_last, month_last)):
        for cell in row:
            if cell.row > 13:
                cell.font = BASE_FONT

    anchor_row = max(cat_last, month_last) + 3
    category_chart = BarChart()
    category_chart.type = "bar"
    category_chart.title = "Revenue by category"
    category_chart.add_data(Reference(ws, min_col=2, min_row=13, max_row=cat_last), titles_from_data=True)
    category_chart.set_categories(Reference(ws, min_col=1, min_row=cat_first, max_row=cat_last))
    category_chart.legend = None
    category_chart.x_axis.delete = False
    category_chart.y_axis.delete = False
    category_chart.height, category_chart.width = 8, 14
    ws.add_chart(category_chart, f"A{anchor_row}")

    month_chart = BarChart()
    month_chart.type = "col"
    month_chart.title = "Revenue by month"
    month_chart.add_data(Reference(ws, min_col=6, min_row=13, max_row=month_last), titles_from_data=True)
    month_chart.set_categories(Reference(ws, min_col=5, min_row=month_first, max_row=month_last))
    month_chart.legend = None
    month_chart.x_axis.delete = False
    month_chart.y_axis.delete = False
    month_chart.height, month_chart.width = 8, 14
    ws.add_chart(month_chart, f"E{anchor_row}")

    for letter, width in zip("ABCDEFG", (24, 16, 10, 4, 14, 16, 14)):
        ws.column_dimensions[letter].width = width


def build_workbook(
    frame: pd.DataFrame,
    issues: Sequence[Issue],
    files: int,
    generated_at: Optional[datetime] = None,
) -> Workbook:
    """Create the workbook with Summary, Data and Issues sheets."""
    if frame.empty:
        raise ValueError("Cannot build a report without any valid rows")
    generated_at = generated_at or datetime.now()
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    data = wb.create_sheet("Data")
    issues_ws = wb.create_sheet("Issues")
    last_row = _write_data_sheet(data, frame)
    _write_summary_sheet(summary, frame, last_row, generated_at, files)
    _write_issues_sheet(issues_ws, issues)
    return wb
