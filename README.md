# Excel Report Automator

Turn a folder of messy sales spreadsheets (CSV and Excel) into **one clean Excel report** with a summary, charts, the cleaned data, and a log of every row that was rejected or removed.

Built with Python, pandas and openpyxl. Runs from the command line, no setup beyond `pip install`.

## The problem it solves

Sales data usually arrives from several sources, each with its own format: different column names, `;` or `,` delimiters, `1.234,56` vs `1,234.56`, `31/01/2026` vs `2026-01-31`, blank lines, duplicates and typos. Cleaning that by hand every month wastes hours and introduces mistakes. This tool does it in seconds and **tells you exactly what it dropped and why**.

## What the report contains

| Sheet | Content |
|---|---|
| **Summary** | Total revenue, transactions, average ticket, date range, top category, revenue by category and by month, two charts. All figures are **live Excel formulas** pointing at the Data sheet. |
| **Data** | The cleaned, merged table (filterable, frozen header) with the source file and original row number of every record. |
| **Issues** | Every rejected row (invalid date, invalid amount, unreadable file) and every duplicate removed, with the raw values. |

## Quick start

```bash
git clone https://github.com/<your-user>/excel-report-automator.git
cd excel-report-automator
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# try it with the included messy sample files
python -m report_automator -i sample_data -o sales_report.xlsx
```

Output:

```
Reading sample_data ...
  - sales_2026_01.csv: 41 valid rows, 0 issue(s)
  - sales_2026_02.csv: 38 valid rows, 2 issue(s)
  - sales_2026_03.xlsx [North]: 30 valid rows, 0 issue(s)
  - sales_2026_03.xlsx [South]: 30 valid rows, 0 issue(s)

Done. 138 rows kept from 3 file(s) (2 rejected, 1 duplicates removed).
Report saved to sales_report.xlsx
```

Use it on your own files: `python -m report_automator -i path/to/folder -o report.xlsx`

### Options

| Option | Description |
|---|---|
| `-i, --input` | A file or a folder with `.csv` / `.xlsx` files (required) |
| `-o, --output` | Report path (default `sales_report.xlsx`) |
| `--month-first` | Read ambiguous dates like `03/04/2026` as month/day (default is day/month) |
| `--keep-duplicates` | Keep rows that are identical instead of removing them |
| `-q, --quiet` | Only print errors |

Exit codes: `0` success, `1` no valid data or output not writable, `2` input not found.

## What the cleaning does

- **Column names**: case, spaces and accents are ignored; common English and Portuguese names are mapped automatically (`Data`, `Order Date` → date; `Valor`, `Total` → amount; `Cliente` → customer; `Qtd` → quantity...). Only **date** and **amount** are required.
- **Numbers**: understands `1234.5`, `1,234.56`, `1.234,56`, `R$ 1.234,56`, `$99.90`, `(45.00)`, `-12,5`. A single `,` or `.` is read as the decimal mark; repeated ones as thousands marks.
- **Dates**: real Excel dates, Excel serial numbers, ISO (`2026-03-05`), `05/03/2026`, `05-03-2026`, `Mar 5, 2026` and more. Impossible dates (`31/02/2026`) are rejected, not guessed.
- **Text**: extra spaces trimmed; missing customer becomes `Unknown`, missing category `Uncategorized`; categories are title-cased so `home` and `HOME` are the same group.
- **CSV files**: delimiter detected automatically (`,` `;` tab), UTF-8 and Windows-1252 encodings.
- **Excel files**: every sheet is read.
- **Duplicates**: rows with the same date, customer, category, product, quantity and amount are removed (first one is kept) and logged. Use `--keep-duplicates` if repeated identical sales are legitimate in your data.
- **Bad files never stop the run**: a corrupt file is logged in the Issues sheet and the rest is processed.

## Project structure

```
report_automator/
  cleaning.py    # parsing and normalization (pure functions, fully tested)
  loader.py      # finds and reads CSV / Excel files
  report.py      # builds the Excel workbook (formulas + charts)
  pipeline.py    # load -> clean -> de-duplicate -> report
  cli.py         # command line interface
scripts/make_sample_data.py   # regenerates the messy sample files
sample_data/                  # example input files
tests/                        # unit and end-to-end tests
```

## Tests

```bash
python -m unittest discover -s tests -v
```

The suite covers number and date parsing, column aliasing, rejected rows, duplicates, corrupt files, the CLI exit codes and a full end-to-end run that checks the generated workbook.

## Ideas for extension

- Add your own column aliases in `COLUMN_ALIASES` (`report_automator/cleaning.py`).
- Add more report sheets (per customer, per product) in `report.py`.
- Schedule it (cron / Task Scheduler) to build the report automatically every week.

## License

MIT
