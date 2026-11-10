import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from report_automator.cli import main
from report_automator.pipeline import NoDataError, run


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def _write_inputs(self):
        (self.dir / "a.csv").write_text(
            "Date;Customer;Category;Amount\n"
            "01/01/2026;Ana;Food;R$ 1.000,50\n"
            "02/01/2026;Bob;Food;200,00\n"
            "02/01/2026;Bob;Food;200,00\n"  # duplicate
            "bad-date;Cy;Food;10\n",
            encoding="utf-8",
        )
        with pd.ExcelWriter(self.dir / "b.xlsx") as writer:
            pd.DataFrame(
                {"data": [pd.Timestamp("2026-02-03")], "cliente": ["Dee"], "categoria": ["Home"], "valor": [300.0]}
            ).to_excel(writer, sheet_name="Sheet A", index=False)
        (self.dir / "ignored.txt").write_text("not a spreadsheet")

    def test_full_run_creates_report(self):
        self._write_inputs()
        out = self.dir / "out" / "report.xlsx"
        stats = run(self.dir, out)
        self.assertTrue(out.exists())
        self.assertEqual(stats.files, 2)
        self.assertEqual(stats.rows_kept, 3)
        self.assertEqual(stats.rejected, 1)
        self.assertEqual(stats.duplicates, 1)

        wb = openpyxl.load_workbook(out)
        self.assertEqual(wb.sheetnames, ["Summary", "Data", "Issues"])
        self.assertEqual(wb["Data"].max_row, 4)  # header + 3 rows
        self.assertEqual(wb["Summary"]["B5"].value, "=SUM(Data!$G$2:$G$4)")
        amounts = sorted(wb["Data"][f"G{r}"].value for r in range(2, 5))
        self.assertEqual(amounts, [200.0, 300.0, 1000.5])
        issue_reasons = [row[2] for row in wb["Issues"].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(len(issue_reasons), 2)

    def test_keep_duplicates(self):
        self._write_inputs()
        stats = run(self.dir, self.dir / "r.xlsx", dedupe=False)
        self.assertEqual(stats.rows_kept, 4)

    def test_no_valid_rows_raises(self):
        (self.dir / "x.csv").write_text("Date,Amount\nnope,nope\n", encoding="utf-8")
        with self.assertRaises(NoDataError):
            run(self.dir, self.dir / "r.xlsx")

    def test_corrupt_excel_file_does_not_stop_the_run(self):
        self._write_inputs()
        (self.dir / "broken.xlsx").write_text("this is not really an xlsx file")
        stats = run(self.dir, self.dir / "r.xlsx")
        self.assertEqual(stats.rows_kept, 3)
        wb = openpyxl.load_workbook(self.dir / "r.xlsx")
        reasons = [row[2] for row in wb["Issues"].iter_rows(min_row=2, values_only=True)]
        self.assertTrue(any("could not be read" in str(r) for r in reasons))

    def test_cli_exit_codes(self):
        self._write_inputs()
        self.assertEqual(main(["-i", str(self.dir), "-o", str(self.dir / "ok.xlsx"), "-q"]), 0)
        self.assertEqual(main(["-i", str(self.dir / "missing"), "-q"]), 2)


if __name__ == "__main__":
    unittest.main()
