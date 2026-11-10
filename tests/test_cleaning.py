import unittest
from datetime import date, datetime

import pandas as pd

from report_automator.cleaning import (
    clean_frame,
    normalize_column_name,
    parse_amount,
    parse_date,
    remove_duplicates,
)


class ParseAmountTests(unittest.TestCase):
    def test_plain_numbers(self):
        self.assertEqual(parse_amount(10), 10.0)
        self.assertEqual(parse_amount(12.5), 12.5)
        self.assertEqual(parse_amount("99.90"), 99.9)

    def test_us_and_brazilian_formats(self):
        self.assertEqual(parse_amount("1,234.56"), 1234.56)
        self.assertEqual(parse_amount("1.234,56"), 1234.56)
        self.assertEqual(parse_amount("R$ 1.234,56"), 1234.56)
        self.assertEqual(parse_amount("$ 99.90"), 99.9)
        self.assertEqual(parse_amount("12,5"), 12.5)
        self.assertEqual(parse_amount("1.234.567"), 1234567.0)

    def test_negative_values(self):
        self.assertEqual(parse_amount("-12,50"), -12.5)
        self.assertEqual(parse_amount("(45.00)"), -45.0)

    def test_invalid_values_return_none(self):
        for value in (None, "", "   ", "abc", "a combinar", "--", "12-34", float("nan"), True):
            with self.subTest(value=value):
                self.assertIsNone(parse_amount(value))


class ParseDateTests(unittest.TestCase):
    def test_text_formats(self):
        self.assertEqual(parse_date("2026-03-05"), date(2026, 3, 5))
        self.assertEqual(parse_date("05/03/2026"), date(2026, 3, 5))
        self.assertEqual(parse_date("05-03-2026"), date(2026, 3, 5))
        self.assertEqual(parse_date("Mar 5, 2026"), date(2026, 3, 5))
        self.assertEqual(parse_date("2026-03-05T10:30:00"), date(2026, 3, 5))

    def test_month_first_option(self):
        self.assertEqual(parse_date("03/05/2026", dayfirst=False), date(2026, 3, 5))
        self.assertEqual(parse_date("03/05/2026", dayfirst=True), date(2026, 5, 3))

    def test_real_date_objects_and_excel_serials(self):
        self.assertEqual(parse_date(datetime(2026, 3, 5, 14, 0)), date(2026, 3, 5))
        self.assertEqual(parse_date(pd.Timestamp("2026-03-05")), date(2026, 3, 5))
        self.assertEqual(parse_date(46086), date(2026, 3, 5))

    def test_invalid_dates_return_none(self):
        for value in (None, "", "31/02/2026", "not a date", "1850-01-01", 5):
            with self.subTest(value=value):
                self.assertIsNone(parse_date(value))


class ColumnNameTests(unittest.TestCase):
    def test_normalization(self):
        self.assertEqual(normalize_column_name("  Valor Total "), "valor_total")
        self.assertEqual(normalize_column_name("Preço"), "preco")
        self.assertEqual(normalize_column_name("Order-Date"), "order_date")


class CleanFrameTests(unittest.TestCase):
    def test_aliases_defaults_and_rejections(self):
        raw = pd.DataFrame(
            {
                "Data": ["01/02/2026", "", "xx", "03/02/2026", "04/02/2026"],
                "Cliente": ["  ana   silva ", "", "Bob", None, "Cy"],
                "Categoria": ["office", "", "home", "", "HOME"],
                "Valor": ["R$ 10,00", "", "5", "1.000,50", "abc"],
            }
        )
        result = clean_frame(raw, "test.csv")
        self.assertEqual(result.rows_read, 4)  # fully empty row is ignored
        self.assertEqual(len(result.frame), 2)
        self.assertEqual(len(result.issues), 2)
        first = result.frame.iloc[0]
        self.assertEqual(first["customer"], "ana silva")
        self.assertEqual(first["category"], "Office")
        self.assertEqual(first["amount"], 10.0)
        second = result.frame.iloc[1]
        self.assertEqual(second["customer"], "Unknown")
        self.assertEqual(second["category"], "Uncategorized")
        self.assertEqual(second["amount"], 1000.5)
        rows = {issue.row for issue in result.issues}
        self.assertEqual(rows, {4, 6})  # spreadsheet row numbers (header is row 1)

    def test_missing_required_column_is_reported_not_raised(self):
        result = clean_frame(pd.DataFrame({"foo": [1], "bar": [2]}), "bad.csv")
        self.assertTrue(result.frame.empty)
        self.assertIn("Missing required column", result.issues[0].reason)

    def test_empty_table(self):
        result = clean_frame(pd.DataFrame(), "empty.csv")
        self.assertTrue(result.frame.empty)
        self.assertEqual(len(result.issues), 1)

    def test_remove_duplicates(self):
        raw = pd.DataFrame(
            {
                "date": ["2026-01-01", "2026-01-01", "2026-01-02"],
                "customer": ["A", "A", "A"],
                "amount": [10, 10, 10],
            }
        )
        frame = clean_frame(raw, "x.csv").frame
        deduped, issues = remove_duplicates(frame)
        self.assertEqual(len(deduped), 2)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].row, 3)


if __name__ == "__main__":
    unittest.main()
