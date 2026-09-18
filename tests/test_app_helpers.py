import datetime as dt
import unittest
from pathlib import Path

from monthly_workbook.app import (
    default_output_path,
    normalize_filename,
    parse_user_date,
    update_filename_dates,
)


class AppHelperTests(unittest.TestCase):
    def test_parse_user_date_accepts_slashes_and_hyphens(self):
        expected = dt.date(2026, 10, 25)
        self.assertEqual(parse_user_date("2026/10/25"), expected)
        self.assertEqual(parse_user_date("2026-10-25"), expected)

    def test_parse_user_date_reports_expected_format(self):
        with self.assertRaisesRegex(ValueError, "年/月/日"):
            parse_user_date("10/25")

    def test_default_output_path_uses_source_directory_and_dates(self):
        value = default_output_path(
            Path("/tmp/旧表.xlsx"), dt.date(2026, 10, 25), dt.date(2026, 11, 24)
        )
        self.assertEqual(value, Path("/tmp/生产、入库、发运表2026-10-25-2026-11-24.xlsx"))

    def test_filename_gets_xlsx_extension(self):
        self.assertEqual(normalize_filename("十月份生产表"), "十月份生产表.xlsx")
        self.assertEqual(normalize_filename("十月份生产表.xlsx"), "十月份生产表.xlsx")

    def test_filename_rejects_path_and_windows_reserved_characters(self):
        for value in ("子目录/生产表", "生产:表", "生产?.xlsx"):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "文件名称"):
                normalize_filename(value)

    def test_filename_dates_replace_existing_range(self):
        value = update_filename_dates(
            "大水矿石产消存计划一览表（9月）-2026-09-17-2026-10-17.xlsx",
            dt.date(2026, 9, 25),
            dt.date(2026, 10, 24),
        )
        self.assertEqual(
            value,
            "大水矿石产消存计划一览表（9月）-2026-09-25-2026-10-24.xlsx",
        )

    def test_filename_dates_preserve_custom_name(self):
        value = update_filename_dates(
            "我的月度报表-2026-09-17-2026-10-17.xlsx",
            dt.date(2026, 10, 25),
            dt.date(2026, 11, 24),
        )
        self.assertEqual(value, "我的月度报表-2026-10-25-2026-11-24.xlsx")

    def test_filename_dates_append_when_range_is_missing(self):
        value = update_filename_dates(
            "我的月度报表.xlsx",
            dt.date(2026, 10, 25),
            dt.date(2026, 11, 24),
        )
        self.assertEqual(value, "我的月度报表-2026-10-25-2026-11-24.xlsx")


if __name__ == "__main__":
    unittest.main()
