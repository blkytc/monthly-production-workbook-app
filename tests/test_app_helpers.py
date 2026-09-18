import datetime as dt
import unittest
from pathlib import Path
from unittest.mock import Mock

from monthly_workbook.app import (
    apply_app_theme,
    default_output_path,
    normalize_filename,
    parse_user_date,
    update_filename_dates,
)


class RecordingStyle:
    def __init__(self):
        self.theme = None
        self.configured = {}
        self.mapped = {}

    def theme_names(self):
        return ("clam", "default")

    def theme_use(self, value):
        self.theme = value

    def configure(self, name, **values):
        self.configured[name] = values

    def map(self, name, **values):
        self.mapped[name] = values


class AppThemeTests(unittest.TestCase):
    def test_dark_theme_configures_portable_controls_and_selection_states(self):
        style = RecordingStyle()
        root = Mock()

        apply_app_theme(style, root)

        self.assertEqual(style.theme, "clam")
        self.assertEqual(root.configure.call_args.kwargs["background"], "#17191c")
        self.assertEqual(style.configured["TFrame"]["background"], "#17191c")
        self.assertEqual(style.configured["TEntry"]["fieldbackground"], "#2b2f34")
        self.assertEqual(style.configured["Treeview"]["background"], "#202327")
        self.assertIn("background", style.mapped["Treeview"])
        self.assertIn("background", style.mapped["TButton"])


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
