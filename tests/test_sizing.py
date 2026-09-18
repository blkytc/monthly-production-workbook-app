import unittest

from monthly_workbook.sizing import (
    column_width,
    parse_columns,
    parse_rows,
    row_height,
)


class RangeParsingTests(unittest.TestCase):
    def test_parse_rows_accepts_lists_and_ranges(self):
        self.assertEqual(parse_rows("4-6, 9, 11-12"), [4, 5, 6, 9, 11, 12])

    def test_parse_columns_accepts_lists_and_ranges(self):
        self.assertEqual(parse_columns("A:C, F, H:I"), ["A", "B", "C", "F", "H", "I"])

    def test_duplicate_members_are_returned_once(self):
        self.assertEqual(parse_rows("1-3,2"), [1, 2, 3])
        self.assertEqual(parse_columns("A:C,B"), ["A", "B", "C"])

    def test_invalid_ranges_raise_chinese_error(self):
        with self.assertRaisesRegex(ValueError, "行范围"):
            parse_rows("3-1")
        with self.assertRaisesRegex(ValueError, "列范围"):
            parse_columns("A:1")


class UnitConversionTests(unittest.TestCase):
    def test_row_height_converts_physical_units_to_points(self):
        self.assertAlmostEqual(row_height(1, "英寸"), 72)
        self.assertAlmostEqual(row_height(2.54, "厘米"), 72)
        self.assertAlmostEqual(row_height(25.4, "毫米"), 72)
        self.assertAlmostEqual(row_height(18, "磅"), 18)

    def test_column_width_keeps_character_units(self):
        self.assertEqual(column_width(38.54, "字符"), 38.54)

    def test_column_width_converts_physical_units(self):
        self.assertAlmostEqual(column_width(1, "英寸"), 13, places=2)

    def test_non_positive_size_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "大于 0"):
            row_height(0, "厘米")


if __name__ == "__main__":
    unittest.main()
