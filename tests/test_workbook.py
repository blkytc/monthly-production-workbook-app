import datetime as dt
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from monthly_workbook.workbook import DimensionRule, generate_workbook, read_workbook_title


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
SOURCE = Path("/Volumes/Seagate/关于/outputs/monthly-production-workbook/生产、入库、发运表2026-10-25-2026-11-24.xlsx")


def q(namespace, name):
    return f"{{{namespace}}}{name}"


def read_output(path):
    archive = zipfile.ZipFile(path)
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rels = {node.attrib["Id"]: node.attrib["Target"] for node in rels_root}
    sheets = []
    for sheet in workbook.find(q(MAIN, "sheets")):
        target = rels[sheet.attrib[q(DOC_REL, "id")]]
        part = "xl/" + target.lstrip("/")
        sheets.append((sheet.attrib["name"], ET.fromstring(archive.read(part))))
    return archive, sheets


def cell(root, address):
    return root.find(f".//{q(MAIN, 'c')}[@r='{address}']")


class WorkbookGenerationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not SOURCE.exists():
            raise unittest.SkipTest(f"缺少测试模板：{SOURCE}")

    def test_generates_arbitrary_range_title_formulas_and_sizes(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "result.xlsx"
            result = generate_workbook(
                SOURCE,
                output,
                dt.date(2026, 10, 3),
                dt.date(2026, 10, 5),
                "测试单位生产记录",
                [DimensionRule("4,6-7", 2.54, "厘米")],
                [DimensionRule("A:C", 20, "字符"), DimensionRule("B", 2.54, "厘米")],
            )

            self.assertEqual(result.sheet_count, 3)
            archive, sheets = read_output(output)
            self.addCleanup(archive.close)
            self.assertEqual([name for name, _ in sheets], ["3", "4", "5"])

            first, second = sheets[0][1], sheets[1][1]
            self.assertEqual(int(cell(first, "G4").find(q(MAIN, "v")).text), 46298)
            self.assertEqual(cell(first, "A1").find(q(MAIN, "is")).find(q(MAIN, "t")).text, "测试单位生产记录")
            self.assertEqual(cell(first, "E8").find(q(MAIN, "f")).text, "D8")
            self.assertIsNone(cell(first, "E17").find(q(MAIN, "f")))
            self.assertEqual(cell(first, "J7").find(q(MAIN, "f")).text, "I7")
            self.assertEqual(cell(first, "K7").find(q(MAIN, "f")).text, "I7+4.37501e+06")
            self.assertEqual(cell(second, "E8").find(q(MAIN, "f")).text, "D8+'3'!E8")
            self.assertEqual(cell(second, "J7").find(q(MAIN, "f")).text, "I7+'3'!J7")
            self.assertIsNone(cell(first, "D9").find(q(MAIN, "v")))

            rows = {row.attrib["r"]: row for row in first.findall(f".//{q(MAIN, 'row')}")}
            self.assertEqual(rows["4"].attrib["ht"], "72")
            self.assertEqual(rows["6"].attrib["ht"], "72")
            self.assertEqual(rows["7"].attrib["ht"], "72")

            columns = {}
            for node in first.find(q(MAIN, "cols")):
                if node.attrib["min"] == node.attrib["max"]:
                    columns[int(node.attrib["min"])] = float(node.attrib["width"])
            self.assertAlmostEqual(columns[1], 20)
            self.assertAlmostEqual(columns[2], 13)
            self.assertAlmostEqual(columns[3], 20)

    def test_rejects_invalid_date_range(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(ValueError, "结束日期"):
                generate_workbook(
                    SOURCE,
                    Path(directory) / "result.xlsx",
                    dt.date(2026, 10, 5),
                    dt.date(2026, 10, 4),
                    "标题",
                )

    def test_reads_title_from_source_workbook(self):
        self.assertEqual(read_workbook_title(SOURCE), "开阳磷矿生产、入库、发运记录")

    def test_first_sheet_carries_annual_total_even_when_year_differs(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "next-year.xlsx"
            generate_workbook(
                SOURCE,
                output,
                dt.date(2027, 2, 25),
                dt.date(2027, 2, 25),
                "测试单位生产记录",
            )
            archive, sheets = read_output(output)
            self.addCleanup(archive.close)
            formula = cell(sheets[0][1], "K7").find(q(MAIN, "f")).text
            self.assertEqual(formula, "I7+4.37501e+06")


if __name__ == "__main__":
    unittest.main()
