#!/usr/bin/env python3
"""Inspect the end-to-end sample workbook produced during release verification."""

from __future__ import annotations

import datetime as dt
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def q(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def main(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        self_test = archive.testzip()
        assert self_test is None, f"ZIP 条目损坏：{self_test}"
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationships = {node.attrib["Id"]: node.attrib["Target"] for node in rel_root}
        sheets = list(workbook.find(q(MAIN, "sheets")))
        assert [node.attrib["name"] for node in sheets] == ["28", "29", "30", "31", "1", "2"]
        for offset, sheet in enumerate(sheets):
            target = relationships[sheet.attrib[q(DOC_REL, "id")]]
            root = ET.fromstring(archive.read("xl/" + target))
            by_address = {node.attrib["r"]: node for node in root.findall(f".//{q(MAIN, 'c')}")}
            expected = dt.date(2026, 12, 28) + dt.timedelta(days=offset)
            actual = dt.date(1899, 12, 30) + dt.timedelta(days=int(by_address["G4"].find(q(MAIN, "v")).text))
            assert actual == expected
            title = by_address["A1"].find(f"{q(MAIN, 'is')}/{q(MAIN, 't')}").text
            assert title == "测试企业生产、入库、发运记录"
            assert by_address["D9"].find(q(MAIN, "v")) is None
            formula = by_address["E8"].find(q(MAIN, "f")).text
            assert formula == ("D8" if offset == 0 else f"D8+'{sheets[offset - 1].attrib['name']}'!E8")
        print(f"检查通过：{path}，共 {len(sheets)} 张工作表")


if __name__ == "__main__":
    main(Path(sys.argv[1]))
