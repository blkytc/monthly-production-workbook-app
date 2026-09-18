"""Create a blank date-ranged workbook while preserving the source layout."""

from __future__ import annotations

import datetime as dt
import posixpath
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

from .sizing import column_width, parse_columns, parse_rows, row_height


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
SHEET_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
SHEET_CONTENT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

ET.register_namespace("", MAIN)
ET.register_namespace("r", DOC_REL)

MANUAL_CELLS = {
    "D9", "D10", "D11", "D13", "D14", "D15", "D16", "D18", "D19",
    "D20", "D22", "D23", "D24", "D25", "D26", "F8", "F12", "F17",
    "F21", "I14", "I15",
}


@dataclass(frozen=True)
class DimensionRule:
    selection: str
    value: float
    unit: str


@dataclass(frozen=True)
class GenerationResult:
    output_path: Path
    sheet_count: int
    sheet_names: tuple[str, ...]


def read_workbook_title(source: str | Path) -> str:
    source_path = Path(source).expanduser().resolve()
    with zipfile.ZipFile(source_path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        first_sheet = next(iter(workbook.find(q(MAIN, "sheets"))))
        rid = first_sheet.attrib[q(DOC_REL, "id")]
        relation = next(node for node in relationships if node.attrib["Id"] == rid)
        part = posixpath.normpath(posixpath.join("xl", relation.attrib["Target"])).lstrip("/")
        root = ET.fromstring(archive.read(part))
        title_cell = cells(root).get("A1")
        if title_cell is None:
            return ""
        inline = title_cell.find(f"{q(MAIN, 'is')}/{q(MAIN, 't')}")
        if inline is not None:
            return inline.text or ""
        value = title_cell.find(q(MAIN, "v"))
        if value is None or not value.text:
            return ""
        if title_cell.attrib.get("t") == "s":
            shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            item = list(shared)[int(value.text)]
            return "".join(node.text or "" for node in item.iter(q(MAIN, "t")))
        return value.text


def q(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def calendar_dates(start: dt.date, end: dt.date) -> list[dt.date]:
    if end < start:
        raise ValueError("结束日期不能早于开始日期")
    return [start + dt.timedelta(days=offset) for offset in range((end - start).days + 1)]


def sheet_names(dates: list[dt.date]) -> list[str]:
    candidates = [[str(value.day) for value in dates], [value.strftime("%m-%d") for value in dates]]
    for names in candidates:
        if len(names) == len(set(names)):
            return names
    return [value.isoformat() for value in dates]


def excel_serial(value: dt.date) -> int:
    return (value - dt.date(1899, 12, 30)).days


def cells(root: ET.Element) -> dict[str, ET.Element]:
    return {node.attrib["r"]: node for node in root.findall(f".//{q(MAIN, 'c')}")}


def remove_child(cell: ET.Element, tag: str) -> None:
    node = cell.find(q(MAIN, tag))
    if node is not None:
        cell.remove(node)


def clear_value(cell: ET.Element) -> None:
    remove_child(cell, "v")
    remove_child(cell, "is")
    cell.attrib.pop("t", None)


def set_number(cell: ET.Element, value: int | float) -> None:
    clear_value(cell)
    ET.SubElement(cell, q(MAIN, "v")).text = str(value)


def set_text(cell: ET.Element, value: str) -> None:
    clear_value(cell)
    cell.attrib["t"] = "inlineStr"
    inline = ET.SubElement(cell, q(MAIN, "is"))
    ET.SubElement(inline, q(MAIN, "t")).text = value


def set_formula(cell: ET.Element, formula: str) -> None:
    clear_value(cell)
    node = cell.find(q(MAIN, "f"))
    if node is None:
        node = ET.Element(q(MAIN, "f"))
        cell.insert(0, node)
    node.attrib.clear()
    node.text = formula


def cached_number(root: ET.Element, address: str) -> float:
    cell = cells(root).get(address)
    value = None if cell is None else cell.find(q(MAIN, "v"))
    return float(value.text) if value is not None and value.text else 0.0


def _stored_number(root: ET.Element, address: str) -> float:
    cell = cells(root).get(address)
    value = None if cell is None else cell.find(q(MAIN, "v"))
    try:
        return float(value.text) if value is not None and value.text else 0.0
    except ValueError:
        return 0.0


def _daily_totals(root: ET.Element) -> dict[int, float]:
    totals = {
        8: sum(_stored_number(root, address) for address in ("D13", "D18", "D22")),
        9: sum(_stored_number(root, address) for address in ("D9", "D14", "D23")),
        10: _stored_number(root, "D24"),
        11: sum(_stored_number(root, address) for address in ("D10", "D15", "D19", "D25")),
        12: sum(_stored_number(root, address) for address in ("D11", "D16", "D20", "D26")),
        14: _stored_number(root, "I14"),
        15: _stored_number(root, "I15"),
    }
    totals[7] = totals[8] + totals[9] + totals[10]
    totals[13] = totals[14] + totals[15]
    return totals


def _formula_baseline(root: ET.Element, row: int) -> float:
    cell = cells(root)[f"K{row}"]
    formula = cell.find(q(MAIN, "f"))
    if formula is None or not formula.text:
        return 0.0
    match = re.fullmatch(rf"I{row}(?:\+([-+0-9.eE]+))?", formula.text)
    return float(match.group(1)) if match and match.group(1) else 0.0


def annual_totals(source_sheets: list[tuple[dt.date, ET.Element]]) -> dict[int, float]:
    ordered = sorted(source_sheets, key=lambda item: item[0])
    first_date, first_root = ordered[0]
    totals = {
        row: 0.0 if first_date.month == 1 and first_date.day == 1 else _formula_baseline(first_root, row)
        for row in range(7, 16)
    }
    for index, (date_value, root) in enumerate(ordered):
        if index and date_value.month == 1 and date_value.day == 1:
            totals = {row: 0.0 for row in range(7, 16)}
        daily = _daily_totals(root)
        totals = {row: totals[row] + daily[row] for row in range(7, 16)}
    return totals


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def apply_row_rules(root: ET.Element, rules: list[DimensionRule]) -> None:
    row_nodes = {int(node.attrib["r"]): node for node in root.findall(f".//{q(MAIN, 'row')}")}
    for rule in rules:
        height = row_height(rule.value, rule.unit)
        for number in parse_rows(rule.selection):
            if number not in row_nodes:
                raise ValueError(f"模板中不存在第 {number} 行")
            row_nodes[number].attrib.update({"ht": f"{height:g}", "customHeight": "1"})


def _set_column(root: ET.Element, number: int, width: float) -> None:
    columns = root.find(q(MAIN, "cols"))
    if columns is None:
        sheet_data = root.find(q(MAIN, "sheetData"))
        columns = ET.Element(q(MAIN, "cols"))
        root.insert(list(root).index(sheet_data), columns)
    source = next(
        (node for node in columns if int(node.attrib["min"]) <= number <= int(node.attrib["max"])),
        None,
    )
    base = {} if source is None else dict(source.attrib)
    if source is not None:
        columns.remove(source)
        minimum, maximum = int(base["min"]), int(base["max"])
        if minimum < number:
            ET.SubElement(columns, q(MAIN, "col"), {**base, "min": str(minimum), "max": str(number - 1)})
        if number < maximum:
            ET.SubElement(columns, q(MAIN, "col"), {**base, "min": str(number + 1), "max": str(maximum)})
    base.update({"min": str(number), "max": str(number), "width": f"{width:g}", "customWidth": "1"})
    ET.SubElement(columns, q(MAIN, "col"), base)
    columns[:] = sorted(columns, key=lambda node: int(node.attrib["min"]))


def apply_column_rules(root: ET.Element, rules: list[DimensionRule]) -> None:
    for rule in rules:
        width = column_width(rule.value, rule.unit)
        for name in parse_columns(rule.selection):
            number = 0
            for character in name:
                number = number * 26 + ord(character) - ord("A") + 1
            _set_column(root, number, width)


def generate_workbook(
    source: str | Path,
    output: str | Path,
    start: dt.date,
    end: dt.date,
    title: str,
    row_rules: list[DimensionRule] | None = None,
    column_rules: list[DimensionRule] | None = None,
) -> GenerationResult:
    source_path = Path(source).expanduser().resolve()
    output_path = Path(output).expanduser().resolve()
    if not source_path.is_file() or source_path.suffix.lower() != ".xlsx":
        raise ValueError("请选择有效的 .xlsx 模板文件")
    if not title.strip():
        raise ValueError("表格标题不能为空")

    dates = calendar_dates(start, end)
    names = sheet_names(dates)
    row_rules = row_rules or []
    column_rules = column_rules or []

    with zipfile.ZipFile(source_path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}

    workbook = ET.fromstring(files["xl/workbook.xml"])
    relationships = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
    content_types = ET.fromstring(files["[Content_Types].xml"])
    sheets_node = workbook.find(q(MAIN, "sheets"))
    if sheets_node is None or not list(sheets_node):
        raise ValueError("工作簿没有工作表")

    rels = {node.attrib["Id"]: node for node in relationships.findall(q(PKG_REL, "Relationship"))}
    source_sheets: list[tuple[dt.date, ET.Element]] = []
    old_sheet_parts: set[str] = set()
    for sheet in list(sheets_node):
        rid = sheet.attrib[q(DOC_REL, "id")]
        relationship = rels.get(rid)
        if relationship is None or relationship.attrib.get("Type") != SHEET_TYPE:
            continue
        target = relationship.attrib["Target"]
        part = posixpath.normpath(posixpath.join("xl", target)).lstrip("/")
        old_sheet_parts.add(part)
        root = ET.fromstring(files[part])
        signature = cells(root)
        if not {"A1", "D4", "G4", "D8", "F7", "I7", "J7", "K7"}.issubset(signature):
            raise ValueError("模板不匹配：关键单元格缺失")
        serial = int(cached_number(root, "G4"))
        source_sheets.append((dt.date(1899, 12, 30) + dt.timedelta(days=serial), root))
    if not source_sheets:
        raise ValueError("模板不匹配：没有可用的日期工作表")

    annual_baseline = annual_totals(source_sheets)
    template_bytes = xml_bytes(source_sheets[0][1])

    for sheet in list(sheets_node):
        rid = sheet.attrib[q(DOC_REL, "id")]
        relationship = rels.get(rid)
        if relationship is not None and relationship.attrib.get("Type") == SHEET_TYPE:
            sheets_node.remove(sheet)
            relationships.remove(relationship)
    for override in list(content_types):
        if override.attrib.get("PartName", "").lstrip("/") in old_sheet_parts:
            content_types.remove(override)
    for part in old_sheet_parts:
        files.pop(part, None)

    used_rids = {
        int(match.group(1))
        for node in relationships.findall(q(PKG_REL, "Relationship"))
        if (match := re.fullmatch(r"rId(\d+)", node.attrib.get("Id", "")))
    }
    next_rid = max(used_rids, default=0) + 1
    previous_name: str | None = None

    for index, (date_value, name) in enumerate(zip(dates, names), start=1):
        root = ET.fromstring(template_bytes)
        by_address = cells(root)
        set_text(by_address["A1"], title.strip())
        clear_value(by_address["D4"])
        set_number(by_address["G4"], excel_serial(date_value))
        for address in MANUAL_CELLS:
            clear_value(by_address[address])

        for row in range(8, 27):
            target = by_address[f"E{row}"]
            if index == 1 and row == 17:
                remove_child(target, "f")
                clear_value(target)
            else:
                formula = f"D{row}" if index == 1 else f"D{row}+'{previous_name}'!E{row}"
                set_formula(target, formula)
        for row in range(7, 16):
            j_formula = f"I{row}" if index == 1 else f"I{row}+'{previous_name}'!J{row}"
            if index == 1:
                k_formula = f"I{row}+{annual_baseline[row]:g}"
            elif date_value.month == 1 and date_value.day == 1:
                k_formula = f"I{row}"
            else:
                k_formula = f"I{row}+'{previous_name}'!K{row}"
            set_formula(by_address[f"J{row}"], j_formula)
            set_formula(by_address[f"K{row}"], k_formula)

        apply_row_rules(root, row_rules)
        apply_column_rules(root, column_rules)
        for formula_cell in root.findall(f".//{q(MAIN, 'c')}"):
            if formula_cell.find(q(MAIN, "f")) is not None:
                remove_child(formula_cell, "v")

        part = f"xl/worksheets/sheet{index}.xml"
        rid = f"rId{next_rid}"
        next_rid += 1
        files[part] = xml_bytes(root)
        ET.SubElement(sheets_node, q(MAIN, "sheet"), {
            "name": name, "sheetId": str(index), q(DOC_REL, "id"): rid,
        })
        ET.SubElement(relationships, q(PKG_REL, "Relationship"), {
            "Id": rid, "Type": SHEET_TYPE, "Target": f"worksheets/sheet{index}.xml",
        })
        ET.SubElement(content_types, q(CONTENT, "Override"), {
            "PartName": f"/xl/worksheets/sheet{index}.xml", "ContentType": SHEET_CONTENT,
        })
        previous_name = name

    calc = workbook.find(q(MAIN, "calcPr"))
    if calc is None:
        calc = ET.SubElement(workbook, q(MAIN, "calcPr"))
    calc.attrib.update({"calcMode": "auto", "fullCalcOnLoad": "1", "forceFullCalc": "1"})
    files["xl/workbook.xml"] = xml_bytes(workbook)
    files["xl/_rels/workbook.xml.rels"] = xml_bytes(relationships)
    files["[Content_Types].xml"] = xml_bytes(content_types)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=output_path.parent, delete=False) as handle:
        temporary = Path(handle.name)
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, data in files.items():
                archive.writestr(name, data)
        temporary.replace(output_path)
    finally:
        temporary.unlink(missing_ok=True)
    return GenerationResult(output_path, len(dates), tuple(names))
