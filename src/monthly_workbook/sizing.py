"""Parse row/column selections and convert display units to Excel units."""

from __future__ import annotations

import re


PHYSICAL_TO_INCHES = {"英寸": 1.0, "厘米": 1 / 2.54, "毫米": 1 / 25.4}


def _unique(values: list[int]) -> list[int]:
    return list(dict.fromkeys(values))


def parse_rows(text: str) -> list[int]:
    values: list[int] = []
    try:
        for token in _tokens(text):
            if "-" in token:
                start_text, end_text = token.split("-", 1)
                start, end = int(start_text), int(end_text)
                if start < 1 or end < start:
                    raise ValueError
                values.extend(range(start, end + 1))
            else:
                value = int(token)
                if value < 1:
                    raise ValueError
                values.append(value)
    except (TypeError, ValueError):
        raise ValueError("行范围格式不正确，例如 4-8,12,15-18") from None
    return _unique(values)


def _column_number(value: str) -> int:
    if not re.fullmatch(r"[A-Z]{1,3}", value):
        raise ValueError
    number = 0
    for character in value:
        number = number * 26 + ord(character) - ord("A") + 1
    if number > 16384:
        raise ValueError
    return number


def column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def parse_columns(text: str) -> list[str]:
    values: list[int] = []
    try:
        for token in _tokens(text.upper()):
            if ":" in token:
                start_text, end_text = token.split(":", 1)
                start, end = _column_number(start_text), _column_number(end_text)
                if end < start:
                    raise ValueError
                values.extend(range(start, end + 1))
            else:
                values.append(_column_number(token))
    except (TypeError, ValueError):
        raise ValueError("列范围格式不正确，例如 A:C,F,H:J") from None
    return [column_name(value) for value in _unique(values)]


def _tokens(text: str) -> list[str]:
    tokens = [item.strip() for item in text.replace("，", ",").split(",")]
    if not tokens or any(not item for item in tokens):
        raise ValueError
    return tokens


def _positive(value: float) -> float:
    number = float(value)
    if number <= 0:
        raise ValueError("尺寸必须大于 0")
    return number


def row_height(value: float, unit: str) -> float:
    number = _positive(value)
    if unit == "磅":
        return number
    if unit == "字符":
        return number * 15
    if unit in PHYSICAL_TO_INCHES:
        return number * PHYSICAL_TO_INCHES[unit] * 72
    raise ValueError(f"不支持的单位：{unit}")


def column_width(value: float, unit: str) -> float:
    number = _positive(value)
    if unit == "字符":
        return number
    if unit == "磅":
        pixels = number * 96 / 72
    elif unit in PHYSICAL_TO_INCHES:
        pixels = number * PHYSICAL_TO_INCHES[unit] * 96
    else:
        raise ValueError(f"不支持的单位：{unit}")
    return max(0, (pixels - 5) / 7)
