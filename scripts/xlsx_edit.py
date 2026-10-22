#!/usr/bin/env python3
"""Declarative XLSX edit helper for the dsh-doc-review plugin.

The plugin submits annotations against a workbook; the agent turns them into a
small ops list and runs this helper once. The helper performs preflight,
backup, matching, edit, read-back verification, zip-part parity check and
atomic replacement, and records failures in the plugin log.

Usage:
  xlsx_edit.py --file /path/book.xlsx --ops '[{"op":"set_cell","sheet":"数据","cell":"B2","value":"999","expect":"100"}]'
  xlsx_edit.py --file /path/book.xlsx --ops ops.json --dry-run
  xlsx_edit.py --file /path/book.xlsx --ops ops.json --allow-protected

Ops:
  {"op":"set_cell","sheet":"数据","cell":"B2","value":"999","expect":"100","value_type":"auto"}
  {"op":"replace_text","sheet":"数据","find":"2025年","replace":"2026年","range":"A1:F20","all":false,"occurrence":1}
  {"op":"set_format","sheet":"数据","cell":"B2","font":"仿宋","size":12,"color":"FF0000",
   "bold":true,"italic":false,"underline":"single","number_format":"#,##0.00",
   "align":"center","wrap":true}

Scope: cell values, in-cell text replacement and cell formatting on a single
worksheet. Charts, images, comments, tables, data validation, conditional
formatting and defined names are preserved by openpyxl; any zip part that
disappears during the round-trip aborts the write. Formula cached values are
dropped (fullCalcOnLoad is set, so Excel/WPS recalculates on open).

Protected worksheets are refused unless --allow-protected is passed, which the
agent may only use after the user confirms.

Exit codes: 0 ok, 2 not-found/ambiguous/value-mismatch/bad-request,
3 unsupported/sheet-protected, 4 verify failed, 5 file/exception.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import time
import unicodedata
import zipfile
from copy import copy
from datetime import date, datetime
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Color
    from openpyxl.utils import range_boundaries
except ImportError as exc:  # pragma: no cover - environment guard
    print(f"ERROR [exception] 缺少 openpyxl：{exc}；请运行 pip install openpyxl", file=sys.stderr)
    raise SystemExit(5)

try:
    from openpyxl.cell.rich_text import CellRichText
except Exception:  # pragma: no cover - older openpyxl
    CellRichText = ()

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from doc_review_log import log_event as _log_event
except Exception:  # pragma: no cover - logger is optional at runtime
    def _log_event(**kwargs):  # type: ignore
        return None


def _compact_ops(ops):
    try:
        from doc_review_log import ops_summary
        return ops_summary(ops)
    except Exception:
        return None


def safe_log(**kwargs):
    """Logging must never fail or undo a successful edit."""
    try:
        return _log_event(**kwargs)
    except Exception:
        return None


class EditError(Exception):
    def __init__(self, code: str, message: str, stage: str = "edit", details: dict | None = None):
        super().__init__(message)
        self.code = code
        self.stage = stage
        self.details = details or {}


CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
INT_RE = re.compile(r"[-+]?\d+")
DEC_RE = re.compile(r"[-+]?\d*\.\d+")
DATE_HINT = re.compile(r"[ymdhs]", re.I)
DATE_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
)
TRUE_WORDS = {"true", "1", "是", "yes", "y"}
FALSE_WORDS = {"false", "0", "否", "no", "n"}


def sha256(path: Path) -> str:
    return "sha256:" + hashlib.sha256(path.read_bytes()).hexdigest()


def zip_parts(path: Path) -> dict:
    with zipfile.ZipFile(path) as archive:
        return {name: hashlib.sha256(archive.read(name)).hexdigest() for name in archive.namelist()}


def norm(text: str) -> str:
    return unicodedata.normalize("NFKC", "" if text is None else str(text)).strip()


def display(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime):
        if value.hour == value.minute == value.second == 0:
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def same_value(actual, expected_text: str) -> bool:
    if actual is None:
        return norm(expected_text) == ""
    text = norm(expected_text)
    if isinstance(actual, bool):
        return text.lower() in (TRUE_WORDS if actual else FALSE_WORDS)
    if isinstance(actual, (int, float)) and not isinstance(actual, bool):
        cleaned = text.replace(",", "")
        try:
            if cleaned.endswith("%"):
                return abs(float(actual) * 100 - float(cleaned[:-1])) < 1e-6
            return abs(float(actual) - float(cleaned)) < 1e-9
        except ValueError:
            return False
    return norm(display(actual)) == text


def is_date_format(fmt: str | None) -> bool:
    if not fmt:
        return False
    cleaned = re.sub(r'"[^"]*"', "", str(fmt))
    cleaned = re.sub(r"\[[^\]]*\]", "", cleaned)
    return bool(DATE_HINT.search(cleaned))


def is_percent_format(fmt: str | None) -> bool:
    return bool(fmt) and "%" in str(fmt)


def parse_date(text: str):
    text = text.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def coerce_value(raw, cell, value_type: str = "auto", warnings: list | None = None):
    kind = (value_type or "auto").lower()
    text = "" if raw is None else str(raw)
    if kind == "string":
        return text
    if kind == "formula":
        stripped = text.strip()
        return stripped if stripped.startswith("=") else "=" + stripped
    if kind == "bool":
        low = text.strip().lower()
        if low in TRUE_WORDS:
            return True
        if low in FALSE_WORDS:
            return False
        raise EditError("bad_request", f"无法识别布尔值：{raw}", "preflight")
    if kind == "number":
        try:
            return float(text.replace(",", ""))
        except ValueError:
            raise EditError("bad_request", f"无法识别数值：{raw}", "preflight")
    if kind == "date":
        parsed = parse_date(text)
        if parsed is None:
            raise EditError("bad_request", f"无法识别日期：{raw}", "preflight")
        return parsed

    stripped = text.strip()
    if stripped.startswith("=") and len(stripped) > 1:
        return stripped
    if stripped.endswith("%"):
        try:
            number = float(stripped[:-1].replace(",", "")) / 100
        except ValueError:
            number = None
        if number is not None:
            if not is_percent_format(cell.number_format) and warnings is not None:
                warnings.append({"code": "percent_written_as_number", "cell": cell.coordinate,
                                 "message": "百分比按数值写入，单元格格式保持原样"})
            return number
    if re.fullmatch(r"[-+]?\d{1,3}(,\d{3})+(\.\d+)?", stripped):
        stripped = stripped.replace(",", "")
    if INT_RE.fullmatch(stripped):
        digits = stripped.lstrip("+-")
        if len(digits) > 1 and digits.startswith("0"):
            return text  # keep 编号/工号 as text
        return int(stripped)
    if DEC_RE.fullmatch(stripped):
        return float(stripped)
    low = stripped.lower()
    if low in {"true", "false"}:
        return low == "true"
    if is_date_format(cell.number_format):
        parsed = parse_date(stripped)
        if parsed is not None:
            return parsed
        if warnings is not None:
            warnings.append({"code": "date_written_as_text", "cell": cell.coordinate,
                             "message": "值为日期文本，单元格日期格式未变更"})
    return text


def normalize_cell(text: str) -> str:
    match = CELL_RE.match(str(text).strip())
    if not match:
        raise EditError("bad_request", f"单元格地址非法：{text}", "preflight")
    return f"{match.group(1).upper()}{int(match.group(2))}"


def parse_range(text: str) -> tuple[int, int, int, int]:
    cleaned = str(text or "").replace("$", "").strip().upper()
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cleaned)
    except Exception:
        raise EditError("bad_request", f"区域地址非法：{text}", "preflight")
    if None in (min_col, min_row, max_col, max_row):
        raise EditError("bad_request", f"区域地址非法：{text}", "preflight")
    return min_col, min_row, max_col, max_row


def get_sheet(workbook, name: str):
    if not name:
        raise EditError("bad_request", "缺少 sheet", "preflight")
    if name in workbook.sheetnames:
        return workbook[name]
    matches = [item for item in workbook.sheetnames if item.lower() == str(name).lower()]
    if len(matches) == 1:
        return workbook[matches[0]]
    if len(matches) > 1:
        raise EditError("ambiguous", f"工作表名 {name} 不唯一", "preflight", {"candidates": matches})
    raise EditError(
        "not_found",
        f"工作表不存在：{name}",
        "preflight",
        {"sheets": workbook.sheetnames[:30]},
    )


def check_sheet_protection(sheet, allow_protected: bool) -> None:
    if getattr(sheet.protection, "sheet", False) and not allow_protected:
        raise EditError(
            "sheet_protected",
            f"工作表「{sheet.title}」受保护，未写入；需向用户确认后加 --allow-protected 重试",
            "preflight",
            {"sheet": sheet.title, "retry": "--allow-protected"},
        )


def get_cell(sheet, ref: str):
    coordinate = normalize_cell(ref)
    cell = sheet[coordinate]
    if isinstance(cell, MergedCell):
        for merged in sheet.merged_cells.ranges:
            if coordinate in merged:
                raise EditError(
                    "unsupported",
                    f"{sheet.title}!{coordinate} 是合并单元格的非左上角，不能单独写入",
                    "preflight",
                    {"sheet": sheet.title, "cell": coordinate, "top_left": merged.start_cell.coordinate,
                     "range": str(merged)},
                )
        raise EditError("unsupported", f"{sheet.title}!{coordinate} 不可写", "preflight")
    if type(cell.value).__name__ == "ArrayFormula":
        raise EditError(
            "unsupported",
            f"{sheet.title}!{coordinate} 是数组公式的一部分，不能单独修改",
            "preflight",
            {"sheet": sheet.title, "cell": coordinate},
        )
    return cell


def text_matches(sheet, needle: str, bounds: tuple[int, int, int, int]):
    min_col, min_row, max_col, max_row = bounds
    matches = []
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            value = cell.value
            if CellRichText and isinstance(value, CellRichText):
                continue
            if not isinstance(value, str) or value.startswith("="):
                continue
            start = value.find(needle)
            while start != -1:
                matches.append({"cell": cell, "start": start, "end": start + len(needle), "text": value})
                start = value.find(needle, start + len(needle))
    return matches


def select_matches(matches: list, op: dict) -> list:
    if not matches:
        raise EditError("not_found", f"未找到引文：{op.get('find')}", "match")
    if op.get("all"):
        return matches
    occurrence = op.get("occurrence")
    if occurrence:
        try:
            return [matches[int(occurrence) - 1]]
        except (IndexError, ValueError):
            raise EditError("not_found", f"occurrence={occurrence} 超出命中数 {len(matches)}", "match")
    cells = {(m["cell"].parent.title, m["cell"].coordinate) for m in matches}
    if len(cells) > 1:
        candidates = [
            {"cell": f"{m['cell'].parent.title}!{m['cell'].coordinate}", "text": m["text"][:80]}
            for m in matches[:8]
        ]
        hint = (
            "；范围内多处命中时加 all:true，或指定 occurrence"
            if op.get("range")
            else "；如全部替换请加 all:true，或用 range 限定范围"
        )
        raise EditError(
            "ambiguous",
            f"引文在 {len(cells)} 个单元格命中{hint}",
            "match",
            {"candidates": candidates},
        )
    return matches


def apply_set_cell(sheet, op: dict, warnings: list) -> dict:
    cell = get_cell(sheet, op.get("cell"))
    if "expect" in op and op.get("expect") is not None:
        if not same_value(cell.value, str(op["expect"])):
            raise EditError(
                "value_mismatch",
                f"{sheet.title}!{cell.coordinate} 现值与批注不符：批注写“{op['expect']}”，实际“{display(cell.value)}”",
                "match",
                {"sheet": sheet.title, "cell": cell.coordinate,
                 "expect": str(op["expect"]), "actual": display(cell.value)},
            )
    before = cell.value
    after = coerce_value(op.get("value"), cell, op.get("value_type"), warnings)
    changed = type(before) is not type(after) or not same_value(before, display(after))
    cell.value = after
    return {
        "op": "set_cell",
        "sheet": sheet.title,
        "targets": [cell.coordinate],
        "cells": [cell.coordinate],
        "before": display(before),
        "after": display(after),
        "changed": changed,
        "expected": display(after),
    }


def apply_replace_text(sheet, op: dict, warnings: list) -> dict:
    needle = str(op.get("find") or "")
    if not needle:
        raise EditError("bad_request", "find 不能为空", "preflight")
    bounds = parse_range(op["range"]) if op.get("range") else parse_range(sheet.calculate_dimension())
    matches = text_matches(sheet, needle, bounds)
    selected = select_matches(matches, op)
    replacement = str(op.get("replace", ""))
    by_cell: dict[str, list] = {}
    for match in selected:
        by_cell.setdefault(match["cell"].coordinate, []).append(match)
    changed_cells = []
    before_text = {}
    after_text = {}
    for coordinate, items in by_cell.items():
        cell = sheet[coordinate]
        original = str(cell.value)
        before_text[coordinate] = original
        rebuilt = original
        for item in sorted(items, key=lambda m: m["start"], reverse=True):
            rebuilt = rebuilt[: item["start"]] + replacement + rebuilt[item["end"]:]
        after_text[coordinate] = rebuilt
        if rebuilt != original:
            changed_cells.append(coordinate)
        cell.value = rebuilt
    return {
        "op": "replace_text",
        "sheet": sheet.title,
        "targets": list(by_cell.keys()),
        "cells": list(by_cell.keys()),
        "hits": len(selected),
        "find": needle,
        "replace": replacement,
        "changed": bool(changed_cells),
        "before_text": before_text,
        "after_text": after_text,
    }


FORMAT_FONT_KEYS = ("font", "size", "color", "bold", "italic", "underline")


def apply_set_format(sheet, op: dict, warnings: list) -> dict:
    if op.get("cell"):
        bounds = parse_range(str(op["cell"]))
    elif op.get("range"):
        bounds = parse_range(str(op["range"]))
    else:
        raise EditError("bad_request", "set_format 需要 cell 或 range", "preflight")
    min_col, min_row, max_col, max_row = bounds
    applied = []
    changed = False
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            touched = False
            if any(key in op for key in FORMAT_FONT_KEYS):
                font = copy(cell.font)
                if op.get("font"):
                    font.name = str(op["font"])
                if op.get("size"):
                    font.size = float(op["size"])
                if op.get("color"):
                    font.color = Color(rgb=str(op["color"]).lstrip("#").upper())
                if "bold" in op:
                    font.bold = bool(op["bold"])
                if "italic" in op:
                    font.italic = bool(op["italic"])
                if op.get("underline"):
                    font.underline = op["underline"]
                cell.font = font
                touched = True
            if op.get("number_format"):
                cell.number_format = str(op["number_format"])
                touched = True
            if any(key in op for key in ("align", "wrap")):
                alignment = copy(cell.alignment)
                if op.get("align"):
                    alignment.horizontal = str(op["align"])
                    alignment.vertical = "center"
                if "wrap" in op:
                    alignment.wrap_text = bool(op["wrap"])
                cell.alignment = alignment
                touched = True
            if touched:
                applied.append(cell.coordinate)
                changed = True
    if not applied:
        raise EditError("bad_request", "set_format 未指定任何格式项", "preflight")
    return {
        "op": "set_format",
        "sheet": sheet.title,
        "targets": applied,
        "cells": applied,
        "format": {key: op[key] for key in (*FORMAT_FONT_KEYS, "number_format", "align", "wrap") if key in op},
        "changed": changed,
    }


APPLIERS = {
    "set_cell": apply_set_cell,
    "replace_text": apply_replace_text,
    "set_format": apply_set_format,
}

ALLOWED_FIELDS = {
    "set_cell": {"op", "type", "sheet", "cell", "value", "expect", "value_type"},
    "replace_text": {"op", "type", "sheet", "find", "replace", "range", "all", "occurrence"},
    "set_format": {"op", "type", "sheet", "cell", "range", "font", "size", "color", "bold",
                   "italic", "underline", "number_format", "align", "wrap"},
}


def load_ops(raw: str) -> list:
    text = str(raw).strip()
    if text.startswith(("[", "{")):
        data_text = text
    else:
        path = Path(raw).expanduser()
        try:
            is_file = path.is_file()
        except OSError:
            is_file = False
        if not is_file:
            raise EditError("bad_request", "ops 既不是 JSON，也不是可读的 ops 文件路径", "preflight")
        data_text = path.read_text("utf-8")
    try:
        data = json.loads(data_text)
    except json.JSONDecodeError as exc:
        raise EditError("bad_request", f"ops JSON 解析失败：{exc}", "preflight")
    if isinstance(data, dict):
        data = data.get("operations") or data.get("ops") or [data]
    if not isinstance(data, list) or not data:
        raise EditError("bad_request", "ops 必须是非空数组", "preflight")
    for item in data:
        if not isinstance(item, dict):
            raise EditError("bad_request", "ops 每项必须是对象", "preflight")
        kind = item.get("op") or item.get("type")
        if kind not in APPLIERS:
            raise EditError("unsupported", f"不支持的 op：{kind}", "preflight",
                            {"supported": sorted(APPLIERS)})
        unknown = sorted(set(item) - ALLOWED_FIELDS[kind])
        if unknown:
            raise EditError(
                "bad_request",
                f"{kind} 不支持字段：{', '.join(unknown)}",
                "preflight",
                {"supported_fields": sorted(ALLOWED_FIELDS[kind])},
            )
        item["op"] = kind

    for item in data:
        kind = item["op"]
        field = "cell" if kind in {"set_cell", "set_format"} else "range"
        raw = item.get(field)
        if raw and "!" in str(raw):
            prefix, rest = str(raw).split("!", 1)
            prefix = prefix.strip().strip("'").strip('"')
            if item.get("sheet") and norm(item["sheet"]) != norm(prefix):
                raise EditError(
                    "bad_request",
                    f"sheet={item['sheet']} 与地址前缀 {prefix} 冲突",
                    "preflight",
                )
            item["sheet"] = item.get("sheet") or prefix
            item[field] = rest.strip()
        if not item.get("sheet"):
            raise EditError(
                "bad_request",
                f"{kind} 缺少 sheet；也可把地址写成 sheet!cell（如 台账!C2）",
                "preflight",
            )
    return data


def formula_map(workbook) -> dict:
    formulas = {}
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    formulas[(sheet.title, cell.coordinate)] = value
    return formulas


def merge_map(workbook) -> dict:
    return {sheet.title: sorted(str(item) for item in sheet.merged_cells.ranges) for sheet in workbook.worksheets}


def defined_map(workbook) -> dict:
    return {name: getattr(item, "attr_text", None) for name, item in workbook.defined_names.items()}


def verify(temp_path: Path, applied: list, before: dict) -> None:
    workbook = load_workbook(str(temp_path), data_only=False)
    if workbook.sheetnames != before["sheets"]:
        raise EditError("verify_failed", "工作表列表发生变化", "verify")
    if merge_map(workbook) != before["merges"]:
        raise EditError("verify_failed", "合并单元格发生变化", "verify")
    if defined_map(workbook) != before["names"]:
        raise EditError("verify_failed", "定义名称发生变化", "verify")

    touched = {(item["sheet"], coordinate) for item in applied for coordinate in item["targets"]}
    after_formulas = formula_map(workbook)
    for key, formula in before["formulas"].items():
        if key in touched:
            continue
        if after_formulas.get(key) != formula:
            raise EditError("verify_failed", f"未操作单元格的公式被改动：{key[0]}!{key[1]}", "verify")
    for key in after_formulas:
        if key in touched or key in before["formulas"]:
            continue
        raise EditError("verify_failed", f"出现意外公式：{key[0]}!{key[1]}", "verify")

    for item in applied:
        sheet = workbook[item["sheet"]]
        for coordinate in item["targets"]:
            cell = sheet[coordinate]
            if item["op"] == "set_cell":
                if not same_value(cell.value, item["expected"]):
                    raise EditError("verify_failed",
                                    f"{item['sheet']}!{coordinate} 回读值不符", "verify")
            elif item["op"] == "replace_text":
                expected = item["after_text"][coordinate]
                if norm(display(cell.value)) != norm(expected):
                    raise EditError("verify_failed",
                                    f"{item['sheet']}!{coordinate} 回读文本不符", "verify")
            elif item["op"] == "set_format":
                fmt = item["format"]
                if fmt.get("font") and cell.font.name != fmt["font"]:
                    raise EditError("verify_failed", f"{item['sheet']}!{coordinate} 字体未生效", "verify")
                if fmt.get("bold") is not None and bool(cell.font.bold) != bool(fmt["bold"]):
                    raise EditError("verify_failed", f"{item['sheet']}!{coordinate} 加粗未生效", "verify")
                if fmt.get("number_format") and cell.number_format != fmt["number_format"]:
                    raise EditError("verify_failed", f"{item['sheet']}!{coordinate} 数字格式未生效", "verify")
                if fmt.get("color"):
                    got = str(cell.font.color.rgb or "").upper()[-6:] if cell.font.color else ""
                    if got != str(fmt["color"]).lstrip("#").upper()[-6:]:
                        raise EditError("verify_failed", f"{item['sheet']}!{coordinate} 颜色未生效", "verify")


def backup_file(path: Path, digest: str) -> Path:
    for candidate in sorted(path.parent.glob(f"{path.stem}.bak-*{path.suffix}")):
        try:
            if sha256(candidate) == digest:
                return candidate
        except OSError:
            continue
    backup = path.with_name(f"{path.stem}.bak-{datetime.now().strftime('%Y%m%d-%H%M%S')}{path.suffix}")
    shutil.copy2(path, backup)
    return backup


def describe(item: dict, warnings: list) -> list[str]:
    lines = []
    if item["op"] == "set_cell":
        lines.append(f"set_cell {item['sheet']}!{item['targets'][0]}: “{item['before']}” -> “{item['after']}”")
    elif item["op"] == "replace_text":
        cells = ", ".join(f"{item['sheet']}!{coord}" for coord in item["targets"][:5])
        suffix = f" 等 {len(item['targets'])} 格" if len(item["targets"]) > 5 else ""
        lines.append(f"replace_text {cells}{suffix}: “{item['find']}” -> “{item['replace']}”（{item['hits']} 处）")
    elif item["op"] == "set_format":
        fmt = item["format"]
        pieces = [f"{key}={value}" for key, value in fmt.items()]
        scope = f"{len(item['targets'])} 个单元格" if len(item["targets"]) > 1 else item["targets"][0]
        lines.append(f"set_format {item['sheet']}!{scope}: {'/'.join(pieces)}")
    return lines


def main() -> None:
    parser = argparse.ArgumentParser(description="xlsx edit helper")
    parser.add_argument("--file", required=True)
    parser.add_argument("--ops", required=True, help="ops JSON file or inline JSON")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--allow-protected", action="store_true", help="仅用户确认后使用")
    parser.add_argument("--contract")
    parser.add_argument("--json", action="store_true", dest="as_json")
    args = parser.parse_args()

    started = time.monotonic()
    path = Path(args.file).expanduser().resolve()
    contract = args.contract or "v1.6"
    warnings: list = []
    applied: list = []
    try:
        if not path.is_file():
            raise EditError("file_error", f"文件不存在：{path}", "preflight")
        suffix = path.suffix.lower()
        if suffix not in {".xlsx", ".xlsm"}:
            raise EditError(
                "unsupported",
                f"不支持的扩展名 {suffix}；请另存为 .xlsx",
                "preflight",
                {"hint": "xls/csv 需先另存为 xlsx"},
            )
        ops = load_ops(args.ops)
        workbook = load_workbook(str(path), keep_vba=(suffix == ".xlsm"), data_only=False)
        before_digest = sha256(path)
        before_parts = zip_parts(path)
        before = {
            "sheets": list(workbook.sheetnames),
            "formulas": formula_map(workbook),
            "merges": merge_map(workbook),
            "names": defined_map(workbook),
        }

        current_sheet = None
        for op in ops:
            sheet_name = op.get("sheet")
            if not sheet_name:
                raise EditError("bad_request", f"{op['op']} 缺少 sheet", "preflight")
            if current_sheet is None or current_sheet.title != sheet_name:
                current_sheet = get_sheet(workbook, sheet_name)
                check_sheet_protection(current_sheet, args.allow_protected)
                if current_sheet.sheet_state != "visible":
                    warnings.append({"code": "hidden_sheet", "sheet": current_sheet.title,
                                     "message": "工作表处于隐藏状态"})
            applied.append(APPLIERS[op["op"]](current_sheet, op, warnings))

        changed = any(item["changed"] for item in applied)
        result = {
            "ok": True,
            "file": str(path),
            "ops": applied,
            "changed": changed,
            "warnings": warnings,
        }

        if args.dry_run:
            print(json.dumps({**result, "dry_run": True}, ensure_ascii=False, indent=2))
            return

        if not changed:
            print(f"OK {path.name}（已是目标状态，未写入）")
            safe_log(event="helper_ok", contract=contract, file=str(path), file_hash=before_digest,
                     duration_ms=int((time.monotonic() - started) * 1000), ops=_compact_ops(ops),
                     note="no-op")
            return

        workbook.calculation.fullCalcOnLoad = True
        backup = backup_file(path, before_digest)
        temp = path.with_name(path.stem + ".dsh-tmp" + path.suffix)
        try:
            workbook.save(str(temp))
            verify(temp, applied, before)
            after_parts = zip_parts(temp)
            missing = sorted(set(before_parts) - set(after_parts) - {"xl/calcChain.xml"})
            if missing:
                raise EditError(
                    "unsupported",
                    "打开-保存会丢失文件部件（透视表/切片器/VBA 等），已放弃写入",
                    "write",
                    {"missing_parts": missing[:20]},
                )
            added = sorted(set(after_parts) - set(before_parts))
            changed_parts = [name for name in sorted(set(before_parts) & set(after_parts))
                             if before_parts[name] != after_parts[name]]
            expected_changes = {
                "xl/sharedStrings.xml",
                "xl/styles.xml",
                "xl/workbook.xml",
                "docProps/app.xml",
                "docProps/core.xml",
            }
            unexpected = [name for name in changed_parts
                          if not name.startswith("xl/worksheets/") and name not in expected_changes]
            os.chmod(temp, path.stat().st_mode)
            try:
                os.replace(temp, path)
            except PermissionError as exc:
                raise EditError(
                    "file_error",
                    f"无法替换原文件，可能被 Excel/WPS 占用：{exc}",
                    "write",
                    {"hint": "关闭占用该文件的程序后重试"},
                )
        finally:
            if temp.exists():
                try:
                    temp.unlink()
                except OSError:
                    pass

        after_digest = sha256(path)
        result.update({
            "backup": str(backup),
            "hash_before": before_digest,
            "hash_after": after_digest,
            "verification": {
                "ok": True,
                "parts_before": len(before_parts),
                "parts_after": len(after_parts),
                "changed_parts": changed_parts,
                "unexpected_parts": unexpected,
                "added_parts": added,
            },
        })
        safe_log(event="helper_ok", contract=contract, file=str(path), file_hash=before_digest,
                 duration_ms=int((time.monotonic() - started) * 1000), ops=_compact_ops(ops))

        if args.as_json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            print(f"OK {path.name}")
            print(f"备份：{backup.name}")
            print(f"hash：{before_digest[7:19]}… -> {after_digest[7:19]}…")
            index = 0
            for item in applied:
                for line in describe(item, warnings):
                    index += 1
                    print(f"{index}. {line}")
            print(f"校验：通过；变更部件：{', '.join(changed_parts) or '无'}；"
                  f"部件守恒 {len(after_parts)}/{len(before_parts)}")
            if unexpected:
                print(f"警告：额外变更部件：{', '.join(unexpected[:5])}")
            for item in warnings:
                print(f"警告：{item.get('sheet', '')} {item['message']}".strip())
    except EditError as exc:
        safe_log(
            event="verify_fail" if exc.code == "verify_failed" else "helper_fail",
            contract=contract,
            file=str(path),
            stage=exc.stage,
            code=exc.code,
            reason=str(exc),
            duration_ms=int((time.monotonic() - started) * 1000),
            ops=_compact_ops(ops) if "ops" in locals() else None,
            note=json.dumps(exc.details, ensure_ascii=False) if exc.details else None,
        )
        print(f"ERROR [{exc.code}] {exc}", file=sys.stderr)
        if exc.details:
            print(json.dumps(exc.details, ensure_ascii=False), file=sys.stderr)
        code = {"verify_failed": 4, "unsupported": 3, "sheet_protected": 3,
                "not_found": 2, "ambiguous": 2, "value_mismatch": 2, "bad_request": 2}.get(exc.code, 5)
        raise SystemExit(code)
    except Exception as exc:  # noqa: BLE001
        safe_log(
            event="helper_fail",
            contract=contract,
            file=str(path),
            stage="exception",
            code="exception",
            reason=f"{type(exc).__name__}: {exc}",
            duration_ms=int((time.monotonic() - started) * 1000),
            ops=_compact_ops(ops) if "ops" in locals() else None,
        )
        print(f"ERROR [exception] {type(exc).__name__}: {exc}", file=sys.stderr)
        raise SystemExit(5)


if __name__ == "__main__":
    main()
