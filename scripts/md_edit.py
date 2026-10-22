#!/usr/bin/env python3
"""Declarative Markdown edit helper for the dsh-doc-review plugin.

The sidebar renders markdown, so the quote captured from the preview is the
rendered text: inline markers such as **, `, ~~ and heading/list prefixes are
missing. This helper matches such quotes against the source file in two passes:
whitespace-normalised first, then with inline markers dropped, so a quote that
spans a bold segment still lands. Every edit is applied in memory, written to a
temp file, read back and compared before the original is replaced.

Usage:
  md_edit.py --file /path/doc.md --ops '[{"op":"replace_text","find":"一、建设目标","replace":"一、建设目标与任务"}]'
  md_edit.py --file /path/doc.md --ops ops.json --dry-run

Ops:
  {"op":"replace_text","find":"...","replace":"...","all":false,"occurrence":1,"line":12}
  {"op":"delete_text","find":"...","all":false,"occurrence":1}
  {"op":"insert_after","find":"...","content":"- 新增一条","occurrence":1}

Scope: plain text edits on a single markdown file. Front matter, code fences and
tables are treated as ordinary text; nothing is reformatted. Exit codes:
0 ok, 2 not-found/ambiguous/bad-request, 3 unsupported, 4 verify failed,
5 file/exception.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

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


MARKERS = "*_`~"
ALLOWED_FIELDS = {
    "replace_text": {"op", "type", "find", "replace", "all", "occurrence", "line"},
    "delete_text": {"op", "type", "find", "all", "occurrence", "line"},
    "insert_after": {"op", "type", "find", "content", "all", "occurrence", "line"},
}


def sha256(path: Path) -> str:
    return "sha256:" + hashlib.sha256(path.read_bytes()).hexdigest()


def norm(text: str) -> str:
    return unicodedata.normalize("NFKC", str(text)).strip()


def read_source(path: Path) -> tuple:
    with path.open("r", encoding="utf-8", newline="") as handle:
        text = handle.read()
    bom = text.startswith("\ufeff")
    return text, bom


def _fold(char: str) -> str:
    """NFKC-fold one character; fall back when folding would change the length."""
    folded = unicodedata.normalize("NFKC", char)
    return folded if len(folded) == 1 else char


def compact(text: str, drop_markers: bool) -> tuple:
    """Whitespace-free projection of text with an index map back to the source."""
    chars = []
    mapping = []
    for index, char in enumerate(text):
        folded = _fold(char)
        if folded.isspace():
            continue
        if drop_markers and folded in MARKERS:
            continue
        chars.append(folded)
        mapping.append(index)
    return "".join(chars), mapping


def line_of(text: str, offset: int) -> int:
    return text.count("\n", 0, offset) + 1


def line_text(text: str, offset: int) -> str:
    start = text.rfind("\n", 0, offset) + 1
    end = text.find("\n", offset)
    if end == -1:
        end = len(text)
    return text[start:end].strip()


def find_quotes(source: str, quote: str) -> tuple:
    """Return (spans, mode). mode is 'exact' or 'markup' for diagnostics."""
    target = str(quote).strip()
    if not target:
        raise EditError("bad_request", "find 不能为空", "preflight")
    for drop_markers in (False, True):
        compact_source, mapping = compact(source, drop_markers)
        compact_quote, _ = compact(target, drop_markers)
        if not compact_quote:
            continue
        spans = []
        start = compact_source.find(compact_quote)
        while start != -1:
            raw_start = mapping[start]
            raw_end = mapping[start + len(compact_quote) - 1] + 1
            spans.append((raw_start, raw_end))
            start = compact_source.find(compact_quote, start + len(compact_quote))
        if spans:
            return spans, ("markup" if drop_markers else "exact")
    return [], "none"


def select_spans(source: str, spans: list, op: dict) -> list:
    if not spans:
        raise EditError("not_found", f"未找到引文：{op.get('find')}", "match")
    line = op.get("line")
    if line:
        try:
            wanted = int(line)
        except (TypeError, ValueError):
            raise EditError("bad_request", f"line 必须是整数：{line}", "preflight")
        spans = [span for span in spans if line_of(source, span[0]) == wanted]
        if not spans:
            raise EditError("not_found", f"第 {wanted} 行没有匹配：{op.get('find')}", "match")
    if op.get("all"):
        return spans
    occurrence = op.get("occurrence")
    if occurrence:
        try:
            return [spans[int(occurrence) - 1]]
        except (IndexError, ValueError):
            raise EditError("not_found", f"occurrence={occurrence} 超出命中数 {len(spans)}", "match")
    if len(spans) > 1:
        candidates = [
            {"line": line_of(source, start), "text": line_text(source, start)[:80]}
            for start, _ in spans[:8]
        ]
        raise EditError(
            "ambiguous",
            f"引文命中 {len(spans)} 处；请指定 all、occurrence 或 line",
            "match",
            {"candidates": candidates},
        )
    return spans


def apply_replace(source: str, spans: list, op: dict) -> tuple:
    replacement = str(op.get("replace", ""))
    ordered = sorted(spans, key=lambda span: span[0], reverse=True)
    changes = []
    for start, end in ordered:
        changes.append({
            "line": line_of(source, start),
            "before": source[start:end],
            "after": replacement,
        })
        source = source[:start] + replacement + source[end:]
    return source, list(reversed(changes))


def apply_delete(source: str, spans: list, op: dict) -> tuple:
    ordered = sorted(spans, key=lambda span: span[0], reverse=True)
    changes = []
    for start, end in ordered:
        changes.append({"line": line_of(source, start), "before": source[start:end], "after": ""})
        source = source[:start] + source[end:]
    return source, list(reversed(changes))


def apply_insert(source: str, spans: list, op: dict) -> tuple:
    content = str(op.get("content") or "")
    if not content.strip():
        raise EditError("bad_request", "insert_after 需要 content", "preflight")
    block = content.rstrip("\n")
    ordered = sorted(spans, key=lambda span: span[1], reverse=True)
    changes = []
    for _, end in ordered:
        newline = source.find("\n", end)
        if newline == -1:
            insert_at = len(source)
            prefix = "" if source.endswith("\n") or not source else "\n"
            suffix = "" if source.endswith("\n") else "\n"
            source = source + prefix + block + suffix
        else:
            insert_at = newline + 1
            source = source[:insert_at] + block + "\n" + source[insert_at:]
        changes.append({"line": line_of(source, insert_at), "inserted": block})
    return source, list(reversed(changes))


APPLIERS = {
    "replace_text": apply_replace,
    "delete_text": apply_delete,
    "insert_after": apply_insert,
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
    return data


def backup_file(path: Path, digest: str) -> Path:
    for candidate in sorted(path.parent.glob(f"{path.name}.bak-*")):
        try:
            if sha256(candidate) == digest:
                return candidate
        except OSError:
            continue
    backup = path.with_name(f"{path.name}.bak-{datetime.now().strftime('%Y%m%d-%H%M%S')}")
    shutil.copy2(path, backup)
    return backup


def main() -> None:
    parser = argparse.ArgumentParser(description="markdown edit helper")
    parser.add_argument("--file", required=True)
    parser.add_argument("--ops", required=True, help="ops JSON file or inline JSON")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--contract")
    parser.add_argument("--json", action="store_true", dest="as_json")
    args = parser.parse_args()

    started = time.monotonic()
    path = Path(args.file).expanduser().resolve()
    contract = args.contract or "v1.7"
    ops: list = []
    try:
        if not path.is_file():
            raise EditError("file_error", f"文件不存在：{path}", "preflight")
        if path.suffix.lower() not in {".md", ".markdown", ".mdx", ".txt"}:
            raise EditError("unsupported", f"不支持的扩展名 {path.suffix}", "preflight")
        ops = load_ops(args.ops)
        source, bom = read_source(path)
        before_digest = sha256(path)

        original = source
        applied = []
        mode = "exact"
        for op in ops:
            spans, found_mode = find_quotes(source, str(op.get("find") or ""))
            if found_mode == "markup":
                mode = "markup"
            selected = select_spans(source, spans, op)
            source, changes = APPLIERS[op["op"]](source, selected, op)
            applied.append({"op": op["op"], "mode": found_mode, "changes": changes})

        changed = source != original
        result = {
            "ok": True,
            "file": str(path),
            "changed": changed,
            "match_mode": mode,
            "ops": applied,
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

        backup = backup_file(path, before_digest)
        temp = path.with_name(path.stem + ".dsh-tmp" + path.suffix)
        try:
            with temp.open("w", encoding="utf-8", newline="") as handle:
                handle.write(("\ufeff" if bom else "") + source)
            written, _ = read_source(temp)
            if written != ("\ufeff" + source if bom else source):
                raise EditError("verify_failed", "回读内容与预期不一致", "verify")
            os.chmod(temp, path.stat().st_mode)
            try:
                os.replace(temp, path)
            except PermissionError as exc:
                raise EditError("file_error", f"无法替换原文件：{exc}", "write",
                                {"hint": "关闭占用该文件的程序后重试"})
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
            "lines_after": source.count("\n") + 1,
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
                for change in item["changes"]:
                    index += 1
                    if item["op"] == "insert_after":
                        print(f"{index}. insert_after 第 {change['line']} 行：{change['inserted']}")
                    else:
                        before = change["before"].replace("\n", " ")[:40]
                        after = change["after"].replace("\n", " ")[:40]
                        print(f"{index}. {item['op']} 第 {change['line']} 行：“{before}” -> “{after}”")
            print(f"校验：通过；匹配方式：{'标记归一化' if mode == 'markup' else '空白归一化'}")
            if mode == "markup":
                print("提示：引文跨行内标记，被替换片段的标记已随之移除")
    except EditError as exc:
        safe_log(
            event="verify_fail" if exc.code == "verify_failed" else "helper_fail",
            contract=contract,
            file=str(path),
            stage=exc.stage,
            code=exc.code,
            reason=str(exc),
            duration_ms=int((time.monotonic() - started) * 1000),
            ops=_compact_ops(ops) if ops else None,
            note=json.dumps(exc.details, ensure_ascii=False) if exc.details else None,
        )
        print(f"ERROR [{exc.code}] {exc}", file=sys.stderr)
        if exc.details:
            print(json.dumps(exc.details, ensure_ascii=False), file=sys.stderr)
        code = {"verify_failed": 4, "unsupported": 3,
                "not_found": 2, "ambiguous": 2, "bad_request": 2}.get(exc.code, 5)
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
            ops=_compact_ops(ops) if ops else None,
        )
        print(f"ERROR [exception] {type(exc).__name__}: {exc}", file=sys.stderr)
        raise SystemExit(5)


if __name__ == "__main__":
    main()
