#!/usr/bin/env python3
"""Tests for scripts/md_edit.py.

Run from the repository root:
  python3 -m unittest discover -s tests -v
"""
from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HELPER = ROOT / "scripts" / "md_edit.py"

SAMPLE = """# 示例文档

## 一、概述

以**示例文本**为主线，构建一份可复用的样例材料。

- 条目数量：不少于 10 个
- 年度预算：**200 元**

## 二、流程

| 事项 | 负责人 |
| --- | --- |
| 起草 | 张三 |
| 复核 | 李四 |
"""


class MarkdownCase(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="md-edit-"))
        self.home = self.tmp / "home"
        self.home.mkdir()
        self.file = self.tmp / "doc.md"
        self.file.write_text(SAMPLE, "utf-8")

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_helper(self, ops, *args: str, file: Path | None = None):
        env = dict(os.environ, DOC_REVIEW_HOME=str(self.home))
        command = [
            sys.executable,
            str(HELPER),
            "--file",
            str(file or self.file),
            "--ops",
            json.dumps(ops, ensure_ascii=False),
            *args,
        ]
        return subprocess.run(command, capture_output=True, text=True, env=env)

    def digest(self) -> str:
        return hashlib.sha256(self.file.read_bytes()).hexdigest()

    def logs(self) -> list:
        directory = self.home / "logs"
        if not directory.exists():
            return []
        return [
            json.loads(line)
            for file in sorted(directory.glob("events-*.jsonl"))
            for line in file.read_text("utf-8").splitlines()
        ]

    def test_replace_text(self) -> None:
        result = self.run_helper(
            [{"op": "replace_text", "find": "一、概述", "replace": "一、概述与范围"}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("第 3 行", result.stdout)
        self.assertIn("一、概述与范围", self.file.read_text("utf-8"))
        self.assertEqual(len(list(self.tmp.glob("doc.md.bak-*"))), 1)

    def test_replace_across_bold(self) -> None:
        result = self.run_helper(
            [{"op": "replace_text", "find": "示例文本为主线", "replace": "样例文本为主线"}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        text = self.file.read_text("utf-8")
        self.assertIn("样例文本为主线", text)
        self.assertIn("标记归一化", result.stdout)

    def test_ambiguous_then_occurrence(self) -> None:
        found = self.run_helper([{"op": "replace_text", "find": "示例", "replace": "样例"}])
        self.assertEqual(found.returncode, 2, found.stderr)
        self.assertIn("ambiguous", found.stderr)
        self.assertIn("candidates", found.stderr)

        second = self.run_helper(
            [{"op": "replace_text", "find": "示例", "replace": "样例", "occurrence": 1}]
        )
        self.assertEqual(second.returncode, 0, second.stderr)
        self.assertIn("样例", self.file.read_text("utf-8"))

    def test_line_scoped(self) -> None:
        result = self.run_helper(
            [{"op": "replace_text", "find": "张三", "replace": "张三（修订）", "line": 14}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        text = self.file.read_text("utf-8")
        self.assertIn("| 起草 | 张三（修订） |", text)

    def test_delete_text(self) -> None:
        result = self.run_helper([{"op": "delete_text", "find": "- 年度预算：**200 元**"}])
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertNotIn("200 元", self.file.read_text("utf-8"))

    def test_insert_after(self) -> None:
        result = self.run_helper(
            [{"op": "insert_after", "find": "- 条目数量：不少于 10 个", "content": "- 覆盖范围：不限"}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        lines = self.file.read_text("utf-8").splitlines()
        index = lines.index("- 条目数量：不少于 10 个")
        self.assertEqual(lines[index + 1], "- 覆盖范围：不限")

    def test_dry_run(self) -> None:
        digest = self.digest()
        result = self.run_helper(
            [{"op": "replace_text", "find": "一、概述", "replace": "x"}], "--dry-run"
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("dry_run", result.stdout)
        self.assertEqual(self.digest(), digest)

    def test_noop(self) -> None:
        digest = self.digest()
        result = self.run_helper(
            [{"op": "replace_text", "find": "一、概述", "replace": "一、概述"}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("已是目标状态", result.stdout)
        self.assertEqual(self.digest(), digest)
        self.assertEqual(list(self.tmp.glob("doc.md.bak-*")), [])

    def test_unknown_field(self) -> None:
        result = self.run_helper(
            [{"op": "replace_text", "find": "起草", "replace": "x", "style": "bold"}]
        )
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertIn("bad_request", result.stderr)
        self.assertIn("style", result.stderr)

    def test_crlf_preserved(self) -> None:
        self.file.write_bytes(SAMPLE.replace("\n", "\r\n").encode("utf-8"))
        result = self.run_helper([{"op": "replace_text", "find": "一、概述", "replace": "一、概述与范围"}])
        self.assertEqual(result.returncode, 0, result.stderr)
        raw = self.file.read_bytes()
        self.assertIn(b"\r\n", raw)
        self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"))

    def test_logs_written(self) -> None:
        self.run_helper([{"op": "replace_text", "find": "一、概述", "replace": "一、概述与范围"}])
        events = self.logs()
        self.assertTrue(events)
        latest = events[-1]
        self.assertEqual(latest.get("event"), "helper_ok")
        self.assertEqual(latest.get("contract"), "v1.7")
        self.assertEqual(latest["ops"][0].get("op"), "replace_text")


if __name__ == "__main__":
    unittest.main(verbosity=2)
