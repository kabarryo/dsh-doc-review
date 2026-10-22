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

SAMPLE = """# 实践教学基地建设要点

## 一、建设目标

以**产教融合**为主线，建设一批示范性实践教学基地。

- 基地数量：不少于 10 个
- 年度经费：**200 万元**

## 二、管理机制

| 事项 | 责任单位 |
| --- | --- |
| 立项 | 教务处 |
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
            [{"op": "replace_text", "find": "一、建设目标", "replace": "一、建设目标与任务"}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("第 3 行", result.stdout)
        self.assertIn("一、建设目标与任务", self.file.read_text("utf-8"))
        self.assertEqual(len(list(self.tmp.glob("doc.md.bak-*"))), 1)

    def test_replace_across_bold(self) -> None:
        result = self.run_helper(
            [{"op": "replace_text", "find": "产教融合为主线", "replace": "科教融汇为主线"}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        text = self.file.read_text("utf-8")
        self.assertIn("科教融汇为主线", text)
        self.assertIn("标记归一化", result.stdout)

    def test_ambiguous_then_occurrence(self) -> None:
        found = self.run_helper([{"op": "replace_text", "find": "基地", "replace": "中心"}])
        self.assertEqual(found.returncode, 2, found.stderr)
        self.assertIn("ambiguous", found.stderr)
        self.assertIn("candidates", found.stderr)

        second = self.run_helper(
            [{"op": "replace_text", "find": "基地", "replace": "中心", "occurrence": 1}]
        )
        self.assertEqual(second.returncode, 0, second.stderr)
        self.assertIn("中心", self.file.read_text("utf-8"))

    def test_line_scoped(self) -> None:
        result = self.run_helper(
            [{"op": "replace_text", "find": "教务处", "replace": "教务处（修订）", "line": 14}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        text = self.file.read_text("utf-8")
        self.assertIn("| 立项 | 教务处（修订） |", text)

    def test_delete_text(self) -> None:
        result = self.run_helper([{"op": "delete_text", "find": "- 年度经费：**200 万元**"}])
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertNotIn("200 万元", self.file.read_text("utf-8"))

    def test_insert_after(self) -> None:
        result = self.run_helper(
            [{"op": "insert_after", "find": "- 基地数量：不少于 10 个", "content": "- 覆盖专业：不少于 5 个"}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        lines = self.file.read_text("utf-8").splitlines()
        index = lines.index("- 基地数量：不少于 10 个")
        self.assertEqual(lines[index + 1], "- 覆盖专业：不少于 5 个")

    def test_dry_run(self) -> None:
        digest = self.digest()
        result = self.run_helper(
            [{"op": "replace_text", "find": "一、建设目标", "replace": "x"}], "--dry-run"
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("dry_run", result.stdout)
        self.assertEqual(self.digest(), digest)

    def test_noop(self) -> None:
        digest = self.digest()
        result = self.run_helper(
            [{"op": "replace_text", "find": "一、建设目标", "replace": "一、建设目标"}]
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("已是目标状态", result.stdout)
        self.assertEqual(self.digest(), digest)
        self.assertEqual(list(self.tmp.glob("doc.md.bak-*")), [])

    def test_unknown_field(self) -> None:
        result = self.run_helper(
            [{"op": "replace_text", "find": "立项", "replace": "x", "style": "bold"}]
        )
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertIn("bad_request", result.stderr)
        self.assertIn("style", result.stderr)

    def test_crlf_preserved(self) -> None:
        self.file.write_bytes(SAMPLE.replace("\n", "\r\n").encode("utf-8"))
        result = self.run_helper([{"op": "replace_text", "find": "一、建设目标", "replace": "一、目标"}])
        self.assertEqual(result.returncode, 0, result.stderr)
        raw = self.file.read_bytes()
        self.assertIn(b"\r\n", raw)
        self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"))

    def test_logs_written(self) -> None:
        self.run_helper([{"op": "replace_text", "find": "一、建设目标", "replace": "一、目标"}])
        events = self.logs()
        self.assertTrue(events)
        latest = events[-1]
        self.assertEqual(latest.get("event"), "helper_ok")
        self.assertEqual(latest.get("contract"), "v1.7")
        self.assertEqual(latest["ops"][0].get("op"), "replace_text")


if __name__ == "__main__":
    unittest.main(verbosity=2)
