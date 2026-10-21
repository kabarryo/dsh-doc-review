#!/usr/bin/env python3
"""Structured failure/fallback log for the dsh-doc-review plugin.

The plugin ships this logger so that helper failures, fallbacks to one-off
scripts, and manual interventions leave one machine-readable trace for later
optimization. It records metadata and short excerpts only; never full documents.

Layout:
  ~/.dsh/doc-review/logs/events-YYYY-MM.jsonl

Commands:
  doc_review_log.py log --event fallback_script --reason "..." --file ...
  doc_review_log.py summary
  doc_review_log.py tail --n 20

Environment:
  DOC_REVIEW_HOME  override the log root (default ~/.dsh/doc-review)
  DSH_SESSION_ID   recorded automatically when present
"""
from __future__ import annotations

import argparse
import json
import os
from collections import Counter
from datetime import datetime
from pathlib import Path

HOME = Path(os.environ.get("DOC_REVIEW_HOME") or (Path.home() / ".dsh" / "doc-review")).expanduser()
LOG_DIR = HOME / "logs"


def now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def clip(value, limit: int = 500):
    if value is None:
        return None
    text = str(value).replace("\n", " ").strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def ops_summary(raw) -> list[dict]:
    if not raw:
        return []
    if isinstance(raw, str):
        path = Path(raw).expanduser()
        if path.is_file():
            raw = path.read_text("utf-8", errors="ignore")
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            return [{"raw": clip(raw, 200)}]
    if isinstance(raw, dict):
        raw = raw.get("operations") or raw.get("ops") or [raw]
    summary = []
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict):
            continue
        summary.append(
            {
                "op": item.get("op") or item.get("type"),
                "scope": item.get("scope"),
                "find": clip(item.get("find"), 120),
                "replace": clip(item.get("replace") or item.get("text"), 120),
                "all": bool(item.get("all")),
            }
        )
    return summary


def log_dirs() -> list[Path]:
    """Central plugin log, then a workspace fallback when the sandbox denies it."""
    dirs = [LOG_DIR]
    fallback = Path.cwd() / ".dsh-doc-review" / "logs"
    try:
        if fallback.resolve() != LOG_DIR.resolve():
            dirs.append(fallback)
    except OSError:
        dirs.append(fallback)
    return dirs


def append(event: dict) -> Path:
    line = json.dumps(event, ensure_ascii=False, separators=(",", ":")) + "\n"
    month = datetime.now().strftime("%Y-%m")
    last_error = None
    for directory in log_dirs():
        try:
            directory.mkdir(parents=True, exist_ok=True)
            path = directory / f"events-{month}.jsonl"
            fd = os.open(path, os.O_APPEND | os.O_CREAT | os.O_WRONLY, 0o644)
            try:
                os.write(fd, line.encode("utf-8"))
            finally:
                os.close(fd)
            return path
        except OSError as exc:
            last_error = exc
    raise last_error or OSError("no writable log directory")


def log_event(**fields) -> Path:
    """Append one event; used by helper scripts and by the agent's CLI calls."""
    event = {"ts": now()}
    event.update({key: value for key, value in fields.items() if value is not None})
    return append(event)


def cmd_log(args) -> None:
    event = {
        "ts": now(),
        "event": args.event,
        "contract": args.contract,
        "session": args.session or os.environ.get("DSH_SESSION_ID"),
        "file": args.file,
        "file_hash": args.hash,
        "stage": args.stage,
        "code": args.code,
        "reason": clip(args.reason),
        "note": clip(args.note),
        "fallback": args.fallback,
        "tool": args.tool,
        "duration_ms": args.duration_ms,
        "ops": ops_summary(args.ops),
    }
    event = {key: value for key, value in event.items() if value is not None}
    path = append(event)
    print(f"logged {args.event} -> {path}")


def cmd_tail(args) -> None:
    files = [path for directory in log_dirs() if directory.exists() for path in sorted(directory.glob("events-*.jsonl"))]
    if not files:
        print(f"no log under {LOG_DIR} or workspace fallback")
        return
    lines = files[-1].read_text("utf-8", errors="ignore").splitlines()[-args.n :]
    for line in lines:
        try:
            print(json.dumps(json.loads(line), ensure_ascii=False, indent=2))
        except json.JSONDecodeError:
            print(line)


def cmd_summary(args) -> None:
    files = [path for directory in log_dirs() if directory.exists() for path in sorted(directory.glob("events-*.jsonl"))]
    if not files:
        print(f"no log under {LOG_DIR} or workspace fallback")
        return
    events: Counter = Counter()
    codes: Counter = Counter()
    documents = set()
    total = 0
    for path in files:
        for line in path.read_text("utf-8", errors="ignore").splitlines():
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                continue
            total += 1
            events[event.get("event", "?")] += 1
            if event.get("code"):
                codes[event["code"]] += 1
            if event.get("file"):
                documents.add(event["file"])
    print("log dirs: " + ", ".join(str(d) for d in log_dirs()))
    print(f"events: {total}  files: {len(documents)}")
    for name, count in events.most_common():
        print(f"  {name}: {count}")
    if codes:
        print("codes:")
        for name, count in codes.most_common():
            print(f"  {name}: {count}")


def main() -> None:
    parser = argparse.ArgumentParser(description="dsh-doc-review failure log")
    sub = parser.add_subparsers(dest="command", required=True)

    log = sub.add_parser("log", help="append one event")
    log.add_argument("--event", required=True, help="helper_ok | helper_fail | fallback_script | manual_intervention | verify_fail")
    log.add_argument("--contract")
    log.add_argument("--session")
    log.add_argument("--file")
    log.add_argument("--hash")
    log.add_argument("--stage", help="preflight | match | edit | verify | write | ui")
    log.add_argument("--code", help="not_found | ambiguous | unsupported | verify_failed | exception | timeout")
    log.add_argument("--reason")
    log.add_argument("--note")
    log.add_argument("--fallback", help="one_off_script | manual | ask_user")
    log.add_argument("--tool")
    log.add_argument("--duration-ms", type=int)
    log.add_argument("--ops", help="ops JSON file or inline JSON")
    log.set_defaults(func=cmd_log)

    tail = sub.add_parser("tail", help="show recent events")
    tail.add_argument("--n", type=int, default=20)
    tail.set_defaults(func=cmd_tail)

    summary = sub.add_parser("summary", help="count events and codes")
    summary.set_defaults(func=cmd_summary)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
