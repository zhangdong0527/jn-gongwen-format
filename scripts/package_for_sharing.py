#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Package this skill into two distributable ZIP files.

Run from anywhere:
    python scripts/package_for_sharing.py

Produces, next to the skill folder:
    dist/jn-gongwen-format.zip          WorkBuddy variant (top-level skill folder)
    dist/jn-gongwen-format-doubao.zip   Doubao variant (SKILL.md at ZIP root)

Zero external dependencies.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parents[1]
NAME = SKILL_DIR.name
DIST = SKILL_DIR.parent / "dist"

EXCLUDE_DIRS = {"__pycache__", ".git", "dist"}
EXCLUDE_SUFFIXES = {".pyc", ".pyo"}


def source_files() -> list[Path]:
    files = []
    for path in sorted(SKILL_DIR.rglob("*")):
        if not path.is_file():
            continue
        if any(part in EXCLUDE_DIRS for part in path.relative_to(SKILL_DIR).parts):
            continue
        if path.suffix in EXCLUDE_SUFFIXES:
            continue
        files.append(path)
    return files


def build(out_path: Path, top_level: bool) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in source_files():
            rel = path.relative_to(SKILL_DIR).as_posix()
            arcname = f"{NAME}/{rel}" if top_level else rel
            zf.write(path, arcname)
    with zipfile.ZipFile(out_path) as zf:
        assert zf.testzip() is None, f"ZIP 校验失败: {out_path}"
        names = zf.namelist()
    entry = "SKILL.md" if not top_level else f"{NAME}/SKILL.md"
    root_ok = entry in names
    print(f"已生成: {out_path} ({len(names)} 个文件, 入口 {entry}: {'正常' if root_ok else '缺失'})")


def main() -> int:
    build(DIST / f"{NAME}.zip", top_level=True)
    build(DIST / f"{NAME}-doubao.zip", top_level=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
