#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Normalize UTF-8/CP1252 mojibake in text files safely.

Only fixes clear mojibake markers like:
  - 'Ã…', 'Ã©', 'Â«', 'Â»'
  - 'â€™', 'â€“', 'â€”', 'â€¦'
  - 'ðŸ' (broken emoji)
  - the replacement char '�'

It will not touch normal accented French letters (é, è, à, â, ç, œ, …).

Usage:
  python tools/normalize_text_encoding.py [paths...]
If no paths given, scans the repository for common text files.
"""
from __future__ import annotations

import sys
from pathlib import Path

SUSPECT_SINGLE = ("Ã", "Â")
SUSPECT_PAIRS = ("â€™", "â€“", "â€”", "â€¦", "ðŸ")
TEXT_EXT = {".py", ".md", ".txt", ".json", ".ini", ".cfg", ".toml", ".yaml", ".yml"}


def looks_mojibake(s: str) -> int:
    score = sum(s.count(c) for c in SUSPECT_SINGLE)
    score += sum(s.count(c) for c in SUSPECT_PAIRS)
    score += s.count("�")  # replacement char
    return score


def fix_text(s: str) -> str:
    out: list[str] = []
    buf: list[str] = []
    def flush():
        nonlocal out, buf
        if not buf:
            return
        seg = "".join(buf)
        try:
            out.append(seg.encode("cp1252", errors="strict").decode("utf-8", errors="strict"))
        except Exception:
            out.append(seg)
        buf = []
    for ch in s:
        if ch != "\ufeff" and ord(ch) <= 255:
            buf.append(ch)
        else:
            flush()
            out.append(ch)
    flush()
    return "".join(out)


def process_file(path: Path) -> bool:
    try:
        raw = path.read_bytes()
        s = raw.decode("utf-8")
    except UnicodeDecodeError:
        try:
            s = raw.decode("cp1252")
        except Exception:
            return False

    before = looks_mojibake(s)
    if before == 0:
        return False

    fixed = fix_text(s)
    after = looks_mojibake(fixed)
    if after >= before:
        return False

    # Preserve original newline style
    nl = "\r\n" if "\r\n" in s else "\n"
    out = fixed.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    path.write_text(nl.join(out), encoding="utf-8")
    return True


def iter_targets(root: Path):
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        if any(part in {".git", ".venv", "__pycache__"} for part in p.parts):
            continue
        if p.suffix.lower() in TEXT_EXT:
            yield p


def main(argv: list[str]) -> int:
    if len(argv) > 1:
        targets: list[Path] = []
        for arg in argv[1:]:
            p = Path(arg)
            if p.is_dir():
                targets.extend(iter_targets(p))
            elif p.is_file():
                targets.append(p)
        # dedupe
        seen = set()
        uniq: list[Path] = []
        for t in targets:
            if t not in seen:
                uniq.append(t); seen.add(t)
        targets = uniq
    else:
        targets = list(iter_targets(Path.cwd()))

    changed = 0
    for p in targets:
        try:
            if process_file(p):
                changed += 1
                print(f"Fixed: {p}")
        except Exception as e:
            print(f"Skip {p}: {e}")
    print(f"Done. Files fixed: {changed}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
