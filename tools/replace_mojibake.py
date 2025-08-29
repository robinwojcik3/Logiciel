#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Replace common CP1252/UTF-8 mojibake sequences by correct characters.
This is a surgical, mapping-based fixer to avoid over-conversion.
"""
from __future__ import annotations

from pathlib import Path

MAP = {
    # Accents (lower)
    "Ã ": "à", "Ã¡": "á", "Ã¢": "â", "Ã£": "ã", "Ã¤": "ä", "Ã¥": "å", "Ã¦": "æ", "Ã§": "ç",
    "Ã¨": "è", "Ã©": "é", "Ãª": "ê", "Ã«": "ë", "Ã¬": "ì", "Ã­": "í", "Ã®": "î", "Ã¯": "ï",
    "Ã°": "ð", "Ã±": "ñ", "Ã²": "ò", "Ã³": "ó", "Ã´": "ô", "Ãµ": "õ", "Ã¶": "ö", "Ã·": "÷",
    "Ã¸": "ø", "Ã¹": "ù", "Ãº": "ú", "Ã»": "û", "Ã¼": "ü", "Ã½": "ý", "Ã¾": "þ", "Ã¿": "ÿ",
    # Accents (upper)
    "Ã€": "À", "Ã�": "Á", "Ã‚": "Â", "Ãƒ": "Ã", "Ã„": "Ä", "Ã…": "Å", "Ã†": "Æ", "Ã‡": "Ç",
    "Ãˆ": "È", "Ã‰": "É", "ÃŠ": "Ê", "Ã‹": "Ë", "ÃŒ": "Ì", "Ã�": "Í", "ÃŽ": "Î", "Ã�": "Ï",
    "Ã�": "Ð", "Ã‘": "Ñ", "Ã’": "Ò", "Ã“": "Ó", "Ã”": "Ô", "Ã•": "Õ", "Ã–": "Ö", "Ã—": "×",
    "Ã˜": "Ø", "Ã™": "Ù", "Ãš": "Ú", "Ã›": "Û", "Ãœ": "Ü", "Ã�": "Ý", "Ãž": "Þ", "ÃŸ": "ß",
    # Guillemets, degré, points
    "Â«": "«", "Â»": "»", "Â°": "°", "Â·": "·", "Â": "",
    # Quotes/dashes/ellipsis
    "â€˜": "‘", "â€™": "’", "â€œ": "“", "â€	": "”", "â€“": "–", "â€”": "—", "â€¦": "…",
    # Arrows
    "â†’": "→", "â†�": "←", "â†‘": "↑", "â†“": "↓",
    # Common UI symbols
    "âœ–": "✖", "âœ”": "✔", "âœ•": "✕", "âœ“": "✓",
    # Emojis commonly used in this project
    "ðŸ“‚": "📂", "ðŸ“�": "📁", "ðŸ§ª": "🧪",
}


def fix_text(s: str) -> str:
    for k, v in MAP.items():
        s = s.replace(k, v)
    return s


def process(paths: list[Path]) -> int:
    changed = 0
    for p in paths:
        try:
            raw = p.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            raw = p.read_text(encoding="cp1252")
        fixed = fix_text(raw)
        if fixed != raw:
            p.write_text(fixed, encoding="utf-8")
            print(f"Fixed: {p}")
            changed += 1
    return changed


if __name__ == "__main__":
    targets = [
        Path("modules/main_app.py"),
        Path("modules/export_worker.py"),
        Path("modules/id_contexte_eco.py"),
        Path("modules/wikipedia_scraper.py"),
        Path("Start.py"),
        Path("README.md"),
        Path("Agent.md"),
        Path("docs/INSTALL.md"),
        Path("requirements.txt"),
        Path("tools/_head.txt"),
    ]
    n = process([p for p in targets if p.exists()])
    print(f"Done. Files fixed: {n}")
