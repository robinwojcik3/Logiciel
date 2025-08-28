# -*- coding: utf-8 -*-
"""CLI pour lancer le worker QGIS dans un sous-processus QGIS Python.

Usage: python -m modules.export_worker_cli <input.json>
L'input JSON doit contenir: {"projects": [...], "cfg": {...}}
Sortie: imprime "OK,KO" sur stdout et un log minimal sur stderr en cas d'erreur.
"""
from __future__ import annotations
import sys
import json
import traceback


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print("Usage: python -m modules.export_worker_cli <input.json>", file=sys.stderr)
        return 2
    in_path = argv[1]
    try:
        with open(in_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        projects = data.get("projects") or []
        cfg = data.get("cfg") or {}
        from .export_worker import worker_run
        ok, ko = worker_run((projects, cfg))
        print(f"{ok},{ko}")
        return 0
    except SystemExit as e:
        raise
    except Exception:
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

