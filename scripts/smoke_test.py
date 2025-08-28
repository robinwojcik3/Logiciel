import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

print("[SMOKE] Repo root:", ROOT)

try:
    import modules
    from modules import main_app
    print("[SMOKE] Imported modules.main_app OK")
except Exception as e:
    print("[SMOKE] Import error (modules/main_app):", e)
    sys.exit(1)

print("[SMOKE] OUT_IMG:", getattr(main_app, "OUT_IMG", None))
print("[SMOKE] DEFAULT_SHAPE_DIR:", getattr(main_app, "DEFAULT_SHAPE_DIR", None))

try:
    projs = main_app.discover_projects()
    print(f"[SMOKE] discover_projects -> {len(projs)} project(s)")
    if projs:
        print("[SMOKE] First project:", projs[0])
except Exception as e:
    print("[SMOKE] discover_projects error:", e)
    sys.exit(1)

print("[SMOKE] OK")

