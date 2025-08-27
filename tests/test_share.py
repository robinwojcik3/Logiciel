import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from utils.fs import share_available


def test_share_available(tmp_path):
    assert share_available(str(tmp_path))


def test_share_unavailable(tmp_path):
    nonexist = tmp_path / "missing"
    assert not share_available(str(nonexist))
