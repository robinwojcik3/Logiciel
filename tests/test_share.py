import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from utils.fs import share_available


def test_share_available_true(tmp_path):
    d = tmp_path / "dir"
    d.mkdir()
    assert share_available(str(d))


def test_share_available_false():
    assert not share_available("/path/that/does/not/exist/xyz")
