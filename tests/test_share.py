import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from utils.fs import share_available


def test_share_available(tmp_path):
    d = tmp_path / "dir"
    d.mkdir()
    assert share_available(str(d))
    assert not share_available(str(d / "missing"))
