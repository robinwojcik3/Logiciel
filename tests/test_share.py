import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from utils.fs import share_available


def test_share_available(tmp_path):
    assert share_available(str(tmp_path))


def test_share_unavailable():
    assert not share_available("/this/path/does/not/exist")
