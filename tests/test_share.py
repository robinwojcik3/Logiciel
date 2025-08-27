import os
import sys
import tempfile

sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from utils.fs import share_available


def test_share_available_true():
    with tempfile.TemporaryDirectory() as tmp:
        assert share_available(tmp)


def test_share_available_false():
    assert not share_available(os.path.join("/nonexistent", "path"))
