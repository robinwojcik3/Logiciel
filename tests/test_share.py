from utils.fs import share_available


def test_share_available(tmp_path):
    existing = tmp_path / "dir"
    existing.mkdir()
    assert share_available(str(existing))
    assert not share_available(str(existing / "missing"))
