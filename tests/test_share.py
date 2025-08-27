from utils.fs import share_available


def test_share_available_existing(tmp_path):
    assert share_available(str(tmp_path))


def test_share_available_missing():
    assert not share_available("/path/that/does/not/exist")
