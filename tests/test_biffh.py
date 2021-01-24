from io import StringIO
from xlrd import biffh


def test_hex_char_dump():
    sio = StringIO()
    sio.debug = sio.write
    biffh.hex_char_dump(b"abc\0e\01", 0, 6, logger=sio)
    s = sio.getvalue()
    assert "61 62 63 00 65 01" in s, s
    assert "abc~e?" in s, s

