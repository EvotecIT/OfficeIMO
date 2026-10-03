"""Inspect finite Decimal128 values in pinned numbers-parser fixture cells."""
from decimal import Decimal


def source_number(cell):
    """Return exact normalized storage text, portable float and precision flag."""
    assert cell._buffer[0] == 5 and cell._flags & 1
    raw = cell._buffer[12:28]
    assert len(raw) == 16 and raw[15] & 0x78 != 0x78
    coefficient = int.from_bytes(raw[:14], "little") + ((raw[14] & 1) << 112)
    assert coefficient < 10**34
    exponent = (((raw[15] & 0x7f) << 7) | (raw[14] >> 1)) - 0x1820
    if coefficient:
        while coefficient % 10 == 0:
            coefficient //= 10
            exponent += 1
        text = ("-" if raw[15] & 0x80 else "") + str(coefficient) + "E" + str(exponent)
    else:
        text = "0"
    return text, float(Decimal(text)), len(str(coefficient)) > 15
