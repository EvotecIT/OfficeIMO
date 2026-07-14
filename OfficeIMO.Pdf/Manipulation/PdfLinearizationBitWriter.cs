namespace OfficeIMO.Pdf;

/// <summary>Writes the bit-packed integer fields used by PDF linearization hint tables.</summary>
internal sealed class PdfLinearizationBitWriter {
    private readonly List<byte> _bytes = new List<byte>();
    private int _bitOffset;

    internal void Write(uint value, int bitCount) {
        if (bitCount < 0 || bitCount > 32) throw new ArgumentOutOfRangeException(nameof(bitCount));
        if (bitCount < 32 && bitCount > 0 && value >= (1U << bitCount)) {
            throw new ArgumentOutOfRangeException(nameof(value), value, "Value does not fit in the requested linearization hint width.");
        }

        for (int bit = bitCount - 1; bit >= 0; bit--) {
            if (_bitOffset == 0) _bytes.Add(0);
            if (((value >> bit) & 1U) != 0U) {
                int index = _bytes.Count - 1;
                _bytes[index] = (byte)(_bytes[index] | (1 << (7 - _bitOffset)));
            }

            _bitOffset = (_bitOffset + 1) & 7;
        }
    }

    internal void AlignToByte() {
        if (_bitOffset != 0) _bitOffset = 0;
    }

    internal byte[] ToArray() => _bytes.ToArray();
}
