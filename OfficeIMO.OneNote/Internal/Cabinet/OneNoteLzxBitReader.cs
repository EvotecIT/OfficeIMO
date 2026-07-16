namespace OfficeIMO.OneNote;

/// <summary>
/// Reads LZX fields from little-endian 16-bit words while exposing the byte-aligned
/// mode used by uncompressed blocks.
/// </summary>
internal sealed class OneNoteLzxBitReader {
    private readonly byte[] _data;
    private int _nextWordOffset;
    private ushort _word;
    private int _bitsRemaining;

    internal OneNoteLzxBitReader(byte[] data) {
        _data = data ?? throw new ArgumentNullException(nameof(data));
    }

    internal int RemainingByteCount {
        get {
            EnsureByteAligned();
            return _data.Length - _nextWordOffset;
        }
    }

    internal uint ReadBits(int count) {
        if (count < 0 || count > 32) throw new ArgumentOutOfRangeException(nameof(count));
        uint value = 0;
        int remaining = count;
        while (remaining > 0) {
            if (_bitsRemaining == 0) LoadWord();
            int take = Math.Min(remaining, _bitsRemaining);
            int shift = _bitsRemaining - take;
            uint mask = take == 32 ? uint.MaxValue : (1u << take) - 1u;
            value = (value << take) | ((uint)(_word >> shift) & mask);
            _bitsRemaining -= take;
            remaining -= take;
        }
        return value;
    }

    internal void AlignToWord() {
        _bitsRemaining = 0;
    }

    internal byte ReadRawByte() {
        EnsureByteAligned();
        EnsureRawBytes(1);
        return _data[_nextWordOffset++];
    }

    internal uint ReadRawUInt32() {
        EnsureByteAligned();
        EnsureRawBytes(4);
        uint value = (uint)(_data[_nextWordOffset] |
            (_data[_nextWordOffset + 1] << 8) |
            (_data[_nextWordOffset + 2] << 16) |
            (_data[_nextWordOffset + 3] << 24));
        _nextWordOffset += 4;
        return value;
    }

    internal void CopyRawBytes(byte[] destination, int destinationOffset, int count) {
        EnsureByteAligned();
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (destinationOffset < 0 || count < 0 || destinationOffset > destination.Length - count) {
            throw new ArgumentOutOfRangeException(nameof(destinationOffset));
        }
        EnsureRawBytes(count);
        Buffer.BlockCopy(_data, _nextWordOffset, destination, destinationOffset, count);
        _nextWordOffset += count;
    }

    private void LoadWord() {
        if (_nextWordOffset > _data.Length - 2) {
            throw new OneNoteFormatException("ONENOTE_CAB_LZX_TRUNCATED", "The CAB LZX bitstream ended inside a 16-bit word.");
        }
        _word = (ushort)(_data[_nextWordOffset] | (_data[_nextWordOffset + 1] << 8));
        _nextWordOffset += 2;
        _bitsRemaining = 16;
    }

    private void EnsureByteAligned() {
        if (_bitsRemaining != 0) {
            throw new OneNoteFormatException("ONENOTE_CAB_LZX_CORRUPT", "The CAB LZX stream entered byte mode without word alignment.");
        }
    }

    private void EnsureRawBytes(int count) {
        if (count < 0 || _nextWordOffset > _data.Length - count) {
            throw new OneNoteFormatException("ONENOTE_CAB_LZX_TRUNCATED", "The CAB LZX byte stream ended unexpectedly.");
        }
    }
}
