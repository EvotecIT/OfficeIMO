namespace OfficeIMO.Archives.Compression;

/// <summary>Reads a Zstandard backward bitstream (RFC 8878, section 4.1): the stream is a little-endian number whose
/// highest set bit marks its end, and fields are read from that end towards the first byte.</summary>
internal sealed class ZstandardBackwardBits {
    private byte[] _data = Array.Empty<byte>();
    private int _start;
    private int _end;
    private long _position;

    /// <summary>Bits left before the start of the stream. Negative after reading past the start.</summary>
    internal long Remaining => _position;

    internal bool Overflowed => _position < 0;

    internal void Reset(byte[] data, int start, int length, long inputOffset, string structure) {
        if (length <= 0) throw ZstandardDecoder.Corrupt(inputOffset, "a non-empty " + structure);
        byte last = data[start + length - 1];
        if (last == 0) throw ZstandardDecoder.Corrupt(inputOffset + length - 1, "the end marker bit of the " + structure);
        _data = data;
        _start = start;
        _end = start + length;
        _position = (long)(length - 1) * 8 + ZstandardDecoder.HighBit(last);
    }

    internal uint Read(int count) {
        if (count == 0) return 0;
        _position -= count;
        return Extract(_position, count);
    }

    internal uint Peek(int count) => Extract(_position - count, count);

    internal void Skip(int count) => _position -= count;

    private uint Extract(long bit, int count) {
        ulong mask = (1UL << count) - 1;
        if (bit >= 0) return (uint)((Load(_start + (int)(bit >> 3)) >> (int)(bit & 7)) & mask);
        if (bit + count <= 0) return 0;
        return (uint)((Load(_start) << (int)-bit) & mask);
    }

    private ulong Load(int index) {
        byte[] data = _data;
        if (index + 8 <= _end) {
            return data[index] | (ulong)data[index + 1] << 8 | (ulong)data[index + 2] << 16 | (ulong)data[index + 3] << 24
                | (ulong)data[index + 4] << 32 | (ulong)data[index + 5] << 40 | (ulong)data[index + 6] << 48 | (ulong)data[index + 7] << 56;
        }
        ulong value = 0;
        for (int shift = 0; index < _end; index++, shift += 8) value |= (ulong)data[index] << shift;
        return value;
    }
}

/// <summary>Reads little-endian bit fields forward from a byte range, as FSE table descriptions are stored.</summary>
internal struct ZstandardForwardBits {
    private readonly byte[] _data;
    private readonly int _start;
    private readonly int _end;

    internal ZstandardForwardBits(byte[] data, int start, int end) {
        _data = data;
        _start = start;
        _end = end;
        Position = 0;
    }

    internal long Position { get; private set; }

    internal int BytesConsumed => (int)((Position + 7) >> 3);

    internal uint Peek(int count) {
        int index = _start + (int)(Position >> 3);
        ulong value = 0;
        for (int shift = 0; shift < 64 && index < _end; index++, shift += 8) value |= (ulong)_data[index] << shift;
        return (uint)((value >> (int)(Position & 7)) & ((1UL << count) - 1));
    }

    internal uint Read(int count) {
        uint value = Peek(count);
        Position += count;
        return value;
    }

    internal void Skip(int count) => Position += count;
}
