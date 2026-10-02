using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>MSB-first syntax cursor bounded to one OBU payload; entropy-coded tiles use a separate reader.</summary>
internal sealed class OfficeAv1Bits {
    private readonly byte[] _bytes;
    private readonly int _start;
    private readonly int _length;
    private readonly CancellationToken _cancellation;
    private int _position;

    internal OfficeAv1Bits(byte[] bytes, int start, int length, CancellationToken cancellation) {
        if (start < 0 || length < 0 || start > bytes.Length - length) throw new FormatException("Invalid AV1 payload bounds.");
        _bytes = bytes; _start = start; _length = checked(length * 8); _cancellation = cancellation;
    }

    internal int Read(int count) {
        _cancellation.ThrowIfCancellationRequested();
        if (count < 0 || count > 16 || _position > _length - count) throw new FormatException("Truncated AV1 syntax.");
        int value = 0;
        for (int i = 0; i < count; i++) {
            value = value << 1 | ((_bytes[_start + (_position >> 3)] >> (7 - (_position & 7))) & 1);
            _position++;
        }
        return value;
    }

    internal bool Flag() => Read(1) != 0;
    internal int Signed(int count) {
        int value = Read(count);
        return (value & (1 << (count - 1))) != 0 ? value - (1 << count) : value;
    }

    /// <summary>AV1 ns(n): the short or long codeword for an integer in [0,n).</summary>
    internal int NonSymmetric(int n) {
        if (n < 1 || n > 65536) throw new FormatException("Invalid AV1 symbol range.");
        int width = 0;
        while ((1 << width) <= n) width++;
        int threshold = (1 << width) - n;
        int value = Read(width - 1);
        return value < threshold ? value : (value << 1) - threshold + Read(1);
    }

    internal int ByteOffset {
        get {
            if ((_position & 7) != 0) throw new FormatException("Unaligned AV1 payload.");
            return _start + (_position >> 3);
        }
    }

    internal void AlignZero() {
        while ((_position & 7) != 0)
            if (Flag()) throw new FormatException("Nonzero AV1 alignment bit.");
    }

    internal void TrailingBits() {
        if (_length - _position > 8 || !Flag()) throw new FormatException("Invalid AV1 trailing bits.");
        while (_position < _length)
            if (Flag()) throw new FormatException("Nonzero AV1 trailing bit.");
    }
}
