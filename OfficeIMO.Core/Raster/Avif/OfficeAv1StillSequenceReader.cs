using System;

namespace OfficeIMO.Drawing;

/// <summary>Checks low-overhead OBU boundaries and 8-bit Main reduced headers, without decoding or allocating planes.</summary>
internal static class OfficeAv1StillSequenceReader {
    internal static bool TryRead(byte[] bytes, OfficeAvifImageItem item, OfficeRasterDecodeOptions options,
        out OfficeAv1StillSequence? sequence) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate();
        options.CancellationToken.ThrowIfCancellationRequested();
        sequence = null;
        if (bytes == null || item == null || bytes.Length > options.MaximumEncodedBytes) return false;
        try {
            Require(item.Offset >= 0 && item.Length > 0 && item.Offset <= bytes.Length - item.Length);
            int end = item.Offset + item.Length;
            int p = item.Offset, count = 0;
            OfficeAv1StillSequence? result = null;
            bool hasFrame = false;
            while (p < end) {
                options.CancellationToken.ThrowIfCancellationRequested();
                Require(++count <= 4096);
                int header = bytes[p++];
                Require((header & 0x81) == 0); // Forbidden and reserved bits.
                int type = (header >> 3) & 15;
                if ((header & 4) != 0) {
                    Require(p < end && bytes[p++] == 0); // Single temporal/spatial layer and reserved bits zero.
                }
                uint size = (header & 2) != 0 ? Leb128(bytes, ref p, end) : (uint)(end - p);
                Require(size <= (uint)(end - p));
                int next = checked(p + (int)size);
                switch (type) {
                    case 1:
                        Require(result == null && !hasFrame);
                        result = ReadSequence(bytes, p, (int)size, item, options);
                        break;
                    case 2:
                        Require(size == 0 && !hasFrame);
                        break;
                    case 6:
                        Require(result != null && !hasFrame && size > 0);
                        result!.FrameOffset = p;
                        result.FrameLength = (int)size;
                        hasFrame = true;
                        break;
                    case 3: case 4: case 5: case 7: case 8:
                        throw new FormatException("Separate tile groups, metadata and layered frames require another AV1 path.");
                    // Padding and reserved OBUs have no picture semantics in this path.
                }
                p = next;
            }
            Require(result != null && hasFrame);
            sequence = result;
            return true;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    private static uint Leb128(byte[] bytes, ref int p, int end) {
        ulong value = 0;
        for (int i = 0; i < 8; i++) {
            Require(p < end);
            int b = bytes[p++];
            value |= (ulong)(b & 127) << (i * 7);
            if ((b & 128) == 0) {
                Require(value <= uint.MaxValue);
                return (uint)value;
            }
        }
        throw new FormatException("Unterminated AV1 OBU size.");
    }

    private static OfficeAv1StillSequence ReadSequence(byte[] bytes, int offset, int length,
        OfficeAvifImageItem item, OfficeRasterDecodeOptions options) {
        var bits = new Bits(bytes, offset, length);
        Require(bits.Read(3) == 0 && bits.Flag() && bits.Flag()); // Main, still picture, reduced header.
        var result = new OfficeAv1StillSequence { Level = bits.Read(5) };
        result.WidthBits = bits.Read(4) + 1;
        result.HeightBits = bits.Read(4) + 1;
        result.MaximumWidth = bits.Read(result.WidthBits) + 1;
        result.MaximumHeight = bits.Read(result.HeightBits) + 1;
        Require(OfficeRasterGuards.TryEnsurePixelCount(result.MaximumWidth, result.MaximumHeight,
            options.MaximumDecodedPixels, out _));
        // Actual UpscaledWidth/FrameHeight are checked against ispe by the frame parser, not inferred from maxima.
        Require(item.Width <= result.MaximumWidth && item.Height <= result.MaximumHeight);
        result.Use128Superblock = bits.Flag();
        result.FilterIntra = bits.Flag();
        result.IntraEdgeFilter = bits.Flag();
        result.SuperResolution = bits.Flag();
        result.Cdef = bits.Flag();
        result.Restoration = bits.Flag();
        Require(!bits.Flag()); // high_bitdepth: this stage owns 8-bit planes.
        result.Monochrome = bits.Flag();
        int primaries = 2, transfer = 2, matrix = 2;
        if (bits.Flag()) { primaries = bits.Read(8); transfer = bits.Read(8); matrix = bits.Read(8); }
        Require(matrix != 0); // MC_IDENTITY requires both subsampling flags zero, including for monochrome streams.
        bool fullRange;
        if (result.Monochrome) {
            fullRange = bits.Flag();
        } else {
            fullRange = bits.Flag();
            result.ChromaSamplePosition = bits.Read(2);
            Require(result.ChromaSamplePosition != 3);
            result.SeparateUvDeltaQ = bits.Flag();
        }
        result.Color = new OfficeAvifColorDescription(primaries, transfer, matrix, fullRange);
        result.FilmGrain = bits.Flag();
        bits.TrailingBits();
        byte[] config = item.Configuration;
        Require(config.Length == 4 && config[0] == 0x81 && config[1] == result.Level);
        int format = (result.Monochrome ? 16 : 0) | 12 | result.ChromaSamplePosition;
        Require(config[2] == format && config[3] == 0);
        Require(result.Monochrome == item.Monochrome && (!item.Monochrome || fullRange));
        return result;
    }

    private static void Require(bool condition) {
        if (!condition) throw new FormatException("Invalid or unsupported AV1 still sequence.");
    }

    /// <summary>Most-significant-bit-first header cursor bounded to one OBU, never to the surrounding item.</summary>
    private sealed class Bits {
        private readonly byte[] _bytes;
        private readonly int _start;
        private readonly int _length;
        private int _position;
        internal Bits(byte[] bytes, int start, int length) { _bytes = bytes; _start = start; _length = checked(length * 8); }
        internal int Read(int count) {
            Require(count >= 0 && count <= 16 && _position <= _length - count);
            int value = 0;
            for (int i = 0; i < count; i++) {
                value = value << 1 | ((_bytes[_start + (_position >> 3)] >> (7 - (_position & 7))) & 1);
                _position++;
            }
            return value;
        }
        internal bool Flag() => Read(1) != 0;
        internal void TrailingBits() {
            Require(_length - _position <= 8); // Reduced header uses alignment bits, not an unbounded zero-padding scan.
            Require(Flag());
            while (_position < _length) Require(Read(1) == 0);
        }
    }
}
