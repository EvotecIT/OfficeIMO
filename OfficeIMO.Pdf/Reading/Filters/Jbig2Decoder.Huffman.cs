namespace OfficeIMO.Pdf.Filters;

internal static partial class Jbig2Decoder {
    /// <summary>Most-significant-bit-first reader for Huffman-coded JBIG2 data (T.88 6.2.2, B.3).</summary>
    private sealed class HuffmanReader {
        private readonly byte[] _data;
        private readonly int _end;
        private long _bit;
        private readonly Action<long> _charge;

        internal HuffmanReader(byte[] data, int start, int end, Action<long> charge) {
            _data = data;
            _end = end;
            _bit = (long)start * 8;
            _charge = charge;
        }

        internal int BytePosition => (int)((_bit + 7) / 8);
        internal int End => _end;

        internal int ReadBit() {
            long index = _bit >> 3;
            if (index >= _end) throw new InvalidDataException("Truncated JBIG2 Huffman data.");
            int value = (_data[index] >> (7 - (int)(_bit & 7))) & 1;
            _bit++;
            return value;
        }

        internal uint ReadBits(int count) {
            _charge(count);
            uint value = 0;
            for (int index = 0; index < count; index++) value = (value << 1) | (uint)ReadBit();
            return value;
        }

        internal void Align() => _bit = (_bit + 7) & ~7L;

        internal void SkipBytes(int count) {
            Align();
            if (count < 0 || (_bit >> 3) + count > _end) throw new InvalidDataException("Truncated JBIG2 Huffman data.");
            _bit += (long)count * 8;
        }
    }

    /// <summary>One table line (B.2): prefix length, range length and range low, plus its role.</summary>
    private readonly struct HuffmanLine {
        internal HuffmanLine(int prefixLength, int rangeLength, int rangeLow, int kind = 0) {
            PrefixLength = prefixLength;
            RangeLength = rangeLength;
            RangeLow = rangeLow;
            Kind = kind;
        }

        internal int PrefixLength { get; }
        internal int RangeLength { get; }
        internal int RangeLow { get; }
        internal int Kind { get; } // 0 normal, 1 lower range, 2 upper range, 3 out-of-band
    }

    /// <summary>Canonical Huffman table built with the code assignment of B.3.</summary>
    private sealed class HuffmanTable {
        private readonly HuffmanLine[] _lines;
        private readonly int[] _codes;
        private readonly int _maximumLength;

        internal HuffmanTable(HuffmanLine[] lines) {
            _lines = lines;
            _codes = new int[lines.Length];
            int maximum = 0;
            foreach (HuffmanLine line in lines) {
                if (line.PrefixLength < 0 || line.PrefixLength > 32 || line.RangeLength < 0 || line.RangeLength > 32)
                    throw new InvalidDataException("JBIG2 Huffman table line is invalid.");
                maximum = Math.Max(maximum, line.PrefixLength);
            }
            _maximumLength = maximum;
            var counts = new int[maximum + 2];
            foreach (HuffmanLine line in lines) counts[line.PrefixLength]++;
            counts[0] = 0;
            int firstCode = 0;
            for (int length = 1; length <= maximum; length++) {
                firstCode = (firstCode + counts[length - 1]) << 1;
                int code = firstCode;
                for (int index = 0; index < lines.Length; index++) {
                    if (lines[index].PrefixLength == length) _codes[index] = code++;
                }
            }
        }

        /// <summary>Decodes one value; null is out-of-band.</summary>
        internal int? Decode(HuffmanReader reader) {
            int code = 0;
            for (int length = 1; length <= _maximumLength; length++) {
                code = (code << 1) | reader.ReadBit();
                for (int index = 0; index < _lines.Length; index++) {
                    HuffmanLine line = _lines[index];
                    if (line.PrefixLength != length || _codes[index] != code) continue;
                    long value = line.Kind switch {
                        3 => 0,
                        1 => (long)line.RangeLow - reader.ReadBits(32),
                        2 => (long)line.RangeLow + reader.ReadBits(32),
                        _ => (long)line.RangeLow + reader.ReadBits(line.RangeLength)
                    };
                    if (line.Kind == 3) return null;
                    if (value < int.MinValue || value > int.MaxValue) throw new InvalidDataException("JBIG2 Huffman value overflows.");
                    return (int)value;
                }
            }
            throw new InvalidDataException("Invalid JBIG2 Huffman code.");
        }

        internal int DecodeValue(HuffmanReader reader, string name) =>
            Decode(reader) ?? throw new InvalidDataException("JBIG2 " + name + " is out of band.");
    }

    private static HuffmanLine L(int prefix, int range, int low) => new(prefix, range, low);
    private static HuffmanLine Lower(int prefix, int low) => new(prefix, 32, low, 1);
    private static HuffmanLine Upper(int prefix, int low) => new(prefix, 32, low, 2);
    private static HuffmanLine Oob(int prefix) => new(prefix, 0, 0, 3);

    // Standard Huffman tables B.1 to B.15 (Annex B.5).
    private static readonly Lazy<HuffmanTable[]> StandardTables = new(() => new[] {
        new HuffmanTable(new[] { L(1, 4, 0), L(2, 8, 16), L(3, 16, 272), Upper(3, 65808) }),
        new HuffmanTable(new[] { L(1, 0, 0), L(2, 0, 1), L(3, 0, 2), L(4, 3, 3), L(5, 6, 11), Upper(6, 75), Oob(6) }),
        new HuffmanTable(new[] { L(8, 8, -256), L(1, 0, 0), L(2, 0, 1), L(3, 0, 2), L(4, 3, 3), L(5, 6, 11), Lower(8, -257), Upper(7, 75), Oob(6) }),
        new HuffmanTable(new[] { L(1, 0, 1), L(2, 0, 2), L(3, 0, 3), L(4, 3, 4), L(5, 6, 12), Upper(5, 76) }),
        new HuffmanTable(new[] { L(7, 8, -255), L(1, 0, 1), L(2, 0, 2), L(3, 0, 3), L(4, 3, 4), L(5, 6, 12), Lower(7, -256), Upper(6, 76) }),
        new HuffmanTable(new[] {
            L(5, 10, -2048), L(4, 9, -1024), L(4, 8, -512), L(4, 7, -256), L(5, 6, -128), L(5, 5, -64), L(4, 5, -32), L(2, 7, 0),
            L(3, 7, 128), L(3, 8, 256), L(4, 9, 512), L(4, 10, 1024), Lower(6, -2049), Upper(6, 2048) }),
        new HuffmanTable(new[] {
            L(4, 9, -1024), L(3, 8, -512), L(4, 7, -256), L(5, 6, -128), L(5, 5, -64), L(4, 5, -32), L(4, 5, 0), L(5, 5, 32),
            L(5, 6, 64), L(4, 7, 128), L(3, 8, 256), L(3, 9, 512), L(3, 10, 1024), Lower(5, -1025), Upper(5, 2048) }),
        new HuffmanTable(new[] {
            L(8, 3, -15), L(9, 1, -7), L(8, 1, -5), L(9, 0, -3), L(7, 0, -2), L(4, 0, -1), L(2, 1, 0), L(5, 0, 2), L(6, 0, 3),
            L(3, 4, 4), L(6, 1, 20), L(4, 4, 22), L(4, 5, 38), L(5, 6, 70), L(5, 7, 134), L(6, 7, 262), L(7, 8, 390),
            L(6, 10, 646), Lower(9, -16), Upper(9, 1670), Oob(2) }),
        new HuffmanTable(new[] {
            L(8, 4, -31), L(9, 2, -15), L(8, 2, -11), L(9, 1, -7), L(7, 1, -5), L(4, 1, -3), L(3, 1, -1), L(3, 1, 1), L(5, 1, 3),
            L(6, 1, 5), L(3, 5, 7), L(6, 2, 39), L(4, 5, 43), L(4, 6, 75), L(5, 7, 139), L(5, 8, 267), L(6, 8, 523),
            L(7, 9, 779), L(6, 11, 1291), Lower(9, -32), Upper(9, 3339), Oob(2) }),
        new HuffmanTable(new[] {
            L(7, 4, -21), L(8, 0, -5), L(7, 0, -4), L(5, 0, -3), L(2, 2, -2), L(5, 0, 2), L(6, 0, 3), L(7, 0, 4), L(8, 0, 5),
            L(2, 6, 6), L(5, 5, 70), L(6, 5, 102), L(6, 6, 134), L(6, 7, 198), L(6, 8, 326), L(6, 9, 582), L(6, 10, 1094),
            L(7, 11, 2118), Lower(8, -22), Upper(8, 4166), Oob(2) }),
        new HuffmanTable(new[] {
            L(1, 0, 1), L(2, 1, 2), L(4, 0, 4), L(4, 1, 5), L(5, 1, 7), L(5, 2, 9), L(6, 2, 13), L(7, 2, 17), L(7, 3, 21),
            L(7, 4, 29), L(7, 5, 45), L(7, 6, 77), Upper(7, 141) }),
        new HuffmanTable(new[] {
            L(1, 0, 1), L(2, 0, 2), L(3, 1, 3), L(5, 0, 5), L(5, 1, 6), L(6, 1, 8), L(7, 0, 10), L(7, 1, 11), L(7, 2, 13),
            L(7, 3, 17), L(7, 4, 25), L(8, 5, 41), Upper(8, 73) }),
        new HuffmanTable(new[] {
            L(1, 0, 1), L(3, 0, 2), L(4, 0, 3), L(5, 0, 4), L(4, 1, 5), L(3, 3, 7), L(6, 1, 15), L(6, 2, 17), L(6, 3, 21),
            L(6, 4, 29), L(6, 5, 45), L(7, 6, 77), Upper(7, 141) }),
        new HuffmanTable(new[] { L(3, 0, -2), L(3, 0, -1), L(1, 0, 0), L(3, 0, 1), L(3, 0, 2) }),
        new HuffmanTable(new[] {
            L(7, 4, -24), L(6, 2, -8), L(5, 1, -4), L(4, 0, -2), L(3, 0, -1), L(1, 0, 0), L(3, 0, 1), L(4, 0, 2), L(5, 1, 3),
            L(6, 2, 5), L(7, 4, 9), Lower(7, -25), Upper(7, 25) })
    });

    /// <summary>Standard table B.<paramref name="number"/>.</summary>
    private static HuffmanTable Standard(int number) => StandardTables.Value[number - 1];

    // Tables segment (7.4.13) decoded with the procedure of B.2.
    private static HuffmanTable ReadCustomTable(Segment segment, Action<long> charge) {
        int position = segment.DataStart;
        if (segment.DataEnd - position < 9) throw new InvalidDataException("Truncated JBIG2 Huffman table segment.");
        int flags = segment.Data[position];
        bool outOfBand = (flags & 1) != 0;
        int prefixBits = ((flags >> 1) & 7) + 1;
        int rangeBits = ((flags >> 4) & 7) + 1;
        int low = (int)((uint)segment.Data[position + 1] << 24 | (uint)segment.Data[position + 2] << 16 | (uint)segment.Data[position + 3] << 8 | segment.Data[position + 4]);
        int high = (int)((uint)segment.Data[position + 5] << 24 | (uint)segment.Data[position + 6] << 16 | (uint)segment.Data[position + 7] << 8 | segment.Data[position + 8]);
        if (high < low) throw new InvalidDataException("JBIG2 Huffman table range is invalid.");
        var reader = new HuffmanReader(segment.Data, position + 9, segment.DataEnd, charge);
        var lines = new List<HuffmanLine>();
        long current = low;
        while (current < high) {
            if (lines.Count > 65536) throw new InvalidDataException("JBIG2 Huffman table has too many lines.");
            int prefix = (int)reader.ReadBits(prefixBits);
            int range = (int)reader.ReadBits(rangeBits);
            if (range > 31) throw new InvalidDataException("JBIG2 Huffman table range length is invalid.");
            lines.Add(L(prefix, range, (int)current));
            current += 1L << range;
        }
        lines.Add(Lower((int)reader.ReadBits(prefixBits), low - 1));
        lines.Add(Upper((int)reader.ReadBits(prefixBits), high));
        if (outOfBand) lines.Add(Oob((int)reader.ReadBits(prefixBits)));
        charge(lines.Count);
        return new HuffmanTable(lines.ToArray());
    }

    /// <summary>Selects standard tables, or the next referred custom table for a custom selection value.</summary>
    private sealed class CustomTableCursor {
        private readonly List<HuffmanTable> _tables;
        private int _next;

        internal CustomTableCursor(List<HuffmanTable> tables) { _tables = tables; }

        internal HuffmanTable Select(int selection, int customValue, params int[] standards) {
            if (selection == customValue) {
                if (_next >= _tables.Count) throw new InvalidDataException("JBIG2 segment lacks a referred custom Huffman table.");
                return _tables[_next++];
            }
            if (selection >= standards.Length) throw new InvalidDataException("JBIG2 Huffman table selection is invalid.");
            return Standard(standards[selection]);
        }
    }
}
