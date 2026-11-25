namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private const int HuffmanFastBits = 9;

    private static void DecodeBlockCoefficients(
        ref JpegBitReader reader,
        in HuffmanTable dcTable,
        in HuffmanTable acTable,
        int[] quant,
        ref int prevDc,
        int[] coeffs) {
        Array.Clear(coeffs, 0, 64);

        var t = DecodeHuffman(ref reader, dcTable, useFast: true);
        var diff = t == 0 ? 0 : Extend(reader.ReadBits(t), t);
        var dc = checked(prevDc + diff);
        prevDc = dc;
        coeffs[0] = checked(dc * quant[0]);

        var k = 1;
        while (k < 64) {
            if (acTable.FastAc is not null && reader.TryPeekBits(HuffmanFastBits, out int peek)) {
                int entry = acTable.FastAc[peek];
                if (entry != 0) {
                    int coefficient = entry >> 8;
                    int run = (entry >> 4) & 15;
                    if (coefficient == 0) {
                        reader.SkipBits(entry & 15);
                        if (run == 15) {
                            k += 16;
                            continue;
                        }
                        break;
                    }
                    // An out-of-range run consumes its Huffman code only.
                    // Retain the existing decoder for that malformed case.
                    if (k + run < 64) {
                        reader.SkipBits(entry & 15);
                        k += run;
                        int index = ZigZag[k++];
                        coeffs[index] = checked(coefficient * quant[index]);
                        continue;
                    }
                }
            }
            var rs = DecodeHuffman(ref reader, acTable, useFast: true);
            if (rs == 0) break;
            var r = rs >> 4;
            var s = rs & 0x0F;
            if (s == 0) {
                if (r == 15) {
                    k += 16;
                    continue;
                }
                break;
            }

            k += r;
            if (k >= 64) break;
            var ac = Extend(reader.ReadBits(s), s);
            var zig = ZigZag[k];
            coeffs[zig] = checked(ac * quant[zig]);
            k++;
        }

    }

    private static int DecodeHuffman(ref JpegBitReader reader, in HuffmanTable table, bool useFast) {
        if (useFast && table.Fast is not null && reader.TryPeekBits(HuffmanFastBits, out int peek)) {
            var entry = table.Fast[peek];
            if (entry >= 0) {
                var size = entry >> 8;
                reader.SkipBits(size);
                return entry & 0xFF;
            }
        }

        var node = 0;
        while (true) {
            var bit = reader.ReadBit();
            node = bit == 0 ? table.Left[node] : table.Right[node];
            if (node < 0) {
                if (reader.AllowTruncated) return 0;
                throw new FormatException("Invalid JPEG Huffman code.");
            }
            var symbol = table.Symbols[node];
            if (symbol >= 0) return symbol;
        }
    }

    private static int Extend(int value, int bits) {
        if (bits == 0) return 0;
        var limit = 1 << (bits - 1);
        if (value < limit) value -= (1 << bits) - 1;
        return value;
    }

    private struct HuffmanTable {
        public int[] Left;
        public int[] Right;
        public int[] Symbols;
        public short[]? Fast;
        public short[]? FastAc;
        public bool IsValid;

        public void PrepareBaselineAcLookup(ref long reservedBytes) {
            if (FastAc is not null || Fast is null) return;
            var lookup = new short[OfficeRasterGuards.EnsureInt16ArrayLength(
                1 << HuffmanFastBits, ref reservedBytes, JpegDimensionsLimitMessage)];
            for (int peek = 0; peek < lookup.Length; peek++) {
                int entry = Fast[peek];
                if (entry < 0) continue;
                int codeBits = entry >> 8;
                int symbol = entry & 255;
                int valueBits = symbol & 15;
                int totalBits = codeBits + valueBits;
                // Signed five-bit magnitudes fit in the upper byte; the low
                // byte packs the zero run and total consumed bit count.
                if (valueBits > 5 || totalBits > HuffmanFastBits) continue;
                int value = (peek >> (HuffmanFastBits - totalBits)) & ((1 << valueBits) - 1);
                int coefficient = Extend(value, valueBits);
                lookup[peek] = (short)((coefficient << 8) | ((symbol >> 4) << 4) | totalBits);
            }
            FastAc = lookup;
        }

        public static HuffmanTable Build(OfficeByteView counts, byte[] values) {
            var left = new int[512];
            var right = new int[512];
            var symbols = new int[512];
            var fast = new short[1 << HuffmanFastBits];
            for (var i = 0; i < left.Length; i++) left[i] = -1;
            for (var i = 0; i < right.Length; i++) right[i] = -1;
            for (var i = 0; i < symbols.Length; i++) symbols[i] = -1;
            for (var i = 0; i < fast.Length; i++) fast[i] = -1;

            var next = 1;
            var code = 0;
            var k = 0;
            for (var i = 1; i <= 16; i++) {
                var count = counts[i - 1];
                for (var j = 0; j < count; j++) {
                    var symbol = values[k++];
                    var node = 0;
                    for (var bit = i - 1; bit >= 0; bit--) {
                        var b = (code >> bit) & 1;
                        if (b == 0) {
                            if (left[node] < 0) {
                                if (next >= left.Length) throw new FormatException("Invalid JPEG Huffman tree.");
                                left[node] = next++;
                            }
                            node = left[node];
                        } else {
                            if (right[node] < 0) {
                                if (next >= right.Length) throw new FormatException("Invalid JPEG Huffman tree.");
                                right[node] = next++;
                            }
                            node = right[node];
                        }
                    }
                    symbols[node] = symbol;
                    if (i <= HuffmanFastBits) {
                        var fill = 1 << (HuffmanFastBits - i);
                        var start = code << (HuffmanFastBits - i);
                        var entry = (short)((i << 8) | symbol);
                        for (var f = 0; f < fill; f++) {
                            fast[start + f] = entry;
                        }
                    }
                    code++;
                }
                code <<= 1;
            }

            return new HuffmanTable {
                Left = left,
                Right = right,
                Symbols = symbols,
                Fast = fast,
                IsValid = true
            };
        }
    }

}
