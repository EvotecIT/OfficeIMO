using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    private static readonly int[] Vp8lCodeLengthOrder = {
        17, 18, 0, 1, 2, 3, 4, 5, 16, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15
    };

    private static bool TryReadHuffmanTree(
        LsbBitReader reader,
        int alphabetSize,
        Vp8lAllocationBudget allocationBudget,
        out Vp8lHuffmanTree tree) {
        tree = Vp8lHuffmanTree.Invalid;
        if (alphabetSize < 1 || !allocationBudget.TryReserveArray(alphabetSize, sizeof(byte))) return false;
        var lengths = new byte[alphabetSize];
        if (reader.ReadBits(1) != 0) {
            int symbolCount = (int)reader.ReadBits(1) + 1;
            int firstBits = reader.ReadBits(1) == 0 ? 1 : 8;
            int first = (int)reader.ReadBits(firstBits);
            if (first >= alphabetSize) return false;
            lengths[first] = 1;
            if (symbolCount == 2) {
                int second = (int)reader.ReadBits(8);
                if (second >= alphabetSize || second == first) return false;
                lengths[second] = 1;
            }
            return Vp8lHuffmanTree.TryCreate(lengths, allocationBudget, out tree);
        }

        int codeLengthCount = 4 + (int)reader.ReadBits(4);
        if (!allocationBudget.TryReserveArray(19, sizeof(byte))) return false;
        var codeLengthLengths = new byte[19];
        for (int index = 0; index < codeLengthCount; index++) {
            codeLengthLengths[Vp8lCodeLengthOrder[index]] = (byte)reader.ReadBits(3);
        }
        if (!Vp8lHuffmanTree.TryCreate(codeLengthLengths, allocationBudget, out Vp8lHuffmanTree codeLengthTree)) return false;

        int maxSymbol = alphabetSize;
        if (reader.ReadBits(1) != 0) {
            int lengthBits = 2 + 2 * (int)reader.ReadBits(3);
            maxSymbol = 2 + (int)reader.ReadBits(lengthBits);
            if (maxSymbol > alphabetSize) return false;
        }

        int position = 0;
        int remainingCodeLengthInstructions = maxSymbol;
        int previous = 8;
        while (position < alphabetSize && remainingCodeLengthInstructions-- > 0) {
            int symbol = codeLengthTree.ReadSymbol(reader);
            if (symbol < 0) return false;
            if (symbol <= 15) {
                lengths[position++] = (byte)symbol;
                if (symbol != 0) previous = symbol;
                continue;
            }
            int repeat;
            int value;
            if (symbol == 16) {
                repeat = 3 + (int)reader.ReadBits(2);
                value = previous;
            } else if (symbol == 17) {
                repeat = 3 + (int)reader.ReadBits(3);
                value = 0;
            } else if (symbol == 18) {
                repeat = 11 + (int)reader.ReadBits(7);
                value = 0;
            } else {
                return false;
            }
            // libwebp treats max_symbol as the encoded instruction count; a repeat may expand
            // beyond it but must remain within the destination alphabet.
            if (repeat > alphabetSize - position) return false;
            for (int index = 0; index < repeat; index++) lengths[position++] = (byte)value;
        }
        return Vp8lHuffmanTree.TryCreate(lengths, allocationBudget, out tree);
    }

    private sealed class Vp8lHuffmanTree {
        private const int LookupBits = 8;
        internal static readonly Vp8lHuffmanTree Invalid = new Vp8lHuffmanTree(Array.Empty<int>(), Array.Empty<int>(), Array.Empty<int>(), -1);
        private readonly int[] _symbols;
        private readonly int[] _counts;
        private readonly int[] _lookup;
        private readonly int _singleSymbol;

        private Vp8lHuffmanTree(int[] symbols, int[] counts, int[] lookup, int singleSymbol) {
            _symbols = symbols;
            _counts = counts;
            _lookup = lookup;
            _singleSymbol = singleSymbol;
        }

        internal static bool TryCreate(
            byte[] lengths,
            Vp8lAllocationBudget allocationBudget,
            out Vp8lHuffmanTree tree) {
            tree = Invalid;
            int used = 0;
            int single = -1;
            if (!allocationBudget.TryReserveArray(16, sizeof(int))) return false;
            var counts = new int[16];
            for (int symbol = 0; symbol < lengths.Length; symbol++) {
                int length = lengths[symbol];
                if (length > 15) return false;
                if (length == 0) continue;
                counts[length]++;
                used++;
                single = symbol;
            }
            if (used == 0) return false;
            if (used == 1) {
                if (lengths[single] != 1) return false;
                if (!allocationBudget.TryReserveBytes(64L)) return false;
                tree = new Vp8lHuffmanTree(Array.Empty<int>(), Array.Empty<int>(), Array.Empty<int>(), single);
                return true;
            }

            int remaining = 1;
            for (int length = 1; length <= 15; length++) {
                remaining = (remaining << 1) - counts[length];
                if (remaining < 0) return false;
            }
            if (remaining != 0) return false;

            if (!allocationBudget.TryReserveArray(used, sizeof(int)) ||
                !allocationBudget.TryReserveArray(1 << LookupBits, sizeof(int)) ||
                !allocationBudget.TryReserveBytes(64L)) return false;
            var symbols = new int[used];
            var lookup = new int[1 << LookupBits];
            int position = 0;
            int firstCode = 0;
            for (int length = 1; length <= 15; length++) {
                int code = firstCode;
                for (int symbol = 0; symbol < lengths.Length; symbol++) {
                    if (lengths[symbol] != length) continue;
                    symbols[position++] = symbol;
                    if (length <= LookupBits) {
                        int reversed = (int)ReverseBits((uint)code, length);
                        int packed = (symbol << 4) | length;
                        for (int index = reversed; index < lookup.Length; index += 1 << length) {
                            lookup[index] = packed;
                        }
                    }
                    code++;
                }
                firstCode = code << 1;
            }
            tree = new Vp8lHuffmanTree(symbols, counts, lookup, -1);
            return true;
        }

        internal int ReadSymbol(LsbBitReader reader) {
            if (_singleSymbol >= 0) return _singleSymbol;
            if (_symbols.Length == 0) return -1;
            // Do not require eight bits at the end of a valid stream: short
            // final codes still use the canonical decoder below.
            if (reader.HasBits(LookupBits)) {
                int packed = _lookup[reader.PeekBits(LookupBits)];
                int length = packed & 15;
                if (length != 0) {
                    reader.ReadBits(length);
                    return packed >> 4;
                }
            }
            int code = 0;
            int firstCode = 0;
            int firstIndex = 0;
            for (int length = 1; length <= 15; length++) {
                code = (code << 1) | (int)reader.ReadBits(1);
                int count = _counts[length];
                int relative = code - firstCode;
                if (relative >= 0 && relative < count) return _symbols[firstIndex + relative];
                firstCode = (firstCode + count) << 1;
                firstIndex += count;
            }
            return -1;
        }
    }
}
