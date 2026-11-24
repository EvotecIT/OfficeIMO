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

    /// <summary>
    /// Canonical VP8L prefix code decoded through a two-level table. The root level resolves codes of up to
    /// eight bits in one lookup; longer codes jump to a second-level table sized for the longest code that shares
    /// the root prefix. Each entry packs the symbol (or second-level offset) above an eight-bit length field.
    /// </summary>
    private sealed class Vp8lHuffmanTree {
        internal const int RootBits = 8;
        private const int RootMask = (1 << RootBits) - 1;
        internal static readonly Vp8lHuffmanTree Invalid = new Vp8lHuffmanTree(Array.Empty<int>(), false);
        private readonly int[] _table;

        private Vp8lHuffmanTree(int[] table, bool isSingleSymbol) {
            _table = table;
            IsSingleSymbol = isSingleSymbol;
        }

        /// <summary>Gets the packed lookup entries; the first <see cref="RootBits"/>-bit block is the root level.</summary>
        internal int[] Table => _table;

        /// <summary>Gets whether every lookup returns one symbol without consuming bits.</summary>
        internal bool IsSingleSymbol { get; }

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
                // A lone symbol is coded with zero bits even though its declared length is one.
                if (lengths[single] != 1) return false;
                if (!allocationBudget.TryReserveArray(1 << RootBits, sizeof(int)) ||
                    !allocationBudget.TryReserveBytes(64L)) return false;
                var singleTable = new int[1 << RootBits];
#if NET8_0_OR_GREATER
                Array.Fill(singleTable, single << 8);
#else
                for (int index = 0; index < singleTable.Length; index++) singleTable[index] = single << 8;
#endif
                tree = new Vp8lHuffmanTree(singleTable, true);
                return true;
            }

            int remaining = 1;
            for (int length = 1; length <= 15; length++) {
                remaining = (remaining << 1) - counts[length];
                if (remaining < 0) return false;
            }
            if (remaining != 0) return false;

            // Canonical codes, assigned in (length, symbol) order, are read least-significant bit first.
            var nextCode = new int[16];
            int code = 0;
            for (int length = 1; length <= 15; length++) {
                code = (code + counts[length - 1]) << 1;
                nextCode[length] = code;
            }

            // Size each second-level table from the longest code sharing its root prefix.
            int[]? subBits = null;
            int longCodes = 0;
            for (int length = RootBits + 1; length <= 15; length++) longCodes += counts[length];
            int tableSize = 1 << RootBits;
            int[]? subOffsets = null;
            var reversedCodes = new int[lengths.Length];
            for (int symbol = 0; symbol < lengths.Length; symbol++) {
                int length = lengths[symbol];
                if (length != 0) reversedCodes[symbol] = (int)ReverseBits((uint)nextCode[length]++, length);
            }
            if (longCodes != 0) {
                subBits = new int[1 << RootBits];
                for (int symbol = 0; symbol < lengths.Length; symbol++) {
                    int length = lengths[symbol];
                    if (length <= RootBits) continue;
                    int root = reversedCodes[symbol] & RootMask;
                    if (length - RootBits > subBits[root]) subBits[root] = length - RootBits;
                }
                subOffsets = new int[1 << RootBits];
                for (int root = 0; root < subBits.Length; root++) {
                    if (subBits[root] == 0) continue;
                    subOffsets[root] = tableSize;
                    tableSize += 1 << subBits[root];
                }
            }

            if (!allocationBudget.TryReserveArray(tableSize, sizeof(int)) ||
                !allocationBudget.TryReserveBytes(64L)) return false;
            var table = new int[tableSize];
            for (int symbol = 0; symbol < lengths.Length; symbol++) {
                int length = lengths[symbol];
                if (length == 0) continue;
                int reversed = reversedCodes[symbol];
                if (length <= RootBits) {
                    int packed = (symbol << 8) | length;
                    for (int index = reversed; index < 1 << RootBits; index += 1 << length) table[index] = packed;
                    continue;
                }
                int root = reversed & RootMask;
                int bits = subBits![root];
                int offset = subOffsets![root];
                int subLength = length - RootBits;
                int subPacked = (symbol << 8) | subLength;
                for (int index = reversed >> RootBits; index < 1 << bits; index += 1 << subLength) {
                    table[offset + index] = subPacked;
                }
                table[root] = (offset << 8) | (RootBits + bits);
            }
            tree = new Vp8lHuffmanTree(table, false);
            return true;
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        internal int ReadSymbol(LsbBitReader reader) => ReadSymbol(_table, reader);

        /// <summary>
        /// Decodes one symbol. Past the end of the payload the reader supplies zero bits and records the overrun,
        /// which callers check before accepting the decoded image.
        /// </summary>
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        internal static int ReadSymbol(int[] table, LsbBitReader reader) {
            if (table.Length == 0) return -1;
            ulong bits = reader.PrefetchBits();
            int entry = table[(int)bits & RootMask];
            int length = entry & 0xFF;
            if (length > RootBits) {
                reader.SkipBits(RootBits);
                entry = table[(entry >> 8) + (int)((bits >> RootBits) & ((1UL << (length - RootBits)) - 1UL))];
                length = entry & 0xFF;
            }
            reader.SkipBits(length);
            return entry >> 8;
        }

        /// <summary>
        /// Looks up the code at the bottom of <paramref name="bits"/> and returns the symbol shifted
        /// left by eight above the total code length. The caller must have buffered 15 bits while
        /// input remains.
        /// </summary>
        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        internal static int Lookup(int[] table, ulong bits) {
#if NET8_0_OR_GREATER
            // Tables built here hold a full root level, and every second-level offset plus its
            // masked index stays inside the table, so both lookups are in range.
            ref int origin = ref System.Runtime.InteropServices.MemoryMarshal.GetArrayDataReference(table);
            int entry = System.Runtime.CompilerServices.Unsafe.Add(ref origin, (int)bits & RootMask);
#else
            int entry = table[(int)bits & RootMask];
#endif
            int length = entry & 0xFF;
            if (length <= RootBits) return entry;
            int index = (entry >> 8) + (int)((bits >> RootBits) & ((1UL << (length - RootBits)) - 1UL));
#if NET8_0_OR_GREATER
            entry = System.Runtime.CompilerServices.Unsafe.Add(ref origin, index);
#else
            entry = table[index];
#endif
            return entry + RootBits;
        }
    }
}
