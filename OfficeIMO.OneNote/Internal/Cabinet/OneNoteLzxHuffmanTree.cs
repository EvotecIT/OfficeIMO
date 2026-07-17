namespace OfficeIMO.OneNote;

/// <summary>
/// Immutable canonical Huffman decoder reconstructed from LZX path lengths.
/// </summary>
internal sealed class OneNoteLzxHuffmanTree {
    private const int MaximumPathLength = 16;
    private readonly int[] _counts;
    private readonly int[] _firstCodes;
    private readonly int[] _firstSymbolIndexes;
    private readonly int[] _symbols;

    private OneNoteLzxHuffmanTree(int[] counts, int[] firstCodes, int[] firstSymbolIndexes, int[] symbols) {
        _counts = counts;
        _firstCodes = firstCodes;
        _firstSymbolIndexes = firstSymbolIndexes;
        _symbols = symbols;
    }

    internal static OneNoteLzxHuffmanTree Empty { get; } = new OneNoteLzxHuffmanTree(
        new int[MaximumPathLength + 1],
        new int[MaximumPathLength + 1],
        new int[MaximumPathLength + 1],
        Array.Empty<int>());

    internal static OneNoteLzxHuffmanTree Create(byte[] pathLengths, bool allowEmpty, string treeName) {
        if (pathLengths == null) throw new ArgumentNullException(nameof(pathLengths));
        if (treeName == null) throw new ArgumentNullException(nameof(treeName));

        var counts = new int[MaximumPathLength + 1];
        for (int symbol = 0; symbol < pathLengths.Length; symbol++) {
            int length = pathLengths[symbol];
            if (length < 0 || length > MaximumPathLength) {
                throw Corrupt(treeName + " tree contains an invalid path length.");
            }
            if (length != 0) counts[length]++;
        }

        int symbolCount = counts.Sum();
        if (symbolCount == 0) {
            if (allowEmpty) return Empty;
            throw Corrupt(treeName + " tree is empty.");
        }
        if (symbolCount == 1) {
            throw Corrupt(treeName + " tree contains only one symbol.");
        }

        var firstCodes = new int[MaximumPathLength + 1];
        var firstSymbolIndexes = new int[MaximumPathLength + 1];
        int code = 0;
        int symbolIndex = 0;
        for (int length = 1; length <= MaximumPathLength; length++) {
            code = checked((code + counts[length - 1]) << 1);
            firstCodes[length] = code;
            firstSymbolIndexes[length] = symbolIndex;
            if (code + counts[length] > 1 << length) {
                throw Corrupt(treeName + " tree is oversubscribed.");
            }
            symbolIndex += counts[length];
        }

        var symbols = new int[symbolCount];
        var nextIndexes = (int[])firstSymbolIndexes.Clone();
        for (int symbol = 0; symbol < pathLengths.Length; symbol++) {
            int length = pathLengths[symbol];
            if (length != 0) symbols[nextIndexes[length]++] = symbol;
        }
        return new OneNoteLzxHuffmanTree(counts, firstCodes, firstSymbolIndexes, symbols);
    }

    internal int Decode(OneNoteLzxBitReader reader) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        if (_symbols.Length == 0) throw Corrupt("The LZX stream uses an empty Huffman tree.");

        int code = 0;
        for (int length = 1; length <= MaximumPathLength; length++) {
            code = checked((code << 1) | (int)reader.ReadBits(1));
            int offset = code - _firstCodes[length];
            if ((uint)offset < (uint)_counts[length]) {
                return _symbols[_firstSymbolIndexes[length] + offset];
            }
        }
        throw Corrupt("The LZX stream contains a Huffman code that is absent from its tree.");
    }

    private static OneNoteFormatException Corrupt(string message) =>
        new OneNoteFormatException("ONENOTE_CAB_LZX_CORRUPT", message);
}
