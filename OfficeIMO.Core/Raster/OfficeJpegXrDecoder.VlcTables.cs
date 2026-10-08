namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    // T.832 Tables 80, 82, 87 and 88. Entries are indexed by decoded symbol.
    private static readonly Codebook[] IndexCodes = {
        new("1", "00000", "001", "00001", "01", "0001"),
        new("01", "0000", "10", "0001", "11", "001"),
        new("0000", "0001", "01", "10", "11", "001"),
        new("00000", "00001", "01", "1", "0001", "001")
    };
    private static readonly Codebook[] FirstCodes = {
        new("00001", "000001", "0000000", "0000001", "00100", "010", "00101", "1", "00110", "0001", "00111", "011"),
        new("0010", "00010", "000000", "000001", "0011", "010", "00011", "11", "011", "100", "00001", "101"),
        new("11", "001", "0000000", "0000001", "00001", "010", "0000010", "011", "100", "101", "0000011", "0001"),
        new("001", "11", "0000000", "00001", "00010", "010", "0000001", "011", "00011", "100", "000001", "101"),
        new("010", "1", "0000001", "0001", "0000010", "011", "00000000", "0010", "0000011", "0011", "00000001", "00001")
    };
    private static readonly int[][] IndexDeltas = {
        new[] { -1, 1, 1, 1, 0, 1 }, new[] { -2, 0, 0, 2, 0, 0 }, new[] { -1, -1, 0, 1, -2, 0 }
    };
    private static readonly int[][] FirstDeltas = {
        new[] { 1, 1, 1, 1, 1, 0, 0, -1, 2, 1, 0, 0 },
        new[] { 2, 2, -1, -1, -1, 0, -2, -1, 0, 0, -2, -1 },
        new[] { -1, 1, 0, 2, 0, 0, 0, 0, -2, 0, 1, 1 },
        new[] { 0, 1, 0, 1, -2, 0, -1, -1, -2, -1, -2, -2 }
    };
    private static readonly Codebook FinalIndexCodes = new("0", "110", "10", "111");
    private static readonly Codebook RunCodes = new("1", "01", "001", "0000", "0001");
    private static readonly Codebook LowpassPresence = new("0", "100", "1010", "1011", "1100", "1101", "1110", "1111");
    private static readonly int[] RunBase = { 1, 2, 3, 5, 7, 1, 2, 3, 5, 7, 1, 2, 3, 4, 5 };
    private static readonly int[] RunBin = { -1, -1, -1, -1, 2, 2, 2, 1, 1, 1, 1, 0, 0, 0, 0 };
    private static readonly int[] RunBits = { 0, 0, 1, 1, 3, 0, 0, 1, 1, 2, 0, 0, 0, 0, 1 };
    private static readonly int[] HorizontalScan = { 0, 4, 1, 5, 8, 2, 9, 6, 12, 3, 10, 13, 7, 14, 11, 15 };
    private static readonly int[] VerticalScan = { 0, 1, 2, 5, 4, 3, 6, 9, 8, 7, 12, 15, 13, 10, 11, 14 };
    private static readonly int[] TransposedScan = { 0, 4, 8, 12, 1, 5, 9, 13, 2, 6, 10, 14, 3, 7, 11, 15 };
    private static readonly int[] AbsLevelDeltas = { 1, 0, -1, -1, -1, -1, -1 };
    private static readonly Codebook[] PatternCountCodes = {
        new("1", "01", "001", "0000", "0001"), new("1", "000", "001", "010", "011")
    };
    private static readonly Codebook[] ChromaPatternCountCodes = {
        new("010", "00000", "0010", "00001", "00010", "1", "011", "00011", "0011"),
        new("1", "001", "010", "0001", "000001", "011", "00001", "0000000", "0000001")
    };
    private static readonly int[] PatternCountDeltas = { 0, -1, 0, 1, 1 };
    private static readonly int[] ChromaPatternCountDeltas = { 2, 2, 1, 1, -1, -2, -2, -2, -3 };
    private static readonly Codebook TernaryCodes = new("1", "01", "00");
    private static readonly Codebook ChromaBlockCodes = new("1", "01", "000", "001");
    private static readonly Codebook DoublePatternCodes = new("00", "01", "100", "101", "110", "111");
    private static readonly int[] DoublePatterns = { 3, 5, 6, 9, 10, 12 };
    private static readonly int[] PatternWidths = { 0, 2, 1, 2, 2, 0 }, PatternOffsets = { 0, 4, 2, 8, 12, 1 };
    private static readonly int[] PatternValues = { 0, 15, 3, 12, 1, 2, 4, 8, 5, 6, 9, 10, 7, 11, 13, 14 };
    private static readonly int[] HierarchicalScan = { 0, 1, 4, 5, 2, 3, 6, 7, 8, 9, 12, 13, 10, 11, 14, 15 };
}
