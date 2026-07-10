namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Original AsciiDoc source and its offset-to-line mapping.
/// </summary>
public sealed class AsciiDocSourceText {
    private readonly int[] _lineStarts;

    internal AsciiDocSourceText(string text) {
        Text = text ?? string.Empty;
        _lineStarts = BuildLineStarts(Text);
        PreferredLineEnding = DetectPreferredLineEnding(Text);
    }

    /// <summary>Original source text.</summary>
    public string Text { get; }

    /// <summary>First line ending used by the document, or <see cref="Environment.NewLine"/> when absent.</summary>
    public string PreferredLineEnding { get; }

    /// <summary>Number of logical lines. Empty source has one logical line.</summary>
    public int LineCount => _lineStarts.Length;

    /// <summary>Creates a half-open span from zero-based offsets.</summary>
    public AsciiDocSourceSpan CreateSpan(int startOffset, int endOffset) {
        if (startOffset < 0 || startOffset > Text.Length) throw new ArgumentOutOfRangeException(nameof(startOffset));
        if (endOffset < startOffset || endOffset > Text.Length) throw new ArgumentOutOfRangeException(nameof(endOffset));
        return new AsciiDocSourceSpan(GetPosition(startOffset), GetPosition(endOffset));
    }

    /// <summary>Maps a zero-based source offset to a one-based line and column.</summary>
    public AsciiDocSourcePosition GetPosition(int offset) {
        if (offset < 0 || offset > Text.Length) throw new ArgumentOutOfRangeException(nameof(offset));

        int lineIndex = FindLineIndex(offset);
        return new AsciiDocSourcePosition(offset, lineIndex + 1, offset - _lineStarts[lineIndex] + 1);
    }

    private int FindLineIndex(int offset) {
        int low = 0;
        int high = _lineStarts.Length - 1;
        while (low <= high) {
            int middle = low + ((high - low) / 2);
            int start = _lineStarts[middle];
            if (start == offset) return middle;
            if (start < offset) low = middle + 1; else high = middle - 1;
        }

        return Math.Max(0, low - 1);
    }

    private static int[] BuildLineStarts(string text) {
        var starts = new List<int> { 0 };
        for (int index = 0; index < text.Length; index++) {
            if (text[index] == '\r') {
                if (index + 1 < text.Length && text[index + 1] == '\n') index++;
                if (index + 1 <= text.Length) starts.Add(index + 1);
            } else if (text[index] == '\n') {
                if (index + 1 <= text.Length) starts.Add(index + 1);
            }
        }

        return starts.ToArray();
    }

    private static string DetectPreferredLineEnding(string text) {
        for (int index = 0; index < text.Length; index++) {
            if (text[index] == '\r') {
                return index + 1 < text.Length && text[index + 1] == '\n' ? "\r\n" : "\r";
            }
            if (text[index] == '\n') return "\n";
        }

        return Environment.NewLine;
    }
}
