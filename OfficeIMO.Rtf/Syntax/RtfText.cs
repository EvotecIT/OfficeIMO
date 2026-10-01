namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Represents literal RTF text.
/// </summary>
public sealed class RtfText : RtfNode {
    internal RtfText(int position, string text, string rawText, int? encodedUnicodeSkipCount = null)
        : base(position) {
        Text = text ?? string.Empty;
        RawText = rawText ?? string.Empty;
        EncodedUnicodeSkipCount = encodedUnicodeSkipCount;
    }

    /// <summary>Literal text.</summary>
    public string Text { get; }

    /// <summary>Raw source text.</summary>
    public string RawText { get; }

    // Newly encoded editor text can use a different fallback count from its insertion scope.
    internal int? EncodedUnicodeSkipCount { get; }
}
