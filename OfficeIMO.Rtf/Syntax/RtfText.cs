namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Represents literal RTF text.
/// </summary>
public sealed class RtfText : RtfNode {
    internal RtfText(int position, string text, string rawText)
        : base(position) {
        Text = text ?? string.Empty;
        RawText = rawText ?? string.Empty;
    }

    /// <summary>Literal text.</summary>
    public string Text { get; }

    /// <summary>Raw source text.</summary>
    public string RawText { get; }
}
