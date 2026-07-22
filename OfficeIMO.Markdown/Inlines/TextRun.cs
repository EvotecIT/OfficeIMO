namespace OfficeIMO.Markdown;

/// <summary>
/// Plain text run.
/// </summary>
public sealed class TextRun : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, ILiteralTextMarkdownInline {
    /// <summary>Text content.</summary>
    public string Text { get; }
    /// <summary>Backslash escape marker text when this run was parsed from an escaped character.</summary>
    public string? EscapeMarker { get; internal set; }
    /// <summary>Source span for the backslash escape marker when this run was parsed from markdown.</summary>
    public MarkdownSourceSpan? EscapeMarkerSourceSpan { get; internal set; }
    /// <summary>Escaped character text when this run was parsed from an escaped character.</summary>
    public string? EscapedCharacter { get; internal set; }
    /// <summary>Source span for the escaped character when this run was parsed from markdown.</summary>
    public MarkdownSourceSpan? EscapedCharacterSourceSpan { get; internal set; }
    /// <summary>Creates a plain text run.</summary>
    public TextRun(string text) { Text = text ?? string.Empty; }

    internal void SetMarkdownSyntaxMetadataSpans(
        string? escapeMarker,
        MarkdownSourceSpan? escapeMarkerSourceSpan,
        string? escapedCharacter,
        MarkdownSourceSpan? escapedCharacterSourceSpan) {
        EscapeMarker = escapeMarker;
        EscapeMarkerSourceSpan = escapeMarkerSourceSpan;
        EscapedCharacter = escapedCharacter;
        EscapedCharacterSourceSpan = escapedCharacterSourceSpan;
    }

    internal string RenderMarkdown() => MarkdownEscaper.EscapeText(Text);
    internal string RenderHtml() => HtmlTextEncoder.Encode(Text, HtmlRenderContext.Options);
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Text);
}
