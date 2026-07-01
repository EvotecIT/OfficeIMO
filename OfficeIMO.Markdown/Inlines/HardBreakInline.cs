namespace OfficeIMO.Markdown;

/// <summary>
/// Hard line break inline. Markdown renders as two spaces + newline; HTML as <br/>.
/// </summary>
public sealed class HardBreakInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    /// <summary>Source marker spelling that produced this hard break when parsed from markdown.</summary>
    public string? Marker { get; internal set; }
    /// <summary>Source span for the hard-break marker when parsed from markdown.</summary>
    public MarkdownSourceSpan? MarkerSourceSpan { get; internal set; }

    internal void SetMarkdownSyntaxMetadataSpans(string? marker, MarkdownSourceSpan? markerSourceSpan) {
        Marker = marker;
        MarkerSourceSpan = markerSourceSpan;
    }

    internal string RenderMarkdown() => "  \n";
    internal string RenderHtml() =>
        Marker is "<br>" or "<br/>" or "<br />" ? Marker : "<br/>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(' ');
}
