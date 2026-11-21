namespace OfficeIMO.Markdown;

/// <summary>
/// Hard line break inline. Markdown renders as two spaces + newline; HTML as <br/>.
/// </summary>
public sealed class HardBreakInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IHtmlAppendableMarkdownInline {
    private string? _marker;

    /// <summary>Source marker spelling that produced this hard break when parsed from markdown.</summary>
    public string? Marker { get => ReadSourceInformation(ref _marker); internal set => WriteSourceInformation(ref _marker, value); }
    /// <summary>Source span for the hard-break marker when parsed from markdown.</summary>
    public MarkdownSourceSpan? MarkerSourceSpan { get => ReadSourceInformation(ref field); internal set => WriteSourceInformation(ref field, value); }

    internal void SetMarkdownSyntaxMetadataSpans(string? marker, MarkdownSourceSpan? markerSourceSpan) {
        Marker = marker;
        MarkerSourceSpan = markerSourceSpan;
    }

    internal string RenderMarkdown() => "  \n";
    // An HTML break marker is recorded the same way with and without source capture, so rendering
    // reads it without capturing deferred source information.
    internal string RenderHtml() =>
        _marker is "<br>" or "<br/>" or "<br />" ? _marker : "<br/>";
    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) => builder.Append(RenderHtml());
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(' ');

    internal override void AdoptCapturedSourceInformation(MarkdownObject captured, MarkdownCapturedObjectMap map) {
        var source = (HardBreakInline)captured;
        Marker = source.Marker;
        MarkerSourceSpan = source.MarkerSourceSpan;
    }
}
