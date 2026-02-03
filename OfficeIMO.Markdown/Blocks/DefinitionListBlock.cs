namespace OfficeIMO.Markdown;

/// <summary>
/// Definition list rendered as term/definition pairs. Markdown output uses
/// a simple "Term: Definition" fallback; HTML uses &lt;dl&gt;.
/// </summary>
public sealed class DefinitionListBlock : IMarkdownBlock {
    /// <summary>List of (term, definition) pairs.</summary>
    public List<(string Term, string Definition)> Items { get; } = new List<(string, string)>();
    internal MarkdownReaderOptions? ReaderOptions { get; private set; }
    internal MarkdownReaderState? ReaderState { get; private set; }

    internal void SetParsingContext(MarkdownReaderOptions options, MarkdownReaderState state) {
        ReaderOptions = options;
        ReaderState = state;
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        StringBuilder sb = new StringBuilder();
        foreach (var (term, def) in Items) sb.AppendLine(term + ": " + def);
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        StringBuilder sb = new StringBuilder();
        sb.Append("<dl>");
        foreach (var (term, def) in Items) {
            sb.Append("<dt>");
            if (!string.IsNullOrEmpty(term)) {
                var termInlines = MarkdownReader.ParseInlineText(term, ReaderOptions, ReaderState);
                sb.Append(termInlines.RenderHtml());
            }
            sb.Append("</dt>");
            sb.Append("<dd>");
            if (!string.IsNullOrEmpty(def)) {
                var inlines = MarkdownReader.ParseInlineText(def, ReaderOptions, ReaderState);
                sb.Append(inlines.RenderHtml());
            }
            sb.Append("</dd>");
        }
        sb.Append("</dl>");
        return sb.ToString();
    }
}
