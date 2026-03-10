namespace OfficeIMO.Markdown;

/// <summary>
/// Definition list rendered as term/definition pairs. Markdown output uses
/// a simple "Term: Definition" fallback; HTML uses &lt;dl&gt;.
/// </summary>
public sealed class DefinitionListBlock : IMarkdownBlock {
    /// <summary>List of (term, definition) pairs.</summary>
    public List<(string Term, string Definition)> Items { get; } = new List<(string, string)>();
    internal List<(InlineSequence Term, InlineSequence Definition)>? ParsedItems { get; private set; }
    internal int? ParsedContentSignature { get; private set; }
    internal List<MarkdownSyntaxNode> SyntaxItems { get; } = new List<MarkdownSyntaxNode>();
    internal MarkdownReaderOptions? ReaderOptions { get; private set; }
    internal MarkdownReaderState? ReaderState { get; private set; }

    internal void SetParsingContext(MarkdownReaderOptions options, MarkdownReaderState state) {
        ReaderOptions = options;
        ReaderState = state;
    }

    internal void SetParsedItems(IReadOnlyList<(InlineSequence Term, InlineSequence Definition)>? parsedItems, int contentSignature) {
        if (parsedItems == null) {
            ParsedItems = null;
        } else {
            ParsedItems = new List<(InlineSequence Term, InlineSequence Definition)>(parsedItems.Count);
            for (int i = 0; i < parsedItems.Count; i++) {
                ParsedItems.Add(parsedItems[i]);
            }
        }

        ParsedContentSignature = contentSignature;
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
        var options = ReaderOptions ?? new MarkdownReaderOptions();
        var state = ReaderState ?? new MarkdownReaderState();
        bool useParsedItems = ParsedContentSignature.HasValue && ParsedContentSignature.Value == ComputeContentSignature();
        for (int index = 0; index < Items.Count; index++) {
            var (term, def) = Items[index];
            var parsed = useParsedItems && ParsedItems != null && index < ParsedItems.Count
                ? ParsedItems[index]
                : default;
            sb.Append("<dt>");
            if (useParsedItems && ParsedItems != null && index < ParsedItems.Count) {
                sb.Append(parsed.Term.RenderHtml());
            } else if (!string.IsNullOrEmpty(term)) {
                sb.Append(MarkdownReader.ParseInlineText(term, options, state).RenderHtml());
            }
            sb.Append("</dt>");
            sb.Append("<dd>");
            if (useParsedItems && ParsedItems != null && index < ParsedItems.Count) {
                sb.Append(parsed.Definition.RenderHtml());
            } else if (!string.IsNullOrEmpty(def)) {
                sb.Append(MarkdownReader.ParseInlineText(def, options, state).RenderHtml());
            }
            sb.Append("</dd>");
        }
        sb.Append("</dl>");
        return sb.ToString();
    }

    internal int ComputeContentSignature() {
        unchecked {
            int hash = 17;
            hash = (hash * 31) + Items.Count;
            for (int i = 0; i < Items.Count; i++) {
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Items[i].Term ?? string.Empty);
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Items[i].Definition ?? string.Empty);
            }
            return hash;
        }
    }
}
