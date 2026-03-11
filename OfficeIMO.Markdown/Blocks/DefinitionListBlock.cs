namespace OfficeIMO.Markdown;

/// <summary>
/// Definition list rendered as term/definition pairs. Markdown output uses
/// a simple "Term: Definition" fallback; HTML uses &lt;dl&gt;.
/// </summary>
public sealed class DefinitionListBlock : IMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>List of (term, definition) pairs.</summary>
    public List<(string Term, string Definition)> Items { get; } = new List<(string, string)>();
    /// <summary>Parsed inline representation of the current definition list items.</summary>
    public IReadOnlyList<DefinitionListInlineItem> InlineItems => BuildInlineItems();
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
        var inlineItems = BuildInlineItems();
        for (int index = 0; index < Items.Count; index++) {
            var parsed = inlineItems[index];
            sb.Append("<dt>");
            sb.Append(parsed.Term.RenderHtml());
            sb.Append("</dt>");
            sb.Append("<dd>");
            sb.Append(parsed.Definition.RenderHtml());
            sb.Append("</dd>");
        }
        sb.Append("</dl>");
        return sb.ToString();
    }

    private IReadOnlyList<DefinitionListInlineItem> BuildInlineItems() {
        if (Items.Count == 0) {
            return Array.Empty<DefinitionListInlineItem>();
        }

        var options = ReaderOptions ?? new MarkdownReaderOptions();
        var state = ReaderState ?? new MarkdownReaderState();
        bool useParsedItems = ParsedContentSignature.HasValue && ParsedContentSignature.Value == ComputeContentSignature();
        var inlineItems = new DefinitionListInlineItem[Items.Count];

        for (int index = 0; index < Items.Count; index++) {
            if (useParsedItems && ParsedItems != null && index < ParsedItems.Count) {
                var parsed = ParsedItems[index];
                inlineItems[index] = new DefinitionListInlineItem(parsed.Term, parsed.Definition);
                continue;
            }

            var (term, def) = Items[index];
            inlineItems[index] = new DefinitionListInlineItem(
                string.IsNullOrEmpty(term) ? new InlineSequence() : MarkdownReader.ParseInlineText(term, options, state),
                string.IsNullOrEmpty(def) ? new InlineSequence() : MarkdownReader.ParseInlineText(def, options, state));
        }

        return inlineItems;
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

    internal IReadOnlyList<MarkdownSyntaxNode> BuildSyntaxItems() {
        if (SyntaxItems.Count > 0) {
            return SyntaxItems;
        }

        var nodes = new List<MarkdownSyntaxNode>();
        foreach (var (term, definition) in Items) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.DefinitionItem,
                literal: term,
                children: new[] {
                    new MarkdownSyntaxNode(MarkdownSyntaxKind.DefinitionTerm, literal: term),
                    new MarkdownSyntaxNode(
                        MarkdownSyntaxKind.DefinitionValue,
                        literal: definition,
                        children: new[] {
                            new MarkdownSyntaxNode(MarkdownSyntaxKind.Paragraph, literal: definition)
                        })
                }));
        }

        return nodes;
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.DefinitionList, span, children: BuildSyntaxItems());
}
