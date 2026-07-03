using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Structured definition body inside a definition-list group.
/// </summary>
public sealed class DefinitionListDefinition : MarkdownObject {
    private readonly List<IMarkdownBlock> _blocks = new List<IMarkdownBlock>();
    private readonly List<MarkdownSourceSpan> _blankLineSourceSpans = new List<MarkdownSourceSpan>();
    private readonly List<MarkdownSourceSpan> _continuationIndentSourceSpans = new List<MarkdownSourceSpan>();

    /// <summary>Structured markdown blocks that belong to this definition body.</summary>
    public List<IMarkdownBlock> Blocks => _blocks;
    /// <summary>Source spans for blank separator lines that belong to this definition body.</summary>
    public IReadOnlyList<MarkdownSourceSpan> BlankLineSourceSpans => _blankLineSourceSpans;
    /// <summary>Source spans for indentation stripped from definition continuation lines.</summary>
    public IReadOnlyList<MarkdownSourceSpan> ContinuationIndentSourceSpans => _continuationIndentSourceSpans;
    internal bool ForceParagraphHtml { get; set; }
    internal bool HasLeadingBlankLineBeforeBody { get; set; }

    /// <summary>Creates a definition body.</summary>
    public DefinitionListDefinition(IEnumerable<IMarkdownBlock>? blocks = null) {
        if (blocks == null) {
            return;
        }

        foreach (var block in blocks) {
            if (block != null) {
                _blocks.Add(block);
            }
        }
    }

    /// <summary>Markdown representation of the full definition body.</summary>
    public string Markdown => RenderMarkdown();

    internal string RenderMarkdown() {
        if (_blocks.Count == 0) {
            return string.Empty;
        }

        var sb = new StringBuilder();
        for (int i = 0; i < _blocks.Count; i++) {
            var rendered = MarkdownBlockRenderDispatcher.RenderMarkdown(_blocks[i]);
            if (string.IsNullOrEmpty(rendered)) {
                continue;
            }

            if (sb.Length > 0) {
                sb.Append("\n\n");
            }

            sb.Append(rendered);
        }

        return sb.ToString();
    }

    internal string RenderHtml() {
        if (_blocks.Count == 0) {
            return string.Empty;
        }

        if (_blocks.Count == 1 && _blocks[0] is ParagraphBlock paragraph && !ForceParagraphHtml && paragraph.Attributes.IsEmpty) {
            return paragraph.Inlines.RenderHtml() + RenderConsumedGenericAttributeWhitespace(paragraph);
        }

        var sb = new StringBuilder();
        for (int i = 0; i < _blocks.Count; i++) {
            var rendered = MarkdownBlockRenderDispatcher.RenderHtml(_blocks[i]);
            sb.Append(rendered);
            if (_blocks[i] is HtmlRawBlock && rendered.Length > 0 && !EndsWithLineBreak(rendered)) {
                sb.Append('\n');
            }
        }

        return sb.ToString();
    }

    internal void ReplaceBlankLineSourceSpans(IEnumerable<MarkdownSourceSpan>? spans) {
        _blankLineSourceSpans.Clear();
        if (spans == null) {
            return;
        }

        _blankLineSourceSpans.AddRange(spans);
    }

    internal void ReplaceContinuationIndentSourceSpans(IEnumerable<MarkdownSourceSpan>? spans) {
        _continuationIndentSourceSpans.Clear();
        if (spans == null) {
            return;
        }

        _continuationIndentSourceSpans.AddRange(spans);
    }

    private static bool EndsWithLineBreak(string value) =>
        !string.IsNullOrEmpty(value) &&
        (value[value.Length - 1] == '\n' || value[value.Length - 1] == '\r');

    private static string RenderConsumedGenericAttributeWhitespace(ParagraphBlock paragraph) {
        if (paragraph == null ||
            paragraph.Attributes.IsEmpty ||
            string.IsNullOrEmpty(paragraph.GenericAttributeConsumedWhitespace)) {
            return string.Empty;
        }

        return HtmlTextEncoder.Encode(paragraph.GenericAttributeConsumedWhitespace, HtmlRenderContext.Options);
    }
}
