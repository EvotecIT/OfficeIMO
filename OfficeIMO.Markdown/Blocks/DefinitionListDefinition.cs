using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Structured definition body inside a definition-list group.
/// </summary>
public sealed class DefinitionListDefinition : MarkdownObject {
    private readonly List<IMarkdownBlock> _blocks = new List<IMarkdownBlock>();

    /// <summary>Structured markdown blocks that belong to this definition body.</summary>
    public List<IMarkdownBlock> Blocks => _blocks;
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

        if (_blocks.Count == 1 && _blocks[0] is ParagraphBlock paragraph && !ForceParagraphHtml) {
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
