using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Typed definition list entry containing a term and one or more definition blocks.
/// </summary>
public sealed class DefinitionListEntry {
    /// <summary>Inline content for the definition term.</summary>
    public InlineSequence Term { get; set; }

    /// <summary>Structured definition content for this term.</summary>
    public List<IMarkdownBlock> DefinitionBlocks { get; } = new List<IMarkdownBlock>();

    /// <summary>Creates a typed definition list entry.</summary>
    public DefinitionListEntry(InlineSequence? term = null, IEnumerable<IMarkdownBlock>? definitionBlocks = null) {
        Term = term ?? new InlineSequence();
        if (definitionBlocks == null) {
            return;
        }

        foreach (var block in definitionBlocks) {
            if (block != null) {
                DefinitionBlocks.Add(block);
            }
        }
    }

    /// <summary>Markdown representation of the term.</summary>
    public string TermMarkdown => Term.RenderMarkdown();

    /// <summary>Markdown representation of the full definition body.</summary>
    public string DefinitionMarkdown => RenderDefinitionMarkdown();

    internal string RenderDefinitionMarkdown() {
        if (DefinitionBlocks.Count == 0) {
            return string.Empty;
        }

        var sb = new StringBuilder();
        for (int i = 0; i < DefinitionBlocks.Count; i++) {
            var rendered = DefinitionBlocks[i].RenderMarkdown();
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

    internal string RenderDefinitionHtml() {
        if (DefinitionBlocks.Count == 0) {
            return string.Empty;
        }

        if (DefinitionBlocks.Count == 1 && DefinitionBlocks[0] is ParagraphBlock paragraph) {
            return paragraph.Inlines.RenderHtml();
        }

        var sb = new StringBuilder();
        for (int i = 0; i < DefinitionBlocks.Count; i++) {
            sb.Append(DefinitionBlocks[i].RenderHtml());
        }
        return sb.ToString();
    }
}
