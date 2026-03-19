using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Typed table cell containing one or more markdown blocks.
/// </summary>
public sealed class TableCell {
    /// <summary>Structured cell content.</summary>
    public List<IMarkdownBlock> Blocks { get; } = new List<IMarkdownBlock>();

    /// <summary>Creates a typed table cell.</summary>
    public TableCell(IEnumerable<IMarkdownBlock>? blocks = null) {
        if (blocks == null) {
            return;
        }

        foreach (var block in blocks) {
            if (block != null) {
                Blocks.Add(block);
            }
        }
    }

    /// <summary>Markdown representation of the full cell body.</summary>
    public string Markdown => RenderMarkdown();

    internal string RenderMarkdown() {
        if (Blocks.Count == 0) {
            return string.Empty;
        }

        if (Blocks.Count == 1 && Blocks[0] is ParagraphBlock paragraph) {
            return paragraph.Inlines.RenderMarkdown();
        }

        var sb = new StringBuilder();
        for (int i = 0; i < Blocks.Count; i++) {
            var rendered = Blocks[i].RenderMarkdown();
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
        if (Blocks.Count == 0) {
            return string.Empty;
        }

        if (Blocks.Count == 1 && Blocks[0] is ParagraphBlock paragraph) {
            var rendered = paragraph.Inlines.RenderHtml();
            return rendered.Contains('\n') ? rendered.Replace("\n", "<br/>") : rendered;
        }

        var sb = new StringBuilder();
        for (int i = 0; i < Blocks.Count; i++) {
            sb.Append(Blocks[i].RenderHtml());
        }

        return sb.ToString();
    }
}
