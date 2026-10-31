namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience wrapper for a table (object-model style).
/// </summary>
public sealed class MarkdownTable : IMarkdownBlock {
    private readonly TableBlock _table = new TableBlock();
    /// <summary>Creates a table with header cells.</summary>
    public MarkdownTable(IEnumerable<string> headers) { _table.Headers.AddRange(headers); }
    /// <summary>Adds a data row and returns its stored cells, which can be edited before rendering.</summary>
    public string[] AddRow(params string[] cells) {
        _table.Rows.Add(cells);
        return cells;
    }
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => ((IMarkdownBlock)_table).RenderMarkdown();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => ((IMarkdownBlock)_table).RenderHtml();
}
