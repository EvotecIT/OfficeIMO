using System.Collections.Generic;

namespace OfficeIMO.Markdown;

/// <summary>
/// Convenience wrapper for a table (object-model style).
/// </summary>
public sealed class Table : IMarkdownBlock {
    private readonly TableBlock _table = new TableBlock();
    /// <summary>Creates a table with header cells.</summary>
    public Table(IEnumerable<string> headers) { _table.Headers.AddRange(headers); }
    /// <summary>Adds a data row.</summary>
    public void AddRow(params string[] cells) => _table.Rows.Add(cells);
    /// <inheritdoc />
    public string RenderMarkdown() => _table.RenderMarkdown();
    /// <inheritdoc />
    public string RenderHtml() => _table.RenderHtml();
}
