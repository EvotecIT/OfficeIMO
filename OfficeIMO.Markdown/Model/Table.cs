using System.Collections.Generic;

namespace OfficeIMO.Markdown;

public sealed class Table : IMarkdownBlock {
    private readonly TableBlock _table = new TableBlock();
    public Table(IEnumerable<string> headers) { _table.Headers.AddRange(headers); }
    public void AddRow(params string[] cells) => _table.Rows.Add(cells);
    public string RenderMarkdown() => _table.RenderMarkdown();
    public string RenderHtml() => _table.RenderHtml();
}

