using System;
using System.Collections.Generic;

namespace OfficeIMO.Markdown;

public sealed class TableBuilder {
    private readonly TableBlock _table = new TableBlock();
    public TableBuilder Headers(params string[] headers) { _table.Headers.AddRange(headers ?? Array.Empty<string>()); return this; }
    public TableBuilder Row(params string[] cells) { _table.Rows.Add(cells?.ToArray() ?? Array.Empty<string>()); return this; }
    public TableBuilder Rows(IEnumerable<IReadOnlyList<string>> rows) { foreach (IReadOnlyList<string> r in rows) _table.Rows.Add(r); return this; }
    public TableBuilder Rows(IEnumerable<(string, string)> rows) { foreach ((string a, string b) in rows) _table.Rows.Add(new[] { a, b }); return this; }
    public TableBuilder Rows(IEnumerable<KeyValuePair<string, string>> rows) { foreach (KeyValuePair<string, string> kv in rows) _table.Rows.Add(new[] { kv.Key, kv.Value }); return this; }
    internal TableBlock Build() => _table;
}

