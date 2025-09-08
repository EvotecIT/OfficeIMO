using System;
using System.Collections.Generic;

namespace OfficeIMO.Markdown;

/// <summary>
/// Builder for pipe tables.
/// </summary>
public sealed class TableBuilder {
    private readonly TableBlock _table = new TableBlock();
    /// <summary>Sets the header row.</summary>
    public TableBuilder Headers(params string[] headers) { _table.Headers.AddRange(headers ?? Array.Empty<string>()); return this; }
    /// <summary>Adds a data row.</summary>
    public TableBuilder Row(params string[] cells) { _table.Rows.Add(cells?.ToArray() ?? Array.Empty<string>()); return this; }
    /// <summary>Adds multiple rows.</summary>
    public TableBuilder Rows(IEnumerable<IReadOnlyList<string>> rows) { foreach (IReadOnlyList<string> r in rows) _table.Rows.Add(r); return this; }
    /// <summary>Adds two-column rows from tuples.</summary>
    public TableBuilder Rows(IEnumerable<(string, string)> rows) { foreach ((string a, string b) in rows) _table.Rows.Add(new[] { a, b }); return this; }
    /// <summary>Adds two-column rows from key/value pairs.</summary>
    public TableBuilder Rows(IEnumerable<KeyValuePair<string, string>> rows) { foreach (KeyValuePair<string, string> kv in rows) _table.Rows.Add(new[] { kv.Key, kv.Value }); return this; }
    internal TableBlock Build() => _table;
}
