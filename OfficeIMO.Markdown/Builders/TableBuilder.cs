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

    /// <summary>
    /// Populates the table from an arbitrary object. If <paramref name="data"/> is
    /// - a sequence of scalars → one-column table with a row per value,
    /// - a sequence of objects → headers from public readable properties and rows per item,
    /// - a dictionary → two columns Key/Value,
    /// - a POCO → two columns Property/Value.
    /// </summary>
    public TableBuilder FromAny(object? data) {
        if (data is null) return this;
        if (data is string || data.GetType().IsPrimitive) {
            if (_table.Headers.Count == 0) _table.Headers.Add("Value");
            _table.Rows.Add(new[] { data.ToString() ?? string.Empty });
            return this;
        }

        if (data is System.Collections.IDictionary dict) {
            if (_table.Headers.Count == 0) _table.Headers.AddRange(new[] { "Key", "Value" });
            foreach (var key in dict.Keys) {
                var val = dict[key];
                _table.Rows.Add(new[] { key?.ToString() ?? string.Empty, FormatValue(val) });
            }
            return this;
        }

        if (data is System.Collections.IEnumerable seq && data is not string) {
            // Determine if scalar or object sequence by peeking first non-null item
            object? first = null;
            foreach (var item in seq) { first = item; if (first != null) break; }
            if (first is null) return this; // empty sequence

            if (IsScalar(first)) {
                if (_table.Headers.Count == 0) _table.Headers.Add("Value");
                foreach (var item in seq) _table.Rows.Add(new[] { FormatValue(item) });
                return this;
            }

            // Object sequence → headers from public readable properties
            var props = GetReadableProperties(first.GetType());
            if (_table.Headers.Count == 0) _table.Headers.AddRange(props.Select(p => p.Name));
            foreach (var item in seq) {
                if (item == null) { _table.Rows.Add(props.Select(_ => string.Empty).ToArray()); continue; }
                var row = props.Select(p => FormatValue(p.GetValue(item, null))).ToArray();
                _table.Rows.Add(row);
            }
            return this;
        }

        // Plain object → two-column property/value table
        if (_table.Headers.Count == 0) _table.Headers.AddRange(new[] { "Property", "Value" });
        var props2 = GetReadableProperties(data.GetType());
        foreach (var p in props2) {
            _table.Rows.Add(new[] { p.Name, FormatValue(p.GetValue(data, null)) });
        }
        return this;
    }

    private static bool IsScalar(object o) {
        var t = o.GetType();
        return t.IsPrimitive || t.IsEnum || t == typeof(string) || t == typeof(decimal) || t == typeof(DateTime) || t == typeof(Guid) || t == typeof(TimeSpan);
    }

    private static string FormatValue(object? value) {
        if (value is null) return "";
        if (value is string s) return s;
        if (value is System.Collections.IEnumerable en && value is not string) {
            var parts = new List<string>();
            foreach (var it in en) parts.Add(it?.ToString() ?? string.Empty);
            return string.Join(", ", parts);
        }
        return value.ToString() ?? string.Empty;
    }

    private static System.Reflection.PropertyInfo[] GetReadableProperties(System.Type t) {
        return t.GetProperties(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public)
            .Where(p => p.CanRead && p.GetIndexParameters().Length == 0)
            .ToArray();
    }
}
