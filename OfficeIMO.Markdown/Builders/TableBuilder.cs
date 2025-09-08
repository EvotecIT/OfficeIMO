using System;
using System.Collections.Generic;

namespace OfficeIMO.Markdown;

/// <summary>
/// Builder for pipe tables.
/// </summary>
public sealed class TableBuilder {
    private const int MaxRows = 10000;
    private const int MaxColumns = 100;
    private readonly TableBlock _table = new TableBlock();
    private TableFromOptions? _defaultOptions;
    private IReadOnlyList<string> NormalizeRow(IReadOnlyList<string> cells) {
        if (cells == null) return Array.Empty<string>();
        int count = cells.Count;
        int target = _table.Headers.Count > 0 ? _table.Headers.Count : count;
        if (target <= 0) return cells;
        if (count == target) return cells;
        var list = new List<string>(cells);
        if (count > target) return list.GetRange(0, target);
        while (list.Count < target) list.Add(string.Empty);
        return list;
    }
    /// <summary>Sets the header row.</summary>
    public TableBuilder Headers(params string[] headers) {
        var hs = headers ?? Array.Empty<string>();
        for (int i = 0; i < hs.Length && i < MaxColumns; i++) _table.Headers.Add(hs[i]?.Trim() ?? string.Empty);
        return this;
    }
    /// <summary>Adds a data row.</summary>
    public TableBuilder Row(params string[] cells) { if (_table.Rows.Count < MaxRows) _table.Rows.Add(NormalizeRow(cells?.ToArray() ?? Array.Empty<string>())); return this; }
    /// <summary>Adds multiple rows.</summary>
    public TableBuilder Rows(IEnumerable<IReadOnlyList<string>> rows) { foreach (IReadOnlyList<string> r in rows) { if (_table.Rows.Count >= MaxRows) break; _table.Rows.Add(NormalizeRow(r)); } return this; }
    /// <summary>Adds two-column rows from tuples.</summary>
    public TableBuilder Rows(IEnumerable<(string, string)> rows) { foreach ((string a, string b) in rows) { if (_table.Rows.Count >= MaxRows) break; _table.Rows.Add(NormalizeRow(new[] { a, b })); } return this; }
    /// <summary>Adds two-column rows from key/value pairs.</summary>
    public TableBuilder Rows(IEnumerable<KeyValuePair<string, string>> rows) { foreach (KeyValuePair<string, string> kv in rows) { if (_table.Rows.Count >= MaxRows) break; _table.Rows.Add(NormalizeRow(new[] { kv.Key, kv.Value })); } return this; }
    internal TableBlock Build() => _table;

    /// <summary>
    /// Populates the table from an arbitrary object. If <paramref name="data"/> is
    /// - a sequence of scalars → one-column table with a row per value,
    /// - a sequence of objects → headers from public readable properties and rows per item,
    /// - a dictionary → two columns Key/Value,
    /// - a POCO → two columns Property/Value.
    /// </summary>
    public TableBuilder FromAny(object? data) { return FromAny(data, _defaultOptions); }

    /// <summary>
    /// Populates the table from arbitrary data with options (include/exclude/order).
    /// </summary>
    public TableBuilder FromAny(object? data, TableFromOptions? options) {
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
            foreach (var item in seq) { if (_table.Rows.Count >= MaxRows) break; _table.Rows.Add(new[] { FormatValue(item) }); }
                return this;
            }

            // Object sequence → headers from public readable properties
            var props = SelectProperties(first.GetType(), options);
            if (_table.Headers.Count == 0) _table.Headers.AddRange(props.Select(p => Rename(p.Name, options)));
            foreach (var item in seq) {
                if (item == null) { _table.Rows.Add(props.Select(_ => string.Empty).ToArray()); continue; }
                var row = props.Select(p => FormatValue(p.GetValue(item, null), p.Name, options)).ToArray();
                if (_table.Rows.Count >= MaxRows) break; _table.Rows.Add(row);
            }
            if (options?.Alignments != null && options.Alignments.Count > 0) { _table.Alignments.Clear(); _table.Alignments.AddRange(options.Alignments); }
            return this;
        }

        // Plain object → either a wide single-row table (when options indicate selection/order/renames)
        // or a two-column Property/Value table by default.
        var props2 = SelectProperties(data.GetType(), options);
        bool wide = options != null && (
            (options.Include != null && options.Include.Count > 0) ||
            (options.Order != null && options.Order.Count > 0) ||
            (options.HeaderRenames != null && options.HeaderRenames.Count > 0) ||
            (options.Formatters != null && options.Formatters.Count > 0)
        );
        if (wide) {
            if (_table.Headers.Count == 0) _table.Headers.AddRange(props2.Select(p => Rename(p.Name, options)));
            var limitedProps = props2.Length > MaxColumns ? props2.AsSpan(0, MaxColumns).ToArray() : props2;
            var row = new List<string>(limitedProps.Length);
            foreach (var p in limitedProps) row.Add(FormatValue(p.GetValue(data, null), p.Name, options));
            _table.Rows.Add(row);
            if (options?.Alignments != null && options.Alignments.Count > 0) { _table.Alignments.Clear(); _table.Alignments.AddRange(options.Alignments); }
            return this;
        } else {
            if (_table.Headers.Count == 0) _table.Headers.AddRange(new[] { "Property", "Value" });
            foreach (var p in props2) {
                _table.Rows.Add(new[] { Rename(p.Name, options), FormatValue(p.GetValue(data, null), p.Name, options) });
            }
            return this;
        }
    }

    /// <summary>
    /// Populates the table from arbitrary data using an inline options configuration.
    /// </summary>
    public TableBuilder FromAny(object? data, System.Action<TableFromOptions> configure) {
        var opts = new TableFromOptions();
        configure(opts);
        return FromAny(data, opts);
    }

    /// <summary>
    /// Populates the table from a sequence using explicit column selectors.
    /// </summary>
    public TableBuilder FromSequence<T>(IEnumerable<T> items, params (string Header, System.Func<T, object?> Selector)[] columns) {
        if (columns == null || columns.Length == 0 || items == null) return this;
        if (_table.Headers.Count == 0) _table.Headers.AddRange(columns.Select(c => c.Header ?? string.Empty));
        foreach (var item in items) {
            var row = new List<string>(columns.Length);
            foreach (var c in columns) {
                var selector = c.Selector ?? (_ => null);
                row.Add(FormatValue(selector(item)));
            }
            _table.Rows.Add(row);
        }
        return this;
    }

    private static bool IsScalar(object o) {
        var t = o.GetType();
        return t.IsPrimitive || t.IsEnum || t == typeof(string) || t == typeof(decimal) || t == typeof(DateTime) || t == typeof(Guid) || t == typeof(TimeSpan);
    }

    private static string FormatValue(object? value, string? propertyName = null, TableFromOptions? options = null) {
        if (propertyName != null && options != null && options.Formatters.TryGetValue(propertyName, out var fmt)) {
            try { return fmt(value); } catch { /* ignore and fallback */ }
        }
        if (value is null) return "";
        if (value is string s) return s;
        if (value is System.Collections.IEnumerable en && value is not string) {
            var parts = new List<string>();
            foreach (var it in en) parts.Add(it?.ToString() ?? string.Empty);
            return string.Join(", ", parts);
        }
        return value.ToString() ?? string.Empty;
    }

    private static readonly System.Collections.Concurrent.ConcurrentDictionary<System.Type, System.Reflection.PropertyInfo[]> _propsCache = new();
    private static System.Reflection.PropertyInfo[] GetReadableProperties(System.Type t) {
        return _propsCache.GetOrAdd(t, static tArg =>
            tArg.GetProperties(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public)
                .Where(p => p.CanRead && p.GetIndexParameters().Length == 0)
                .ToArray());
    }

    private static System.Reflection.PropertyInfo[] SelectProperties(System.Type t, TableFromOptions? options) {
        var props = GetReadableProperties(t);
        if (options == null) return props;
        var list = new List<System.Reflection.PropertyInfo>(props);
        if (options.Include != null && options.Include.Count > 0) {
            list = list.Where(p => options.Include.Contains(p.Name)).ToList();
        }
        if (options.Exclude != null && options.Exclude.Count > 0) {
            list = list.Where(p => !options.Exclude.Contains(p.Name)).ToList();
        }
        if (options.Order != null && options.Order.Count > 0) {
            var orderMap = options.Order.Select((name, idx) => new { name, idx }).ToDictionary(x => x.name, x => x.idx);
            list = list.OrderBy(p => orderMap.ContainsKey(p.Name) ? orderMap[p.Name] : int.MaxValue).ToList();
        }
        return list.ToArray();
    }

    private static string Rename(string name, TableFromOptions? options) {
        if (options?.HeaderRenames != null && options.HeaderRenames.TryGetValue(name, out var newName)) return newName;
        if (options?.HeaderTransform != null) return options.HeaderTransform(name);
        return name;
    }

    /// <summary>Sets column alignments for the table header/columns.</summary>
    public TableBuilder Align(params ColumnAlignment[] alignments) { _table.Alignments.Clear(); _table.Alignments.AddRange(alignments); return this; }
    /// <summary>Sets a uniform alignment for all columns (applied to header + cells).</summary>
    public TableBuilder AlignAll(ColumnAlignment alignment) {
        int cols = _table.Headers.Count > 0 ? _table.Headers.Count : (_table.Rows.Count > 0 ? _table.Rows[0].Count : 0);
        _table.Alignments.Clear();
        for (int i = 0; i < cols; i++) _table.Alignments.Add(alignment);
        return this;
    }

    /// <summary>Set left alignment on specified 0-based column indexes. If none provided, all columns.</summary>
    public TableBuilder AlignLeft(params int[] cols) => AlignPreset(ColumnAlignment.Left, cols);
    /// <summary>Set right alignment on specified 0-based column indexes. If none provided, all columns.</summary>
    public TableBuilder AlignRight(params int[] cols) => AlignPreset(ColumnAlignment.Right, cols);
    /// <summary>Set center alignment on specified 0-based column indexes. If none provided, all columns.</summary>
    public TableBuilder AlignCenter(params int[] cols) => AlignPreset(ColumnAlignment.Center, cols);
    /// <summary>Remove alignment (default) on specified 0-based column indexes. If none provided, all columns.</summary>
    public TableBuilder AlignNone(params int[] cols) => AlignPreset(ColumnAlignment.None, cols);

    /// <summary>Align columns by matching header names (case-insensitive).</summary>
    public TableBuilder AlignByHeaders(ColumnAlignment alignment, params string[] headerNames) {
        if (headerNames == null || headerNames.Length == 0) return this;
        if (_table.Headers.Count == 0) return this;
        var set = new HashSet<string>(headerNames.Where(h => !string.IsNullOrWhiteSpace(h)), StringComparer.OrdinalIgnoreCase);
        var idxs = new List<int>();
        for (int i = 0; i < _table.Headers.Count; i++) {
            if (set.Contains(_table.Headers[i])) idxs.Add(i);
        }
        return AlignPreset(alignment, idxs.ToArray());
    }

    private TableBuilder AlignPreset(ColumnAlignment alignment, params int[] cols) {
        int count = _table.Headers.Count > 0 ? _table.Headers.Count : (_table.Rows.Count > 0 ? _table.Rows[0].Count : 0);
        if (count <= 0) return this;
        if (_table.Alignments.Count < count) { for (int i = _table.Alignments.Count; i < count; i++) _table.Alignments.Add(ColumnAlignment.None); }
        if (cols == null || cols.Length == 0) { for (int i = 0; i < count; i++) _table.Alignments[i] = alignment; }
        else { foreach (var c in cols) if (c >= 0 && c < count) _table.Alignments[c] = alignment; }
        return this;
    }

    /// <summary>Guess numeric columns by sampling values and align them right. Threshold is fraction of numeric-like values required (0..1).</summary>
    public TableBuilder AlignNumericRight(double threshold = 0.8) {
        if (_table.Rows.Count == 0) return this;
        int cols = _table.Headers.Count > 0 ? _table.Headers.Count : _table.Rows[0].Count;
        if (_table.Alignments.Count < cols) { for (int i = _table.Alignments.Count; i < cols; i++) _table.Alignments.Add(ColumnAlignment.None); }
        for (int c = 0; c < cols; c++) {
            int total = 0; int numeric = 0;
            foreach (var row in _table.Rows) {
                if (c >= row.Count) continue;
                var cell = row[c];
                total++;
                if (LooksNumeric(cell)) numeric++;
            }
            if (total > 0 && (double)numeric / total >= threshold) _table.Alignments[c] = ColumnAlignment.Right;
        }
        // Default unspecified alignments to Left to produce explicit alignment markers (:---) for readability
        for (int c = 0; c < cols; c++) if (_table.Alignments[c] == ColumnAlignment.None) _table.Alignments[c] = ColumnAlignment.Left;
        return this;
    }

    private static bool LooksNumeric(string? s) {
        if (string.IsNullOrWhiteSpace(s)) return false;
        string s2 = s!.Trim();
        // Strip percent at end
        if (s2.EndsWith("%")) s2 = s2.Substring(0, s2.Length - 1).Trim();
        // Remove currency symbols anywhere
        var chars = new System.Text.StringBuilder(s2.Length);
        foreach (var ch in s2) {
            var cat = char.GetUnicodeCategory(ch);
            if (cat == System.Globalization.UnicodeCategory.CurrencySymbol) continue;
            chars.Append(ch);
        }
        s2 = chars.ToString();
        // Try parse with both invariant and current cultures
        var inv = System.Globalization.CultureInfo.InvariantCulture;
        var cur = System.Globalization.CultureInfo.CurrentCulture;
        return double.TryParse(s2, System.Globalization.NumberStyles.Any, inv, out _)
            || double.TryParse(s2, System.Globalization.NumberStyles.Any, cur, out _);
    }

    private static bool LooksDate(string? s) {
        if (string.IsNullOrWhiteSpace(s)) return false;
        string s2 = s!.Trim();
        // Try DateTimeOffset (captures ISO 8601 etc.) and DateTime using current culture
        if (DateTimeOffset.TryParse(s2, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AllowWhiteSpaces, out _)) return true;
        if (DateTimeOffset.TryParse(s2, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.AllowWhiteSpaces, out _)) return true;
        if (DateTime.TryParse(s2, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.AllowWhiteSpaces, out _)) return true;
        // Common explicit formats
        string[] formats = new [] { "yyyy-MM-dd", "MM/dd/yyyy", "dd.MM.yyyy", "yyyyMMdd", "dd/MM/yyyy" };
        foreach (var fmt in formats) if (DateTime.TryParseExact(s2, fmt, null, System.Globalization.DateTimeStyles.None, out _)) return true;
        return false;
    }

    /// <summary>Configures default column options for subsequent FromAny calls.</summary>
    public TableBuilder Columns(System.Action<TableFromOptions> configure) { _defaultOptions ??= new TableFromOptions(); configure(_defaultOptions); return this; }
    /// <summary>Sets a simple header transform used for generating header text (e.g., prettifying PascalCase).</summary>
    public TableBuilder Columns(System.Func<string, string> headerTransform) { _defaultOptions ??= new TableFromOptions(); _defaultOptions.HeaderTransform = headerTransform; return this; }

    /// <summary>Guess date-like columns and center-align them. Threshold is fraction of date-like values required (0..1).</summary>
    public TableBuilder AlignDatesCenter(double threshold = 0.6) {
        if (_table.Rows.Count == 0) return this;
        int cols = _table.Headers.Count > 0 ? _table.Headers.Count : _table.Rows[0].Count;
        if (_table.Alignments.Count < cols) { for (int i = _table.Alignments.Count; i < cols; i++) _table.Alignments.Add(ColumnAlignment.None); }
        for (int c = 0; c < cols; c++) {
            int total = 0; int dates = 0;
            foreach (var row in _table.Rows) {
                if (c >= row.Count) continue;
                var cell = row[c];
                total++;
                if (LooksDate(cell)) dates++;
            }
            if (total > 0 && (double)dates / total >= threshold) _table.Alignments[c] = ColumnAlignment.Center;
        }
        return this;
    }
}
