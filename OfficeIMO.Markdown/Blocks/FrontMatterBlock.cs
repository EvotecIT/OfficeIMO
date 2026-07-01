using System.Globalization;

using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Markdown;

/// <summary>
/// YAML front matter block rendered at the beginning of the document.
/// </summary>
public sealed class FrontMatterBlock : MarkdownBlock, IFrontMatterMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Single front matter entry.</summary>
    public sealed class Entry {
        /// <summary>Entry key.</summary>
        public string Key { get; }
        /// <summary>Entry value.</summary>
        public object? Value { get; }
        /// <summary>Source span for the whole entry when parsed from markdown.</summary>
        public MarkdownSourceSpan? SourceSpan { get; }
        /// <summary>Source span for the key token when parsed from markdown.</summary>
        public MarkdownSourceSpan? KeySourceSpan { get; }
        /// <summary>Source span for the value token or literal-block payload when parsed from markdown.</summary>
        public MarkdownSourceSpan? ValueSourceSpan { get; }

        internal Entry(
            string key,
            object? value,
            MarkdownSourceSpan? keySourceSpan = null,
            MarkdownSourceSpan? valueSourceSpan = null,
            MarkdownSourceSpan? sourceSpan = null) {
            Key = key;
            Value = value;
            SourceSpan = sourceSpan;
            KeySourceSpan = keySourceSpan;
            ValueSourceSpan = valueSourceSpan;
        }
    }

    private readonly List<Entry> _entries = new List<Entry>();

    /// <summary>Structured front matter entries in insertion order.</summary>
    public IReadOnlyList<Entry> Entries => _entries;

    /// <summary>Raw YAML payload between the opening and closing front matter fences when parsed from markdown.</summary>
    public string? RawYaml { get; private set; }

    /// <summary>Source span for the raw YAML payload between the opening and closing front matter fences.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; private set; }

    /// <summary>Finds a front matter entry by key.</summary>
    public Entry? FindEntry(string key, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
        if (string.IsNullOrEmpty(key)) {
            return null;
        }

        for (int i = 0; i < Entries.Count; i++) {
            if (string.Equals(Entries[i].Key, key, comparison)) {
                return Entries[i];
            }
        }

        return null;
    }

    /// <summary>Checks whether the front matter contains an entry with the specified key.</summary>
    public bool HasEntry(string key, StringComparison comparison = StringComparison.OrdinalIgnoreCase) =>
        FindEntry(key, comparison) != null;

    /// <summary>Gets a typed front matter value by key when available.</summary>
    public bool TryGetValue<T>(string key, out T? value) {
        var entry = FindEntry(key);
        if (entry?.Value is T typedValue) {
            value = typedValue;
            return true;
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Creates front matter from an anonymous object or dictionary.
    /// </summary>
    [RequiresUnreferencedCode("Uses reflection over arbitrary runtime types. For AOT-safe usage, use FromObject<T> or pass a dictionary.")]
    public static FrontMatterBlock FromObject(object data) {
        FrontMatterBlock fm = new FrontMatterBlock();
        if (data is IEnumerable<KeyValuePair<string, object?>> dict) {
            foreach (KeyValuePair<string, object?> kv in dict) fm._entries.Add(new Entry(kv.Key, kv.Value));
        } else {
            var props = data.GetType().GetProperties()
                .Where(p => p.CanRead && p.GetIndexParameters().Length == 0 && p.GetMethod != null && p.GetMethod.GetParameters().Length == 0);
            foreach (var p in props) fm._entries.Add(new Entry(p.Name, p.GetValue(data)));
        }
        return fm;
    }

    /// <summary>
    /// Creates front matter from a typed object using public readable properties.
    /// </summary>
    public static FrontMatterBlock FromObject<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(T data) {
        if (data is null) throw new ArgumentNullException(nameof(data));        
        FrontMatterBlock fm = new FrontMatterBlock();
        if (data is IEnumerable<KeyValuePair<string, object?>> dict) {
            foreach (KeyValuePair<string, object?> kv in dict) fm._entries.Add(new Entry(kv.Key, kv.Value));
            return fm;
        }
        var props = typeof(T).GetProperties()
            .Where(p => p.CanRead && p.GetIndexParameters().Length == 0 && p.GetMethod != null && p.GetMethod.GetParameters().Length == 0);
        foreach (var p in props) fm._entries.Add(new Entry(p.Name, p.GetValue(data)));
        return fm;
    }

    internal static FrontMatterBlock FromEntries(IEnumerable<Entry> entries, string? rawYaml = null, MarkdownSourceSpan? bodySourceSpan = null) {
        var fm = new FrontMatterBlock();
        fm.RawYaml = rawYaml;
        fm.BodySourceSpan = bodySourceSpan;
        foreach (var entry in entries) {
            if (entry == null) {
                continue;
            }

            fm._entries.Add(entry);
        }

        return fm;
    }

    /// <summary>Renders the front matter including '---' fences.</summary>
    public string Render() {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("---");
        if (RawYaml != null) {
            sb.Append(RawYaml);
            if (RawYaml.Length > 0 && RawYaml[RawYaml.Length - 1] != '\n') {
                sb.AppendLine();
            }
        } else {
            for (int i = 0; i < Entries.Count; i++) {
                var entry = Entries[i];
                sb.AppendLine(entry.Key + ": " + YamlValue(entry.Value));
            }
        }

        sb.Append("---");
        return sb.ToString();
    }

    string IFrontMatterMarkdownBlock.RenderFrontMatter() => Render();

    private static string YamlValue(object? value) {
        switch (value) {
            case null:
                return string.Empty;
            case string s:
                return EscapeYamlString(s);
            case bool b:
                return b ? "true" : "false";
            case Enum e:
                return EscapeYamlString(e.ToString());
            case IEnumerable<string> ss:
                return "[" + string.Join(", ", ss.Select(EscapeYamlBareOrQuoted)) + "]";
            case System.Collections.IEnumerable ie:
                List<string> vals = new List<string>();
                foreach (object? item in ie) vals.Add(EscapeYamlBareOrQuoted(item?.ToString() ?? string.Empty));
                return "[" + string.Join(", ", vals) + "]";
            default:
                if (IsNumeric(value)) return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                return EscapeYamlString(value.ToString() ?? string.Empty);
        }
    }

    private static bool IsNumeric(object o) => o is sbyte || o is byte || o is short || o is ushort || o is int || o is uint || o is long || o is ulong || o is float || o is double || o is decimal;

    private static string EscapeYamlString(string s) {
        if (string.IsNullOrEmpty(s)) return "\"\"";
        // Multi-line: render as a literal block scalar to preserve newlines safely
        if (s.Contains('\n') || s.Contains('\r')) {
            var lines = s.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var sb = new StringBuilder();
            sb.AppendLine("|");
            foreach (var line in lines) sb.AppendLine("  " + line);
            return sb.ToString().TrimEnd();
        }
        bool needsQuotes = s.IndexOfAny(new[] { ':', '#', '\'', '"', '\t' }) >= 0 || s.Contains(' ');
        if (!needsQuotes) return s;
        // minimal escaping for quotes and backslashes
        string escaped = s.Replace("\\", "\\\\").Replace("\"", "\\\"");
        return "\"" + escaped + "\"";
    }

    private static string EscapeYamlBareOrQuoted(string s) => EscapeYamlString(s);

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => Render();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => string.Empty;
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var children = new List<MarkdownSyntaxNode>();
        var openingFenceSpan = CreateFenceSpan(span, opening: true);
        if (openingFenceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterOpeningFence, openingFenceSpan, "---", associatedObject: this));
        }

        for (int i = 0; i < Entries.Count; i++) {
            var entry = Entries[i];
            if (entry.SourceSpan.HasValue) {
                children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterEntry, entry.SourceSpan, null, associatedObject: entry));
            }

            if (entry.KeySourceSpan.HasValue) {
                children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterKey, entry.KeySourceSpan, entry.Key, associatedObject: entry));
            }

            if (entry.ValueSourceSpan.HasValue) {
                children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterValue, entry.ValueSourceSpan, FormatSyntaxValue(entry.Value), associatedObject: entry));
            }
        }

        if (RawYaml != null && BodySourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterBody, BodySourceSpan, RawYaml, associatedObject: this));
        }

        var closingFenceSpan = CreateFenceSpan(span, opening: false);
        if (closingFenceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterClosingFence, closingFenceSpan, "---", associatedObject: this));
        }

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatter, span, Render(), children, this);
    }

    private static MarkdownSourceSpan? CreateFenceSpan(MarkdownSourceSpan? span, bool opening) {
        if (!span.HasValue) {
            return null;
        }

        var value = span.Value;
        var line = opening ? value.StartLine : value.EndLine;
        int? startOffset = null;
        int? endOffset = null;
        if (opening && value.StartOffset.HasValue) {
            startOffset = value.StartOffset.Value;
            endOffset = startOffset.Value + 2;
        } else if (!opening && value.EndOffset.HasValue && value.EndColumn.HasValue) {
            startOffset = value.EndOffset.Value - Math.Max(0, value.EndColumn.Value - 1);
            endOffset = startOffset.Value + 2;
        }

        return new MarkdownSourceSpan(line, 1, line, 3, startOffset, endOffset);
    }

    internal static string? FormatSyntaxValue(object? value) {
        switch (value) {
            case null:
                return null;
            case bool boolean:
                return boolean ? "true" : "false";
            case IFormattable formattable:
                return formattable.ToString(null, CultureInfo.InvariantCulture);
            case IEnumerable<string> strings:
                return string.Join(", ", strings);
            case System.Collections.IEnumerable values when value is not string:
                var items = new List<string>();
                foreach (object? item in values) {
                    items.Add(item?.ToString() ?? string.Empty);
                }

                return string.Join(", ", items);
            default:
                return value.ToString();
        }
    }
}
