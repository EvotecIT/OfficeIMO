using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using OfficeIMO.Toml;
using OfficeIMO.Yaml;

namespace OfficeIMO.Markdown;

/// <summary>
/// Front matter at the start of a document: YAML between <c>---</c> fences, read and written by <c>OfficeIMO.Yaml</c>,
/// or TOML between <c>+++</c> fences, read and written by <c>OfficeIMO.Toml</c>. <see cref="Yaml"/> or <see cref="Toml"/>
/// holds the lossless document: edits made to it are saved with only the changed text rewritten, and an unedited block
/// renders back to the bytes it was read from. <see cref="Entries"/>, <see cref="Select"/> and <see cref="TryGetValue{T}"/>
/// read both formats the same way.
/// </summary>
public sealed class FrontMatterBlock : MarkdownBlock, IFrontMatterMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>One top-level entry of the front matter, with its Markdown source spans when parsed.</summary>
    public sealed class Entry {
        internal Entry(YamlNode keyNode, string key, YamlNode value, EntrySpans? spans) {
            KeyNode = keyNode;
            Key = key;
            Value = value;
            SourceSpan = spans?.Entry;
            KeySourceSpan = spans?.Key;
            ValueSourceSpan = spans?.Value;
        }

        internal YamlNode KeyNode { get; }

        /// <summary>Entry key text.</summary>
        public string Key { get; }

        /// <summary>
        /// Entry value as a YAML node: read scalars with <see cref="YamlNode.GetValue{T}"/> or
        /// <see cref="YamlNode.AsString"/>, and nested values with <see cref="YamlNode.AsMapping"/>,
        /// <see cref="YamlNode.AsSequence"/> or <see cref="YamlNode.Select"/>. TOML values are read through the same
        /// node types: tables are mappings, arrays and arrays of tables are sequences, and dates keep their TOML
        /// spelling. Edit TOML front matter through <see cref="Toml"/>.
        /// </summary>
        public YamlNode Value { get; }

        /// <summary>Source span from the key to the end of the value when parsed from Markdown.</summary>
        public MarkdownSourceSpan? SourceSpan { get; }

        /// <summary>Source span of the key when parsed from Markdown.</summary>
        public MarkdownSourceSpan? KeySourceSpan { get; }

        /// <summary>
        /// Source span of the value node, including a YAML block scalar's header, when parsed from Markdown. TOML
        /// tables and dotted keys have no single value span.
        /// </summary>
        public MarkdownSourceSpan? ValueSourceSpan { get; }
    }

    internal sealed class EntrySpans {
        public EntrySpans(MarkdownSourceSpan? entry, MarkdownSourceSpan? key, MarkdownSourceSpan? value) {
            Entry = entry;
            Key = key;
            Value = value;
        }

        public MarkdownSourceSpan? Entry { get; }
        public MarkdownSourceSpan? Key { get; }
        public MarkdownSourceSpan? Value { get; }
    }

    internal const string YamlFence = "---";
    internal const string TomlFence = "+++";

    private readonly IReadOnlyDictionary<YamlNode, EntrySpans>? _spans;
    private readonly IReadOnlyDictionary<string, EntrySpans>? _keySpans;
    private readonly string _openingFence;
    private readonly string _closingFence;
    private Entry[]? _entries;
    private string? _tomlViewText;
    private YamlMapping? _tomlView;

    private FrontMatterBlock(MarkdownFrontMatterFormat format, YamlDocument? yaml, TomlDocument? toml, string? rawText,
        MarkdownSourceSpan? bodySourceSpan, IReadOnlyList<OfficeDiagnostic> diagnostics,
        IReadOnlyDictionary<YamlNode, EntrySpans>? spans, IReadOnlyDictionary<string, EntrySpans>? keySpans,
        string? openingFence = null, string? closingFence = null) {
        Format = format;
        Yaml = yaml;
        Toml = toml;
        RawText = rawText;
        BodySourceSpan = bodySourceSpan;
        Diagnostics = diagnostics;
        _spans = spans;
        _keySpans = keySpans;
        _openingFence = openingFence ?? FenceText;
        _closingFence = closingFence ?? FenceText;
    }

    /// <summary>Gets the front matter language: YAML between <c>---</c> fences or TOML between <c>+++</c> fences.</summary>
    public MarkdownFrontMatterFormat Format { get; }

    /// <summary>
    /// Gets the YAML document of YAML front matter, or <see langword="null"/> for TOML front matter and when the text
    /// between the fences is not valid YAML; then <see cref="Diagnostics"/> holds the error and the block renders
    /// <see cref="RawText"/> unchanged.
    /// </summary>
    public YamlDocument? Yaml { get; }

    /// <summary>
    /// Gets the TOML document of TOML front matter, or <see langword="null"/> for YAML front matter and when the text
    /// between the fences is not valid TOML; then <see cref="Diagnostics"/> holds the error and the block renders
    /// <see cref="RawText"/> unchanged.
    /// </summary>
    public TomlDocument? Toml { get; }

    /// <summary>
    /// Gets the root mapping of the front matter, or <see langword="null"/> when it is empty, invalid or not a mapping.
    /// For TOML front matter it is a read-only view of the root table that follows edits made through <see cref="Toml"/>.
    /// </summary>
    public YamlMapping? Mapping {
        get {
            if (Yaml?.Root is { } root) return Resolve(root) as YamlMapping;
            YamlMapping? view = Toml == null ? null : TomlView(Toml);
            return view == null || view.Count == 0 ? null : view;
        }
    }

    /// <summary>
    /// Gets the top-level entries of <see cref="Mapping"/> in source order, with Markdown source spans for entries that
    /// were read from Markdown. The list follows edits made through <see cref="Yaml"/> or <see cref="Toml"/>.
    /// </summary>
    public IReadOnlyList<Entry> Entries {
        get {
            YamlMapping? mapping = Mapping;
            if (mapping == null || mapping.Count == 0) return Array.Empty<Entry>();
            if (_entries != null && IsCurrent(_entries, mapping)) return _entries;

            var entries = new Entry[mapping.Count];
            for (int i = 0; i < entries.Length; i++) {
                YamlMappingEntry entry = mapping[i];
                string key = KeyText(entry.Key);
                EntrySpans? spans = null;
                if (_spans != null) _spans.TryGetValue(entry.Key, out spans);
                else _keySpans?.TryGetValue(key, out spans);
                entries[i] = new Entry(entry.Key, key, entry.Value, spans);
            }

            _entries = entries;
            return entries;
        }
    }

    /// <summary>Gets the text between the fences as it was read from Markdown, or <see langword="null"/> for front matter built in code.</summary>
    public string? RawText { get; }

    /// <summary>Gets the source span of the text between the fences when parsed from Markdown.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; }

    /// <summary>
    /// Gets findings from reading the front matter: a <see cref="OfficeDiagnosticCodes.MarkdownFrontMatterInvalidYaml"/>
    /// or <see cref="OfficeDiagnosticCodes.MarkdownFrontMatterInvalidToml"/> error located in the Markdown source when the
    /// text is invalid, and otherwise the YAML document's own findings, such as repeated keys, located within the front
    /// matter text. Valid TOML has no findings.
    /// </summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }

    /// <summary>
    /// Finds a top-level entry by key. When keys repeat under the comparison, the last one wins, as it does for YAML
    /// loaders that accept repeated keys.
    /// </summary>
    public Entry? FindEntry(string key, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
        if (string.IsNullOrEmpty(key)) return null;

        IReadOnlyList<Entry> entries = Entries;
        for (int i = entries.Count - 1; i >= 0; i--) {
            if (string.Equals(entries[i].Key, key, comparison)) return entries[i];
        }

        return null;
    }

    /// <summary>Checks whether the front matter has a top-level entry with the key.</summary>
    public bool HasEntry(string key, StringComparison comparison = StringComparison.OrdinalIgnoreCase) =>
        FindEntry(key, comparison) != null;

    /// <summary>
    /// Reads a top-level scalar entry as <typeparamref name="T"/> through the YAML schema, as
    /// <see cref="YamlNode.TryGetValue{T}"/> does: <see cref="string"/>, <see cref="bool"/>, <see cref="int"/>,
    /// <see cref="long"/>, <see cref="double"/>, <see cref="float"/>, <see cref="decimal"/> and their nullable forms.
    /// Returns <see langword="false"/> when the key is missing, the value is a collection or it does not convert.
    /// </summary>
    /// <exception cref="NotSupportedException"><typeparamref name="T"/> is not a supported type.</exception>
    public bool TryGetValue<T>(string key, out T? value) {
        Entry? entry = FindEntry(key);
        if (entry != null && entry.Value.TryGetValue(out T converted)) {
            value = converted;
            return true;
        }

        value = default;
        return false;
    }

    /// <summary>Returns the node a path such as <c>author.name</c> or <c>tags[0]</c> addresses, or <see langword="null"/>.</summary>
    /// <remarks>
    /// Paths use the grammar of <see cref="YamlNode.Select"/>, which TOML paths share, and match keys exactly. TOML front
    /// matter is read through the same view as <see cref="Mapping"/>.
    /// </remarks>
    public YamlNode? Select(string path) {
        if (Yaml != null) return Yaml.Root == null ? null : Yaml.Select(path);
        return Toml == null ? null : TomlView(Toml).Select(path);
    }

    /// <summary>
    /// Creates front matter from an anonymous object, a dictionary or another object's public properties. Nested
    /// dictionaries, objects and collections become YAML mappings and sequences, and <see cref="YamlNode"/> values are
    /// used as they are.
    /// </summary>
    [RequiresUnreferencedCode("Uses reflection over arbitrary runtime types. For AOT-safe usage, use FromObject<T> or pass a dictionary.")]
    public static FrontMatterBlock FromObject(object data) {
        ArgumentNullException.ThrowIfNull(data);
        return FromMapping(FrontMatterValues.ToMapping(data, FrontMatterValues.ReflectObject));
    }

    /// <summary>
    /// Creates front matter from a typed object's public properties or from a dictionary. Nested dictionaries and
    /// collections become YAML mappings and sequences; other nested objects are written as their invariant text.
    /// </summary>
    public static FrontMatterBlock FromObject<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(T data) {
        ArgumentNullException.ThrowIfNull(data);
        YamlMapping mapping = FrontMatterValues.TryToMapping(data, complexValue: null) ?? FrontMatterValues.FromProperties(data, typeof(T), complexValue: null);
        return FromMapping(mapping);
    }

    /// <summary>
    /// Creates TOML front matter, written between <c>+++</c> fences, from a TOML document. The block keeps the document,
    /// so edits made to it later are part of the rendered front matter.
    /// </summary>
    public static FrontMatterBlock FromToml(TomlDocument document) {
        ArgumentNullException.ThrowIfNull(document);
        return new FrontMatterBlock(MarkdownFrontMatterFormat.Toml, null, document, rawText: null, bodySourceSpan: null,
            Array.Empty<OfficeDiagnostic>(), spans: null, keySpans: null);
    }

    private static FrontMatterBlock FromMapping(YamlMapping mapping) {
        YamlDocument yaml = YamlDocument.Create();
        yaml.Root = mapping;
        return new FrontMatterBlock(MarkdownFrontMatterFormat.Yaml, yaml, null, rawText: null, bodySourceSpan: null,
            Array.Empty<OfficeDiagnostic>(), spans: null, keySpans: null);
    }

    internal static FrontMatterBlock FromParsed(YamlDocument yaml, string rawText, MarkdownSourceSpan? bodySourceSpan,
        IReadOnlyDictionary<YamlNode, EntrySpans>? spans, string openingFence, string closingFence) =>
        new FrontMatterBlock(MarkdownFrontMatterFormat.Yaml, yaml, null, rawText, bodySourceSpan, yaml.Diagnostics, spans,
            keySpans: null, openingFence, closingFence);

    internal static FrontMatterBlock FromParsed(TomlDocument toml, string rawText, MarkdownSourceSpan? bodySourceSpan,
        IReadOnlyDictionary<string, EntrySpans>? keySpans, string openingFence, string closingFence) =>
        new FrontMatterBlock(MarkdownFrontMatterFormat.Toml, null, toml, rawText, bodySourceSpan, Array.Empty<OfficeDiagnostic>(),
            spans: null, keySpans, openingFence, closingFence);

    internal static FrontMatterBlock FromInvalid(MarkdownFrontMatterFormat format, string rawText, MarkdownSourceSpan? bodySourceSpan,
        OfficeDiagnostic diagnostic, string openingFence, string closingFence) =>
        new FrontMatterBlock(format, null, null, rawText, bodySourceSpan, new[] { diagnostic }, spans: null, keySpans: null,
            openingFence, closingFence);

    /// <summary>The fence line of this block's format: <c>---</c> for YAML, <c>+++</c> for TOML.</summary>
    internal string FenceText => Format == MarkdownFrontMatterFormat.Toml ? TomlFence : YamlFence;

    /// <summary>
    /// Renders the front matter with its fences. Unedited front matter read from Markdown renders the text it was read
    /// from; edits made through <see cref="Yaml"/> or <see cref="Toml"/> rewrite only what changed.
    /// </summary>
    public string Render() {
        string body = Yaml?.ToString() ?? Toml?.ToString() ?? RawText ?? string.Empty;
        var sb = new StringBuilder(body.Length + _openingFence.Length + _closingFence.Length + 2);
        sb.Append(_openingFence).Append('\n').Append(body);
        if (body.Length > 0 && body[body.Length - 1] != '\n') sb.Append('\n');
        sb.Append(_closingFence);
        return sb.ToString();
    }

    /// <summary>
    /// Returns the YAML view of the TOML root table, rebuilt when the document's text has changed since the last call, so
    /// it follows edits without a change counter on the document.
    /// </summary>
    private YamlMapping TomlView(TomlDocument toml) {
        string text = toml.ToString();
        if (_tomlView == null || !string.Equals(text, _tomlViewText, StringComparison.Ordinal)) {
            _tomlView = FrontMatterTomlView.ToMapping(toml.Root);
            _tomlViewText = text;
        }

        return _tomlView;
    }

    string IFrontMatterMarkdownBlock.RenderFrontMatter() => Render();

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => Render();

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => string.Empty;

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var children = new List<MarkdownSyntaxNode>();
        var openingFenceSpan = CreateFenceSpan(span, opening: true);
        if (openingFenceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterOpeningFence, openingFenceSpan, FenceText, associatedObject: this));
        }

        IReadOnlyList<Entry> entries = Entries;
        for (int i = 0; i < entries.Count; i++) {
            var entry = entries[i];
            var entryChildren = new List<MarkdownSyntaxNode>(2);
            if (entry.KeySourceSpan.HasValue) {
                entryChildren.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterKey, entry.KeySourceSpan, entry.Key, associatedObject: entry));
            }

            if (entry.ValueSourceSpan.HasValue) {
                entryChildren.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterValue, entry.ValueSourceSpan, FormatValue(entry.Value), associatedObject: entry));
            }

            if (entry.SourceSpan.HasValue || entryChildren.Count > 0) {
                children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterEntry, entry.SourceSpan, children: entryChildren, associatedObject: entry));
            }
        }

        if (RawText != null && BodySourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterBody, BodySourceSpan, RawText, associatedObject: this));
        }

        var closingFenceSpan = CreateFenceSpan(span, opening: false);
        if (closingFenceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterClosingFence, closingFenceSpan, FenceText, associatedObject: this));
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

    private static bool IsCurrent(Entry[] entries, YamlMapping mapping) {
        if (entries.Length != mapping.Count) return false;
        for (int i = 0; i < entries.Length; i++) {
            YamlMappingEntry entry = mapping[i];
            if (!ReferenceEquals(entries[i].KeyNode, entry.Key) || !ReferenceEquals(entries[i].Value, entry.Value)) return false;
        }

        return true;
    }

    internal static string KeyText(YamlNode key) => Resolve(key) is YamlScalar scalar ? scalar.AsString() : key.ToString();

    private static YamlNode Resolve(YamlNode node) => node is YamlAlias alias ? alias.Target : node;

    /// <summary>
    /// Formats a value for people to read: a scalar's text without trailing line breaks, sequence items joined with
    /// commas and mapping entries as <c>key: value</c> joined with semicolons; nested collections are bracketed.
    /// </summary>
    internal static string? FormatValue(YamlNode? value) {
        if (value == null || value.IsNull) return null;
        var sb = new StringBuilder();
        AppendValue(sb, value, nested: false);
        return sb.ToString();
    }

    private static void AppendValue(StringBuilder sb, YamlNode node, bool nested) {
        node = Resolve(node);
        switch (node) {
            case YamlSequence sequence:
                if (nested) sb.Append('[');
                for (int i = 0; i < sequence.Count; i++) {
                    if (i > 0) sb.Append(", ");
                    AppendValue(sb, sequence[i], nested: true);
                }

                if (nested) sb.Append(']');
                break;
            case YamlMapping mapping:
                if (nested) sb.Append('{');
                for (int i = 0; i < mapping.Count; i++) {
                    if (i > 0) sb.Append(nested ? ", " : "; ");
                    YamlMappingEntry entry = mapping[i];
                    sb.Append(KeyText(entry.Key)).Append(": ");
                    AppendValue(sb, entry.Value, nested: true);
                }

                if (nested) sb.Append('}');
                break;
            default:
                if (!node.IsNull) sb.Append(node.AsString().TrimEnd('\r', '\n'));
                break;
        }
    }
}
