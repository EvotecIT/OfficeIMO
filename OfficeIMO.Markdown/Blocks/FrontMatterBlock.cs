using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using OfficeIMO.Yaml;

namespace OfficeIMO.Markdown;

/// <summary>
/// YAML front matter at the start of a document, read and written by <c>OfficeIMO.Yaml</c>. <see cref="Yaml"/> holds
/// the lossless document: nested values are reachable through it, edits made to it are saved with only the changed text
/// rewritten, and an unedited block renders back to the bytes it was read from.
/// </summary>
public sealed class FrontMatterBlock : MarkdownBlock, IFrontMatterMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>One top-level entry of the front matter mapping, with its Markdown source spans when parsed.</summary>
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
        /// <see cref="YamlNode.AsSequence"/> or <see cref="YamlNode.Select"/>.
        /// </summary>
        public YamlNode Value { get; }

        /// <summary>Source span from the key to the end of the value when parsed from Markdown.</summary>
        public MarkdownSourceSpan? SourceSpan { get; }

        /// <summary>Source span of the key when parsed from Markdown.</summary>
        public MarkdownSourceSpan? KeySourceSpan { get; }

        /// <summary>Source span of the value node, including a block scalar's header, when parsed from Markdown.</summary>
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

    internal const string Fence = "---";

    private readonly IReadOnlyDictionary<YamlNode, EntrySpans>? _spans;
    private readonly string _openingFence;
    private readonly string _closingFence;
    private Entry[]? _entries;

    private FrontMatterBlock(YamlDocument? yaml, string? rawYaml, MarkdownSourceSpan? bodySourceSpan,
        IReadOnlyList<OfficeDiagnostic> diagnostics, IReadOnlyDictionary<YamlNode, EntrySpans>? spans,
        string openingFence = Fence, string closingFence = Fence) {
        Yaml = yaml;
        RawYaml = rawYaml;
        BodySourceSpan = bodySourceSpan;
        Diagnostics = diagnostics;
        _spans = spans;
        _openingFence = openingFence;
        _closingFence = closingFence;
    }

    /// <summary>
    /// Gets the YAML document, or <see langword="null"/> when the text between the fences is not valid YAML; then
    /// <see cref="Diagnostics"/> holds the error and the block renders <see cref="RawYaml"/> unchanged.
    /// </summary>
    public YamlDocument? Yaml { get; }

    /// <summary>Gets the root mapping of the front matter, or <see langword="null"/> when it is empty, invalid or not a mapping.</summary>
    public YamlMapping? Mapping => Yaml?.Root is { } root ? Resolve(root) as YamlMapping : null;

    /// <summary>
    /// Gets the top-level entries of <see cref="Mapping"/> in source order, with Markdown source spans for entries that
    /// were read from Markdown. The list follows edits made through <see cref="Yaml"/>.
    /// </summary>
    public IReadOnlyList<Entry> Entries {
        get {
            YamlMapping? mapping = Mapping;
            if (mapping == null || mapping.Count == 0) return Array.Empty<Entry>();
            if (_entries != null && IsCurrent(_entries, mapping)) return _entries;

            var entries = new Entry[mapping.Count];
            for (int i = 0; i < entries.Length; i++) {
                YamlMappingEntry entry = mapping[i];
                EntrySpans? spans = null;
                _spans?.TryGetValue(entry.Key, out spans);
                entries[i] = new Entry(entry.Key, KeyText(entry.Key), entry.Value, spans);
            }

            _entries = entries;
            return entries;
        }
    }

    /// <summary>Gets the text between the fences as it was read from Markdown, or <see langword="null"/> for front matter built in code.</summary>
    public string? RawYaml { get; }

    /// <summary>Gets the source span of the text between the fences when parsed from Markdown.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; }

    /// <summary>
    /// Gets findings from reading the front matter: a <see cref="OfficeDiagnosticCodes.MarkdownFrontMatterInvalidYaml"/>
    /// error located in the Markdown source when the YAML is invalid, and otherwise the YAML document's own findings,
    /// such as repeated keys, located within the front matter text.
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

    /// <summary>Returns the node a YAML path such as <c>author.name</c> or <c>tags[0]</c> addresses, or <see langword="null"/>.</summary>
    /// <remarks>Paths use the grammar of <see cref="YamlNode.Select"/> and match keys exactly.</remarks>
    public YamlNode? Select(string path) => Yaml?.Root == null ? null : Yaml.Select(path);

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

    private static FrontMatterBlock FromMapping(YamlMapping mapping) {
        YamlDocument yaml = YamlDocument.Create();
        yaml.Root = mapping;
        return new FrontMatterBlock(yaml, rawYaml: null, bodySourceSpan: null, Array.Empty<OfficeDiagnostic>(), spans: null);
    }

    internal static FrontMatterBlock FromParsed(YamlDocument yaml, string rawYaml, MarkdownSourceSpan? bodySourceSpan,
        IReadOnlyDictionary<YamlNode, EntrySpans>? spans, string openingFence, string closingFence) =>
        new FrontMatterBlock(yaml, rawYaml, bodySourceSpan, yaml.Diagnostics, spans, openingFence, closingFence);

    internal static FrontMatterBlock FromInvalid(string rawYaml, MarkdownSourceSpan? bodySourceSpan, OfficeDiagnostic diagnostic,
        string openingFence, string closingFence) =>
        new FrontMatterBlock(null, rawYaml, bodySourceSpan, new[] { diagnostic }, spans: null, openingFence, closingFence);

    /// <summary>
    /// Renders the front matter with its fences. Unedited front matter read from Markdown renders the text it was read
    /// from; edits made through <see cref="Yaml"/> rewrite only what changed.
    /// </summary>
    public string Render() {
        string yaml = Yaml?.ToString() ?? RawYaml ?? string.Empty;
        var sb = new StringBuilder(yaml.Length + _openingFence.Length + _closingFence.Length + 2);
        sb.Append(_openingFence).Append('\n').Append(yaml);
        if (yaml.Length > 0 && yaml[yaml.Length - 1] != '\n') sb.Append('\n');
        sb.Append(_closingFence);
        return sb.ToString();
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
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterOpeningFence, openingFenceSpan, Fence, associatedObject: this));
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

        if (RawYaml != null && BodySourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterBody, BodySourceSpan, RawYaml, associatedObject: this));
        }

        var closingFenceSpan = CreateFenceSpan(span, opening: false);
        if (closingFenceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.FrontMatterClosingFence, closingFenceSpan, Fence, associatedObject: this));
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
