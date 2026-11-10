namespace OfficeIMO.Yaml;

/// <summary>
/// One document of a YAML stream: its directives, its <c>---</c> and <c>...</c> markers and its root node. A file with
/// several documents (a Kubernetes manifest list, for example) has one part per document.
/// </summary>
public sealed class YamlDocumentPart {
    internal readonly YamlDocument _document;
    internal YamlNode? _root;
    /// <summary>Start of the part: its first directive, its <c>---</c>, or its first content when both are absent.</summary>
    internal int _start;
    /// <summary>End of the <c>---</c> marker, or <see cref="_start"/> without one.</summary>
    internal int _markerEnd;
    internal int _end;
    private readonly YamlTagDirective[] _tagDirectives;

    internal YamlDocumentPart(YamlDocument document, YamlSchema schema, string? version, YamlTagDirective[] tagDirectives, bool hasStartMarker, int start, int markerEnd) {
        _document = document;
        Schema = schema;
        Version = version;
        _tagDirectives = tagDirectives;
        HasStartMarker = hasStartMarker;
        _start = start;
        _markerEnd = markerEnd;
    }

    /// <summary>Gets the root node. A document without content has an empty plain scalar root, which reads as null.</summary>
    public YamlNode Root => _root!;

    /// <summary>
    /// Gets the schema the typed reads of this document use: the schema forced through
    /// <see cref="YamlLoadOptions.Schema"/>, otherwise <see cref="YamlSchema.Yaml11"/> for a <c>%YAML 1.1</c> document and
    /// <see cref="YamlSchema.Core"/> for every other document.
    /// </summary>
    public YamlSchema Schema { get; }

    /// <summary>Gets the version a <c>%YAML</c> directive declares, such as <c>1.1</c>, or <see langword="null"/> without one.</summary>
    public string? Version { get; }

    /// <summary>Gets the document's <c>%TAG</c> directives in source order.</summary>
    public IReadOnlyList<YamlTagDirective> TagDirectives => _tagDirectives;

    /// <summary>Gets whether the document starts with an explicit <c>---</c> marker.</summary>
    public bool HasStartMarker { get; }

    /// <summary>Gets whether the document ends with an explicit <c>...</c> marker.</summary>
    public bool HasEndMarker { get; internal set; }
}
