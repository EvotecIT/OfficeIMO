namespace OfficeIMO.Yaml;

/// <summary>
/// One document of a YAML stream: its directives, its <c>---</c> and <c>...</c> markers, its root node and the
/// comments around it. A file with several documents (a Kubernetes manifest list, for example) has one part per document.
/// </summary>
public sealed class YamlDocumentPart {
    internal readonly YamlDocument _document;
    internal YamlNode? _root;
    /// <summary>Start of the part: its first directive, its <c>---</c>, or its first content when both are absent.</summary>
    internal int _start;
    /// <summary>End of the <c>---</c> marker, or <see cref="_start"/> without one.</summary>
    internal int _markerEnd;
    internal int _end;
    /// <summary>Whether the part was added by editing and has no source text; saving writes it whole.</summary>
    internal readonly bool _new;
    internal bool _commentBeforeSet;
    internal string? _commentBefore;
    internal bool _commentAfterSet;
    internal string? _commentAfter;
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

    /// <summary>Creates a part added by editing.</summary>
    internal YamlDocumentPart(YamlDocument document, YamlSchema schema, bool hasStartMarker) {
        _document = document;
        Schema = schema;
        _tagDirectives = Array.Empty<YamlTagDirective>();
        HasStartMarker = hasStartMarker;
        _new = true;
        _start = _markerEnd = _end = -1;
    }

    /// <summary>
    /// Gets the root node, or sets a new one in its place. A document without content has an empty plain scalar root,
    /// which reads as null. Setting writes the new root in place of the old one and leaves the comments around it.
    /// </summary>
    /// <exception cref="InvalidOperationException">Setting: the node already belongs to a document, or aliases refer to the old root.</exception>
    public YamlNode Root {
        get => _root!;
        set => YamlEdit.ReplaceRoot(this, value ?? new YamlScalar((string?)null));
    }

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

    /// <summary>
    /// Gets the document's header comments, such as a license, joined with line breaks and without their <c>#</c> and one
    /// following space; <see langword="null"/> when there are none. They are the comment lines before <c>---</c>, or, in a
    /// document without one, the comment lines at its top that a blank line separates from the first entry.
    /// </summary>
    public string? CommentBefore => YamlTrivia.PartComments(this, header: true);

    /// <summary>Gets the comment lines after the root's last line, joined with line breaks; <see langword="null"/> when there are none.</summary>
    public string? CommentAfter => YamlTrivia.PartComments(this, header: false);

    /// <summary>
    /// Replaces the header comments; each line of <paramref name="text"/> becomes one <c># line</c> comment. In a document
    /// without <c>---</c> a blank line follows them so they stay the document's. <see langword="null"/> removes them.
    /// </summary>
    /// <returns>This part.</returns>
    public YamlDocumentPart SetCommentBefore(string? text) {
        _commentBeforeSet = true;
        _commentBefore = string.IsNullOrEmpty(text) ? null : text;
        _document._partsDirty = true;
        return this;
    }

    /// <summary>Replaces the comment lines after the root's last line; <see langword="null"/> removes them.</summary>
    /// <returns>This part.</returns>
    public YamlDocumentPart SetCommentAfter(string? text) {
        _commentAfterSet = true;
        _commentAfter = string.IsNullOrEmpty(text) ? null : text;
        _document._partsDirty = true;
        return this;
    }

    /// <summary>Whether a <c>%TAG</c> directive redefines <c>!!</c>, so standard tags must be written verbatim.</summary>
    internal bool RedefinesSecondaryHandle {
        get {
            foreach (YamlTagDirective directive in _tagDirectives) {
                if (directive.Handle == "!!") return true;
            }
            return false;
        }
    }
}
