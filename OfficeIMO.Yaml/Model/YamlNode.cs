namespace OfficeIMO.Yaml;

/// <summary>Model state flags. Editing sets the dirty flags; saving copies the source of every node without them.</summary>
[Flags]
internal enum YamlNodeState : byte {
    None = 0,
    /// <summary>The node's own text changed: saving emits it instead of copying its source span.</summary>
    Dirty = 1,
    /// <summary>A descendant is dirty: saving walks into the node instead of copying its span whole.</summary>
    ContainsDirty = 2,
    /// <summary>The collection is still being loaded; an alias to it would make the graph cyclic.</summary>
    Open = 4
}

/// <summary>Anchor, tag and alias-expansion size of a node that carries properties. Most nodes have none, so they share no object.</summary>
internal sealed class YamlNodeProperties {
    public string? Anchor;
    public string? Tag;
    /// <summary>Nodes the anchored node stands for when every alias inside it is resolved, saturating.</summary>
    public long ExpandedSize;
}

/// <summary>
/// A node of a loaded <see cref="YamlDocument"/>: a <see cref="YamlScalar"/>, <see cref="YamlMapping"/>,
/// <see cref="YamlSequence"/> or <see cref="YamlAlias"/>. Nodes are views over the retained source text: they store
/// offsets, not copies, and decode values only when asked.
/// </summary>
/// <remarks>
/// <para>Typed reads (<see cref="AsString"/>, <see cref="GetValue{T}"/>, <see cref="AsMapping"/> and the others) look
/// through aliases to the anchored node; the tree itself keeps every alias as a <see cref="YamlAlias"/>. Reading never
/// changes the source text.</para>
/// <para>A loaded document can be read from several threads at once. Lazily built caches (decoded values, key indexes,
/// positions) are computed idempotently.</para>
/// </remarks>
public abstract class YamlNode {
    internal readonly YamlDocumentPart _part;
    internal YamlNode? _parent;
    internal YamlNodeProperties? _properties;
    internal int _start;
    internal int _end;
    internal YamlStyle _style;
    internal YamlNodeState _state;

    private protected YamlNode(YamlDocumentPart part, YamlStyle style, int start, int end) {
        _part = part;
        _style = style;
        _start = start;
        _end = end;
    }

    /// <summary>
    /// Gets the presentation style: <see cref="YamlStyle.Plain"/>, a quoted or block scalar style, or
    /// <see cref="YamlStyle.Block"/>/<see cref="YamlStyle.Flow"/> for collections. Aliases report <see cref="YamlStyle.None"/>.
    /// </summary>
    public YamlStyle Style => _style;

    /// <summary>
    /// Gets the node's resolved tag (<c>!!str</c> becomes <c>tag:yaml.org,2002:str</c>, <c>!Ref</c> stays <c>!Ref</c>),
    /// or <see langword="null"/> when the source has none. Custom tags are kept verbatim and never fail a load.
    /// </summary>
    public string? Tag => _properties?.Tag;

    /// <summary>Gets the anchor name (without <c>&amp;</c>) defined on this node, or <see langword="null"/>.</summary>
    public string? Anchor => _properties?.Anchor;

    /// <summary>Gets where the node starts: its first property (anchor or tag) when present, otherwise its content.</summary>
    public YamlPosition Start => _part._document.GetPosition(_start);

    /// <summary>
    /// Gets where the node ends. A block collection ends where its last content ends; an empty node is zero-width at the
    /// point where the omitted node would be.
    /// </summary>
    public YamlPosition End => _part._document.GetPosition(_end);

    /// <summary>
    /// Gets whether the node resolves to null under its document's schema: an untagged plain scalar such as <c>null</c>,
    /// <c>~</c> or an empty node (core schema), or a scalar tagged <c>!!null</c>. Aliases answer for their anchored node;
    /// collections and quoted scalars are never null.
    /// </summary>
    public bool IsNull => Resolved is YamlScalar scalar && scalar.ResolvesToNull();

    /// <summary>Gets the node itself, or the anchored node for an alias. The anchored node is never itself an alias.</summary>
    internal YamlNode Resolved => this is YamlAlias alias ? alias.Target : this;

    internal YamlDocument Document => _part._document;

    /// <summary>Returns the scalar this node is, or the scalar an alias refers to.</summary>
    /// <exception cref="InvalidOperationException">The node is a mapping or sequence.</exception>
    public YamlScalar AsScalar() => Resolved as YamlScalar ?? throw WrongKind("a scalar");

    /// <summary>Returns the mapping this node is, or the mapping an alias refers to.</summary>
    /// <exception cref="InvalidOperationException">The node is a scalar or sequence.</exception>
    public YamlMapping AsMapping() => Resolved as YamlMapping ?? throw WrongKind("a mapping");

    /// <summary>Returns the sequence this node is, or the sequence an alias refers to.</summary>
    /// <exception cref="InvalidOperationException">The node is a scalar or mapping.</exception>
    public YamlSequence AsSequence() => Resolved as YamlSequence ?? throw WrongKind("a sequence");

    /// <summary>
    /// Returns the scalar's decoded text (folding, escapes and block chomping applied) whatever it resolves to, so
    /// <c>~</c> returns <c>"~"</c> and <c>0x1F</c> returns <c>"0x1F"</c>. Use <see cref="GetValue{T}"/> for typed values.
    /// </summary>
    /// <exception cref="InvalidOperationException">The node is a mapping or sequence.</exception>
    public string AsString() => AsScalar().Value;

    /// <summary>
    /// Reads the scalar as <typeparamref name="T"/> through its document's schema. Supported types are
    /// <see cref="string"/>, <see cref="bool"/>, <see cref="int"/>, <see cref="long"/>, <see cref="double"/>,
    /// <see cref="float"/>, <see cref="decimal"/> and their nullable forms. A null scalar reads as <see langword="null"/>
    /// for <see cref="string"/> and nullable types. A <see cref="string"/> read returns the text of any non-null scalar,
    /// including custom-tagged ones.
    /// </summary>
    /// <exception cref="InvalidOperationException">The node is a mapping or sequence.</exception>
    /// <exception cref="FormatException">The scalar does not resolve to <typeparamref name="T"/>.</exception>
    /// <exception cref="NotSupportedException"><typeparamref name="T"/> is not a supported type.</exception>
    public T GetValue<T>() {
        YamlScalar scalar = AsScalar();
        if (YamlValueConverter.TryConvert(scalar, out T value)) return value;
        throw new FormatException("The scalar '" + YamlValueConverter.Describe(scalar) + "' at " + scalar.Start
            + " is not a " + typeof(T).Name + " under the " + scalar._part.Schema + " schema.");
    }

    /// <summary>
    /// Tries to read the node as <typeparamref name="T"/> through its document's schema, with the conversions of
    /// <see cref="GetValue{T}"/>. Returns <see langword="false"/> for collections and for scalars that do not resolve.
    /// </summary>
    /// <exception cref="NotSupportedException"><typeparamref name="T"/> is not a supported type.</exception>
    public bool TryGetValue<T>(out T value) {
        if (Resolved is YamlScalar scalar) return YamlValueConverter.TryConvert(scalar, out value);
        YamlValueConverter.RequireSupported<T>();
        value = default!;
        return false;
    }

    /// <summary>
    /// Returns the node a path addresses from here, or <see langword="null"/> when nothing matches. Paths share their
    /// grammar with the paths <c>OfficeIMO.Reader.Yaml</c> reports:
    /// <list type="bullet">
    /// <item><description><c>.name</c> selects a mapping value by key text; a name runs to the next <c>.</c> or <c>[</c>,
    /// and the leading dot of the first step may be omitted (<c>spec.replicas</c>).</description></item>
    /// <item><description><c>["key"]</c> selects a key containing dots, brackets or other characters, with JSON escapes
    /// (<c>\"</c>, <c>\\</c>, <c>\n</c>, <c>\uXXXX</c>…).</description></item>
    /// <item><description><c>[n]</c> selects the zero-based item of a sequence.</description></item>
    /// <item><description>A leading <c>$</c> stands for this node.</description></item>
    /// </list>
    /// Key steps match scalar keys by decoded text and follow aliases and, under the YAML 1.1 schema, <c>&lt;&lt;</c>
    /// merge keys. Paths address nodes; they do not evaluate expressions.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is malformed.</exception>
    public YamlNode? Select(string path) => YamlPath.Select(this, path, 0);

    /// <summary>
    /// Gets the comments that precede the node on lines of their own, in source order, as the text after <c>#</c>.
    /// </summary>
    /// <remarks>
    /// A comment belongs to the innermost node that starts after it, so a comment above <c>key: value</c> belongs to the
    /// key, and a comment above a block collection without an anchor or tag belongs to the collection's first key or item.
    /// The comments read are those between the node and the content before it (the previous sibling, its key, the opening
    /// bracket or properties of its collection, or the document start), excluding the comment on the line where that
    /// previous node ends, which is the previous node's <see cref="TrailingComment"/>. Comments before a <c>---</c> marker
    /// and after the last node of a document belong to no node; they are still kept in the source and saved.
    /// </remarks>
    public IReadOnlyList<string> LeadingComments => YamlTrivia.GetLeadingComments(this);

    /// <summary>
    /// Gets the comment on the line where the node ends, as the text after <c>#</c>, or <see langword="null"/>. For a block
    /// scalar it is the comment on the header line (<c>key: | # comment</c>). A block collection has none of its own: the
    /// comment after its last line belongs to the innermost node ending there. A key whose empty value ends before the
    /// comment (<c>key: # comment</c>) leaves the comment to that value.
    /// </summary>
    public string? TrailingComment => YamlTrivia.GetTrailingComment(this);

    /// <summary>Returns the node's source text, from <see cref="Start"/> to <see cref="End"/>.</summary>
    public override string ToString() => Document._source.Substring(_start, _end - _start);

    /// <summary>Writes the replacement text of a dirty node. Nothing marks nodes dirty until editing exists.</summary>
    internal virtual void WriteChanged(List<YamlTextSegment> output) =>
        throw new NotSupportedException("This node has no emitter for changed text.");

    private InvalidOperationException WrongKind(string expected) {
        YamlNode resolved = Resolved;
        string actual = resolved switch {
            YamlMapping => "a mapping",
            YamlSequence => "a sequence",
            _ => "a scalar"
        };
        string via = ReferenceEquals(resolved, this) ? string.Empty : " (through alias *" + ((YamlAlias)this).Name + ")";
        return new InvalidOperationException("Expected " + expected + " but the node at " + Start + " is " + actual + via + ".");
    }
}
