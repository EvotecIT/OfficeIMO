namespace OfficeIMO.Yaml;

/// <summary>Model state flags. Editing sets them; saving copies the source of every node without dirty flags.</summary>
[Flags]
internal enum YamlNodeState : ushort {
    None = 0,
    /// <summary>The node's text is emitted in place of its region (its own source span, or the span of the node it replaced).</summary>
    Dirty = 1,
    /// <summary>A descendant has an edit: saving walks into the node instead of copying its span whole.</summary>
    ContainsDirty = 2,
    /// <summary>The collection is still being loaded; an alias to it would make the graph cyclic.</summary>
    Open = 4,
    /// <summary>At least one alias refers to the node.</summary>
    Aliased = 8,
    /// <summary>The node's content comes from the model, not the source: created, re-valued or moved by editing.</summary>
    New = 16,
    /// <summary>The anchor or tag changed on a node whose content is still copied from the source.</summary>
    PropertiesDirty = 32,
    /// <summary>A comment or blank line owned by the node changed.</summary>
    CommentsDirty = 64,
    /// <summary>Entries or items were inserted or removed; the collection keeps a snapshot of its loaded children.</summary>
    Restructured = 128,
    /// <summary>The node took the region of a block collection, so it starts on the key's line rather than below it.</summary>
    ReplacedBlock = 256
}

/// <summary>
/// Anchor, tag, alias-expansion size and comment edits of a node. Most nodes have none of these, so they share no object.
/// </summary>
internal sealed class YamlNodeProperties {
    public string? Anchor;
    public string? Tag;
    /// <summary>Nodes the anchored node stands for when every alias inside it is resolved, saturating.</summary>
    public long ExpandedSize;
    public bool CommentBeforeSet;
    public string? CommentBefore;
    public bool CommentSet;
    public string? Comment;
    public bool SpaceBeforeSet;
    public bool SpaceBefore;
}

/// <summary>
/// A node of a <see cref="YamlDocument"/>: a <see cref="YamlScalar"/>, <see cref="YamlMapping"/>,
/// <see cref="YamlSequence"/> or <see cref="YamlAlias"/>. Loaded nodes are views over the retained source text: they
/// store offsets, not copies, and decode values only when asked. Nodes created by editing hold their values directly.
/// </summary>
/// <remarks>
/// <para>Typed reads (<see cref="AsString"/>, <see cref="GetValue{T}"/>, <see cref="AsMapping"/> and the others) look
/// through aliases to the anchored node; the tree itself keeps every alias as a <see cref="YamlAlias"/>. Reading never
/// changes the source text.</para>
/// <para>Editing changes only what it touches. Saving copies every untouched node from the source and writes only the
/// edited ones, so comments, quoting, indentation and spellings such as <c>TRUE</c> survive everywhere else. Methods
/// that configure a node return it, so calls chain: <c>mapping.Add("image", "nginx").SetComment("pinned by CI")</c>.</para>
/// <para>A loaded document can be read from several threads at once. Lazily built caches (decoded values, key indexes,
/// positions) are computed idempotently. Editing is not thread-safe.</para>
/// </remarks>
public abstract class YamlNode {
    internal YamlDocumentPart? _part;
    internal YamlNode? _parent;
    internal YamlNodeProperties? _properties;
    internal int _start;
    internal int _end;
    /// <summary>The node's position in its parent, a hint that makes finding its entry constant time.</summary>
    internal int _index;
    internal YamlStyle _style;
    internal YamlNodeState _state;

    private protected YamlNode(YamlDocumentPart part, YamlStyle style, int start, int end) {
        _part = part;
        _style = style;
        _start = start;
        _end = end;
    }

    /// <summary>Creates a node that belongs to no document yet.</summary>
    private protected YamlNode(YamlStyle style) {
        _style = style;
        _start = -1;
        _end = -1;
        _state = YamlNodeState.New;
    }

    /// <summary>
    /// Gets the presentation style: <see cref="YamlStyle.Plain"/>, a quoted or block scalar style, or
    /// <see cref="YamlStyle.Block"/>/<see cref="YamlStyle.Flow"/> for collections. Aliases report <see cref="YamlStyle.None"/>,
    /// and so does a new string scalar until it joins a document, which picks its style from its context.
    /// </summary>
    public YamlStyle Style => _style;

    /// <summary>
    /// Gets the node's resolved tag (<c>!!str</c> becomes <c>tag:yaml.org,2002:str</c>, <c>!Ref</c> stays <c>!Ref</c>),
    /// or <see langword="null"/> when it has none. Custom tags are kept verbatim and never fail a load.
    /// </summary>
    public string? Tag => _properties?.Tag;

    /// <summary>Gets the anchor name (without <c>&amp;</c>) defined on this node, or <see langword="null"/>.</summary>
    public string? Anchor => _properties?.Anchor;

    /// <summary>Gets the mapping or sequence that holds the node, or <see langword="null"/> for a document root or a node not in a document.</summary>
    public YamlNode? Parent => _parent;

    /// <summary>
    /// Gets where the node starts in the loaded source, with its offset and one-based line and column: its first property
    /// (anchor or tag) when present, otherwise its content. Nodes whose content comes from editing (created, re-valued
    /// or moved) have no source position and return the default value, whose <see cref="YamlPosition.Line"/> is 0.
    /// </summary>
    /// <remarks>
    /// Positions refer to the text as loaded; edits elsewhere never shift them. The first position read builds the
    /// document's line table once; nodes store only offsets. <see cref="IsModified"/> tells whether the node's text in
    /// a save would still be the text at this position.
    /// </remarks>
    public YamlPosition Start => HasSourceSpan ? _part!._document.GetPosition(_start) : default;

    /// <summary>
    /// Gets where the node ends in the loaded source. A block collection ends where its last content ends; an empty node
    /// is zero-width at the point where the omitted node would be. Nodes created by editing return the default value.
    /// </summary>
    public YamlPosition End => HasSourceSpan ? _part!._document.GetPosition(_end) : default;

    /// <summary>
    /// Gets whether editing changed the node or anything inside it since the document was loaded: its value, anchor,
    /// tag or comments, or a descendant's, or entries or items added or removed. It is <see langword="true"/> for nodes
    /// created by editing, and <see langword="false"/> for an untouched loaded node, whose source text between
    /// <see cref="Start"/> and <see cref="End"/> is what saving writes for it.
    /// </summary>
    public bool IsModified => !HasSourceSpan || (_state & (YamlNodeState.Dirty | YamlNodeState.ContainsDirty | YamlNodeState.PropertiesDirty
        | YamlNodeState.CommentsDirty | YamlNodeState.Restructured)) != 0;

    /// <summary>
    /// Gets whether the node resolves to null under its document's schema: an untagged plain scalar such as <c>null</c>,
    /// <c>~</c> or an empty node (core schema), or a scalar tagged <c>!!null</c>. Aliases answer for their anchored node;
    /// collections and quoted scalars are never null.
    /// </summary>
    public bool IsNull => Resolved is YamlScalar scalar && scalar.ResolvesToNull();

    /// <summary>Gets the node itself, or the anchored node for an alias. The anchored node is never itself an alias.</summary>
    internal YamlNode Resolved => this is YamlAlias alias ? alias.Target : this;

    internal YamlDocument? Document => _part?._document;

    /// <summary>The schema typed reads use: the document part's, or the core schema for a node not in a document.</summary>
    internal YamlSchema Schema => _part?.Schema ?? YamlSchema.Core;

    /// <summary>Whether the node's content is still the loaded source text at <see cref="_start"/>.</summary>
    internal bool HasSourceSpan => (_state & YamlNodeState.New) == 0 && _part is not null && _start >= 0;

    /// <summary>Whether the node is a block collection without anchor or tag, which starts where its first child starts.</summary>
    internal bool IsTransparentBlock => _style == YamlStyle.Block && this is YamlMapping or YamlSequence && _properties?.Anchor is null && _properties?.Tag is null;

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
        throw new FormatException("The scalar '" + YamlValueConverter.Describe(scalar) + "'" + scalar.Where()
            + " is not a " + typeof(T).Name + " under the " + scalar.Schema + " schema.");
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
    /// (<c>\"</c>, <c>\\</c>, <c>\n</c>, <c>\uXXXX</c>...).</description></item>
    /// <item><description><c>[n]</c> selects the zero-based item of a sequence.</description></item>
    /// <item><description>A leading <c>$</c> stands for this node.</description></item>
    /// </list>
    /// Key steps match scalar keys by decoded text and follow aliases and, under the YAML 1.1 schema, <c>&lt;&lt;</c>
    /// merge keys. Paths address nodes; they do not evaluate expressions. Because steps match text, a path cannot reach
    /// values under collection keys (<c>[a, b]: x</c>), and of two keys with the same text but different types
    /// (<c>1</c> and <c>"1"</c>) it finds the one the document's duplicate-key policy picks; enumerate
    /// <see cref="YamlMapping"/> entries to reach those. <see cref="YamlDocument.Set"/>,
    /// <see cref="YamlDocument.Add"/> and <see cref="YamlDocument.Remove"/> edit by the same paths.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is malformed.</exception>
    public YamlNode? Select(string path) => YamlPath.Select(this, path, 0);

    /// <summary>Returns the path of this node from its document root, in the grammar <see cref="Select"/> reads.</summary>
    /// <remarks>Keys that are not plain names (dots, brackets, quotes) use the <c>["key"]</c> form.</remarks>
    /// <exception cref="InvalidOperationException">The node is a mapping key, or sits under a collection key; see <see cref="Select"/>.</exception>
    public string GetPath() => YamlPath.Format(this);

    /// <summary>
    /// Gets the comment lines above the node's entry, joined with line breaks, without their <c>#</c> and one following
    /// space; <see langword="null"/> when there are none.
    /// </summary>
    /// <remarks>
    /// <para>The comments belong to the line where the node's entry starts. A mapping value shares them with its key, and
    /// the first entry of a block collection that is a sequence item or the document root shares them with that
    /// collection; reading or setting them through any of these nodes reaches the same lines.</para>
    /// <para>They are the own-line comments directly above the entry, with no blank line between. Comments that a blank
    /// line separates from the entry belong to no entry, such as a commented-out example, and stay when the entry is
    /// removed; above the first entry of a document they are the document's header
    /// (<see cref="YamlDocumentPart.CommentBefore"/>). Comments inside flow collections are kept in the source but not
    /// reported.</para>
    /// </remarks>
    public string? CommentBefore => YamlTrivia.GetCommentBefore(this);

    /// <summary>
    /// Gets the comment at the end of the node's line, without its <c>#</c> and one following space, or
    /// <see langword="null"/>. A block scalar's comment is the one on its header line (<c>key: | # comment</c>); a block
    /// collection that is a mapping value reports its key's line (<c>key: # comment</c>). Other block collections have no
    /// line of their own: the comment after their last line belongs to the innermost node ending there. A key whose value
    /// starts on the same line has none; the comment after that value is the value's.
    /// </summary>
    public string? Comment => YamlTrivia.GetComment(this);

    /// <summary>
    /// Gets whether a blank line separates the node's entry from the previous content, above any
    /// <see cref="CommentBefore"/> lines. Mapping values and first entries share this with the same nodes as
    /// <see cref="CommentBefore"/>.
    /// </summary>
    public bool SpaceBefore => YamlTrivia.GetSpaceBefore(this);

    /// <summary>
    /// Replaces the comment lines above the node's entry; each line of <paramref name="text"/> becomes one <c># line</c>
    /// comment at the entry's indentation. <see langword="null"/> or an empty string removes them.
    /// </summary>
    /// <returns>This node.</returns>
    /// <exception cref="InvalidOperationException">The node is inside a flow collection, where comment lines cannot be placed.</exception>
    public YamlNode SetCommentBefore(string? text) {
        YamlEdit.SetCommentBefore(this, string.IsNullOrEmpty(text) ? null : text);
        return this;
    }

    /// <summary>
    /// Sets the comment at the end of the node's line (<c>value # text</c>), or removes it with <see langword="null"/>.
    /// </summary>
    /// <returns>This node.</returns>
    /// <exception cref="ArgumentException">The text contains a line break.</exception>
    /// <exception cref="InvalidOperationException">The node has no line of its own: a block collection that is a sequence item or
    /// the root, a key whose value starts on its line, or a node inside a flow collection.</exception>
    public YamlNode SetComment(string? text) {
        if (text is not null && (text.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0)) throw new ArgumentException("A trailing comment is one line.", nameof(text));
        YamlEdit.SetComment(this, text);
        return this;
    }

    /// <summary>Sets whether a blank line separates the node's entry from the previous content.</summary>
    /// <returns>This node.</returns>
    /// <exception cref="InvalidOperationException">The node is inside a flow collection.</exception>
    public YamlNode SetSpaceBefore(bool value) {
        YamlEdit.SetSpaceBefore(this, value);
        return this;
    }

    /// <summary>
    /// Sets, renames or removes (<see langword="null"/>) the node's anchor. Renaming updates every alias that refers to
    /// the node; removing an anchor that aliases still use fails.
    /// </summary>
    /// <returns>This node.</returns>
    /// <exception cref="ArgumentException">The name is not a valid anchor, or another node of the document already uses it.</exception>
    /// <exception cref="InvalidOperationException">The node is an alias, or aliases refer to the anchor being removed.</exception>
    public YamlNode SetAnchor(string? name) {
        YamlEdit.SetAnchor(this, name);
        return this;
    }

    /// <summary>
    /// Sets or removes (<see langword="null"/>) the node's tag, in the resolved form <see cref="Tag"/> reports
    /// (<c>tag:yaml.org,2002:str</c>, <c>!Ref</c>). Tags change how a scalar resolves; they never activate .NET types.
    /// </summary>
    /// <returns>This node.</returns>
    /// <exception cref="InvalidOperationException">The node is an alias.</exception>
    public YamlNode SetTag(string? tag) {
        YamlEdit.SetTag(this, string.IsNullOrEmpty(tag) ? null : tag);
        return this;
    }

    /// <summary>
    /// Enumerates this node and every node below it, depth first in document order: for a mapping, each key and then its
    /// value. Aliases are returned as themselves and not followed. The walk keeps one small stack and allocates nothing
    /// per node.
    /// </summary>
    public YamlNodeWalk Descendants() => new YamlNodeWalk(this);

    /// <summary>
    /// Returns the node's YAML text: its source text for a loaded, unchanged node, otherwise the text saving would write
    /// for it on its own.
    /// </summary>
    public override string ToString() {
        if (HasSourceSpan && (_state & (YamlNodeState.Dirty | YamlNodeState.ContainsDirty | YamlNodeState.PropertiesDirty)) == 0) {
            return _part!._document._source.Substring(_start, _end - _start);
        }
        return YamlModelWriter.Render(this);
    }

    /// <summary>Creates a string scalar; its style is chosen from its context when it joins a document.</summary>
    public static implicit operator YamlNode(string? value) => new YamlScalar(value);

    /// <summary>Creates an integer scalar.</summary>
    public static implicit operator YamlNode(long value) => new YamlScalar(value);

    /// <summary>Creates a floating-point scalar.</summary>
    public static implicit operator YamlNode(double value) => new YamlScalar(value);

    /// <summary>Creates a boolean scalar.</summary>
    public static implicit operator YamlNode(bool value) => new YamlScalar(value);

    /// <summary>Formats " at line:column" for messages, or nothing for a node without a source position.</summary>
    internal string Where() => HasSourceSpan ? " at " + Start : string.Empty;

    private InvalidOperationException WrongKind(string expected) {
        YamlNode resolved = Resolved;
        string actual = resolved switch {
            YamlMapping => "a mapping",
            YamlSequence => "a sequence",
            _ => "a scalar"
        };
        string via = ReferenceEquals(resolved, this) ? string.Empty : " (through alias *" + ((YamlAlias)this).Name + ")";
        return new InvalidOperationException("Expected " + expected + " but the node" + Where() + " is " + actual + via + ".");
    }
}
