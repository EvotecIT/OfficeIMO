namespace OfficeIMO.Toml;

/// <summary>Model state flags. Editing sets them; saving copies the source of every node without them.</summary>
[Flags]
internal enum TomlNodeState : byte {
    None = 0,
    /// <summary>The node's own value text is re-rendered from the model at its source span.</summary>
    Dirty = 1,
    /// <summary>A descendant changed, or children were added or removed: saving walks into the node.</summary>
    Changed = 2,
    /// <summary>The trailing comment was set and is written from <see cref="TomlNode.ChangedComment"/>.</summary>
    CommentChanged = 4,
    /// <summary>The node was removed from its container.</summary>
    Detached = 8,
    /// <summary>The inline table is complete; nothing outside its braces may add to it (load-time rule).</summary>
    Sealed = 16,
    /// <summary>A loaded implicit or dotted table or array of tables: it has no span of its own but its contents have text.</summary>
    Loaded = 32
}

/// <summary>
/// A node of a <see cref="TomlDocument"/>: a <see cref="TomlTable"/>, <see cref="TomlTableArray"/>,
/// <see cref="TomlArray"/> or <see cref="TomlValue"/>. Loaded nodes are views over the retained source text: they store
/// offsets, not copies, and decode values only when asked.
/// </summary>
/// <remarks>
/// A loaded document can be read from several threads at once; lazily built caches are computed idempotently. Editing
/// is not thread-safe.
/// </remarks>
public abstract class TomlNode {
    /// <summary>
    /// The parent node, or the document for the root and for a node not yet added to a container. One field serves
    /// both, since a loaded document has a node per value and most of them never need the document directly.
    /// </summary>
    private object _owner;
    internal string? _key;
    internal int _start = -1;
    internal int _end = -1;
    /// <summary>Start of the first key segment of the key/value pair that defines this node, or -1.</summary>
    internal int _keyStart = -1;
    internal TomlNodeState _state;

    private protected TomlNode(TomlDocument document) {
        _owner = document;
    }

    /// <summary>Gets or sets the containing node; <see langword="null"/> for the root and nodes not added yet.</summary>
    internal TomlNode? _parent {
        get => _owner as TomlNode;
        set => _owner = value ?? (object)_document;
    }

    /// <summary>Gets the document the node belongs to, through its parents.</summary>
    internal TomlDocument _document {
        get {
            object owner = _owner;
            while (owner is TomlNode node) owner = node._owner;
            return (TomlDocument)owner;
        }
    }

    /// <summary>Gets what kind of node this is.</summary>
    public abstract TomlNodeKind Kind { get; }

    /// <summary>Gets the table, array or array of tables that holds the node, or <see langword="null"/> for the root and removed nodes.</summary>
    public TomlNode? Parent => (_state & TomlNodeState.Detached) != 0 ? null : _parent;

    /// <summary>Gets the node's key in its parent table, or <see langword="null"/> for the root, array items and tables of an array of tables.</summary>
    public string? Key => _key;

    /// <summary>Gets where the node starts in the loaded text: its value, or the <c>[</c> of a header. Nodes added by editing have the default position.</summary>
    public TomlPosition Start => _start < 0 ? default : _document.GetPosition(_start);

    /// <summary>Gets where the node ends (exclusive) in the loaded text. Nodes added by editing have the default position.</summary>
    public TomlPosition End => _end < 0 ? default : _document.GetPosition(_end);

    /// <summary>
    /// Gets where the node's key starts in the loaded text: its segment of the key/value pair or <c>[header]</c> that
    /// defines it (<c>b</c> in <c>a.b.c = 1</c> for table <c>b</c>), or, for a table made only by dotted keys or deeper
    /// headers, its segment of the first such key. Quoted keys include their quotes. The root, array items, tables of an
    /// array of tables and nodes added by editing have the default position.
    /// </summary>
    public TomlPosition KeyStart => TomlKeyLocator.TryFind(this, out int start, out _) ? _document.GetPosition(start) : default;

    /// <summary>Gets where the node's key ends (exclusive) in the loaded text; see <see cref="KeyStart"/>.</summary>
    public TomlPosition KeyEnd => TomlKeyLocator.TryFind(this, out _, out int end) ? _document.GetPosition(end) : default;

    internal bool IsSource => _start >= 0;

    /// <summary>Gets whether the node came from the loaded text, including implicit and dotted tables without a span.</summary>
    internal bool IsLoaded => _start >= 0 || (_state & TomlNodeState.Loaded) != 0;

    /// <summary>Gets whether the node is defined by a <c>key = value</c> pair rather than a header or as an array item.</summary>
    internal bool IsKeyValue => _keyStart >= 0;

    /// <summary>
    /// Gets the end of the last key segment of the key/value pair that defines this node, or -1. Only spaces, tabs and
    /// the <c>=</c> separate it from the value, so it is found from the value's start instead of being stored.
    /// </summary>
    internal int PairKeyEnd {
        get {
            if (_keyStart < 0) return -1;
            string source = _document._source;
            int i = _start;
            while (i > _keyStart && (source[i - 1] == ' ' || source[i - 1] == '\t')) i--;
            if (i > _keyStart && source[i - 1] == '=') i--;
            while (i > _keyStart && (source[i - 1] == ' ' || source[i - 1] == '\t')) i--;
            return i;
        }
    }

    /// <summary>Gets the trailing comment set with <see cref="SetComment"/>; meaningful with <see cref="TomlNodeState.CommentChanged"/>.</summary>
    internal string? ChangedComment => _document.GetChangedComment(this);

    /// <summary>Returns the node as a table.</summary>
    /// <exception cref="InvalidOperationException">The node is not a table.</exception>
    public TomlTable AsTable() => this as TomlTable ?? throw WrongKind("a table");

    /// <summary>Returns the node as an array value.</summary>
    /// <exception cref="InvalidOperationException">The node is not an array.</exception>
    public TomlArray AsArray() => this as TomlArray ?? throw WrongKind("an array");

    /// <summary>Returns the node as an array of tables.</summary>
    /// <exception cref="InvalidOperationException">The node is not an array of tables.</exception>
    public TomlTableArray AsTableArray() => this as TomlTableArray ?? throw WrongKind("an array of tables");

    /// <summary>Returns the node as a string, number, boolean or date-time value.</summary>
    /// <exception cref="InvalidOperationException">The node is a table or array.</exception>
    public TomlValue AsValue() => this as TomlValue ?? throw WrongKind("a value");

    /// <summary>
    /// Returns a value's text: the decoded string of a string, and the TOML spelling of other values as written in the
    /// source (<c>0xFF</c>, <c>1979-05-27</c>, <c>true</c>). Use <see cref="GetValue{T}"/> for typed values.
    /// </summary>
    /// <exception cref="InvalidOperationException">The node is a table or array.</exception>
    public string AsString() => AsValue().GetText();

    /// <summary>
    /// Reads the value as <typeparamref name="T"/>. Supported types are <see cref="string"/> (as <see cref="AsString"/>),
    /// <see cref="bool"/>, <see cref="int"/>, <see cref="long"/>, <see cref="double"/>, <see cref="float"/>,
    /// <see cref="decimal"/>, <see cref="TomlDateTime"/>, <see cref="DateTime"/>, <see cref="DateTimeOffset"/> and
    /// <see cref="TimeSpan"/> (time of day), plus <c>DateOnly</c> and <c>TimeOnly</c> on .NET 8 and later. Integers read
    /// as floating types; a float reads as an integer type only when it is whole and in range.
    /// </summary>
    /// <exception cref="InvalidOperationException">The node is a table or array.</exception>
    /// <exception cref="FormatException">The value cannot be read as <typeparamref name="T"/>.</exception>
    /// <exception cref="NotSupportedException"><typeparamref name="T"/> is not a supported type.</exception>
    public T GetValue<T>() {
        TomlValue value = AsValue();
        if (TomlValueConverter.TryConvert(value, out T result)) return result;
        throw new FormatException("The " + value.Kind + " value '" + value.GetText() + "'" + (IsSource ? " at " + Start : string.Empty) + " cannot be read as " + typeof(T).Name + ".");
    }

    /// <summary>
    /// Tries to read the node as <typeparamref name="T"/> with the conversions of <see cref="GetValue{T}"/>. Returns
    /// <see langword="false"/> for tables, arrays and values that do not convert.
    /// </summary>
    /// <exception cref="NotSupportedException"><typeparamref name="T"/> is not a supported type.</exception>
    public bool TryGetValue<T>(out T value) {
        if (this is TomlValue scalar) return TomlValueConverter.TryConvert(scalar, out value);
        TomlValueConverter.RequireSupported<T>();
        value = default!;
        return false;
    }

    /// <summary>
    /// Returns the node a path addresses from here, or <see langword="null"/> when nothing matches:
    /// <list type="bullet">
    /// <item><description><c>.name</c> selects a table entry by key; a name runs to the next <c>.</c> or <c>[</c>, and the
    /// leading dot of the first step may be omitted (<c>tool.ruff.line-length</c>).</description></item>
    /// <item><description><c>["key"]</c> selects a key containing dots, brackets or other characters, with JSON escapes
    /// (<c>\"</c>, <c>\\</c>, <c>\n</c>, <c>\uXXXX</c>…).</description></item>
    /// <item><description><c>[n]</c> selects the zero-based item of an array or table of an array of tables
    /// (<c>package[3].name</c>).</description></item>
    /// <item><description>A leading <c>$</c> stands for this node.</description></item>
    /// </list>
    /// The grammar is the one <c>OfficeIMO.Reader.Toml</c> reports paths in, and the same as YAML node paths.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is malformed.</exception>
    public TomlNode? Select(string path) => TomlPath.Select(this, path);

    /// <summary>
    /// Gets the comments on the lines directly above the node (no blank line between), in source order, as the text after
    /// <c>#</c>. Nodes added by editing have none.
    /// </summary>
    public IReadOnlyList<string> LeadingComments => TomlTrivia.GetLeadingComments(this);

    /// <summary>
    /// Gets the comment at the end of the line where the node ends, as the text after <c>#</c>, or <see langword="null"/>.
    /// For a key/value pair or a table header that is the comment after it; a comment set with
    /// <see cref="SetComment"/> is returned before the document is saved.
    /// </summary>
    public string? TrailingComment => (_state & TomlNodeState.CommentChanged) != 0 ? ChangedComment : TomlTrivia.GetTrailingComment(this);

    /// <summary>
    /// Sets or, with <see langword="null"/>, removes the comment at the end of the node's line and returns the node.
    /// Works for key/value pairs, table headers and tables of an array of tables; on save only that comment's text changes.
    /// </summary>
    /// <param name="comment">The comment text without <c>#</c>. It cannot contain line breaks.</param>
    /// <exception cref="ArgumentException"><paramref name="comment"/> contains a line break or a control character.</exception>
    /// <exception cref="InvalidOperationException">The node is an array item, the root, an implicit or dotted table, or an array of tables.</exception>
    public TomlNode SetComment(string? comment) {
        if (comment != null) {
            for (int i = 0; i < comment.Length; i++) {
                if (TomlScalars.IsForbiddenControl(comment[i])) throw new ArgumentException("A comment cannot contain line breaks or control characters.", nameof(comment));
            }
        }
        if (!CanHaveComment()) throw new InvalidOperationException("Only key/value pairs and table headers have trailing comments.");
        _document.SetChangedComment(this, comment);
        _state |= TomlNodeState.CommentChanged;
        TomlEditing.MarkChanged(this);
        return this;
    }

    internal bool CanHaveComment() {
        if (this is TomlTable table && table._tableKind is TomlTableKind.Header or TomlTableKind.ArrayItem) return true;
        bool valueNode = this is TomlValue || this is TomlArray || (this is TomlTable inline && inline._tableKind == TomlTableKind.Inline);
        return valueNode && _key != null && _parent is TomlTable && !InInlineContext();
    }

    /// <summary>Returns whether the node sits inside an inline table or array value.</summary>
    internal bool InInlineContext() {
        for (TomlNode? node = _parent; node != null; node = node._parent) {
            if (node is TomlArray || (node is TomlTable table && table._tableKind == TomlTableKind.Inline)) return true;
        }
        return false;
    }

    /// <summary>Returns the node as TOML text: its source text when unchanged, otherwise the text saving would write for it.</summary>
    public override string ToString() => TomlRendering.NodeText(this);

    internal InvalidOperationException WrongKind(string expected) =>
        new InvalidOperationException("Expected " + expected + " but the node" + (IsSource ? " at " + Start : string.Empty) + " is " + Describe(Kind) + ".");

    internal static string Describe(TomlNodeKind kind) => kind switch {
        TomlNodeKind.Table => "a table",
        TomlNodeKind.TableArray => "an array of tables",
        TomlNodeKind.Array => "an array",
        TomlNodeKind.Integer => "an integer",
        _ => "a " + kind.ToString().ToLowerInvariant()
    };
}
