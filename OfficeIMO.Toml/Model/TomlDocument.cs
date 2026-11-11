namespace OfficeIMO.Toml;

/// <summary>
/// A TOML file loaded into a lossless, editable model. Every key, value, comment, blank line and spelling of the input
/// is kept: saving an unchanged document reproduces the input byte for byte, including a UTF-8 byte-order mark, and
/// saving an edited document changes only the text of what was edited.
/// </summary>
/// <remarks>
/// <para>The document retains its decoded source text once; loaded nodes are offsets into it, so a node costs about one
/// object and comments and whitespace cost nothing until asked for. Values are decoded lazily and cached.</para>
/// <para>Loading checks TOML 1.1 completely, including duplicate keys and the rules for defining and extending tables,
/// and enforces the limits in <see cref="TomlLoadOptions"/>. Invalid input throws <see cref="TomlException"/>.</para>
/// </remarks>
public sealed partial class TomlDocument : IDisposable {
    internal readonly string _source;
    internal readonly bool _byteOrderMark;
    internal readonly TomlTable _root;
    /// <summary>Tables with a <c>[header]</c> or <c>[[header]]</c> line, in source order.</summary>
    internal readonly List<TomlTable> _headers = new List<TomlTable>();
    private int[]? _lineStarts;
    private TomlLayout? _layout;

    private TomlDocument(string source, bool byteOrderMark, string? filePath) {
        _source = source;
        _byteOrderMark = byteOrderMark;
        FilePath = filePath;
        _root = new TomlTable(this, TomlTableKind.Root) { _state = TomlNodeState.Loaded };
    }

    /// <summary>Creates an empty document in memory. Save it with <see cref="Save(string)"/> or <see cref="Save(Stream)"/>.</summary>
    public static TomlDocument Create() => new TomlDocument(string.Empty, byteOrderMark: false, filePath: null);

    /// <summary>
    /// Creates an empty document that <see cref="Save()"/> writes to <paramref name="path"/>. Nothing is written until
    /// it is saved.
    /// </summary>
    /// <param name="path">The file the document saves to.</param>
    public static TomlDocument Create(string path) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return new TomlDocument(string.Empty, byteOrderMark: false, path);
    }

    /// <summary>Gets the file the document was loaded from or created for, which <see cref="Save()"/> writes; <see langword="null"/> for parsed or stream documents.</summary>
    public string? FilePath { get; private set; }

    /// <summary>Gets the root table: the key/value pairs before the first header, and every top-level table.</summary>
    public TomlTable Root => _root;

    /// <summary>
    /// Returns the node a path addresses from the root, or <see langword="null"/>. See <see cref="TomlNode.Select"/> for
    /// the grammar; <c>tool.ruff.line-length</c> and <c>package[0].name</c> are typical paths.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is malformed.</exception>
    public TomlNode? Select(string path) => _root.Select(path);

    /// <summary>
    /// Sets the value a path addresses and returns the node now there. The last step must be a key or an array index;
    /// missing tables on the way are added with <see cref="TomlTable.AddTable"/>, so
    /// <c>Set("tool.ruff.line-length", 100)</c> works on a file without a <c>[tool.ruff]</c> table.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> or <paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is malformed or <paramref name="value"/> is not representable.</exception>
    /// <exception cref="InvalidOperationException">A step goes through a value, or an index is outside its array.</exception>
    public TomlNode Set(string path, object value) => TomlPath.Set(_root, path, value);

    /// <summary>Removes the node a path addresses and returns whether it existed.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is malformed or addresses the root.</exception>
    public bool Remove(string path) => TomlPath.Remove(_root, path);

    /// <summary>Releases nothing; a document holds only managed memory. It exists so <c>using</c> works as for every OfficeIMO document.</summary>
    public void Dispose() {
    }

    internal TomlLayout Layout => _layout ??= new TomlLayout(this);

    /// <summary>Converts an offset into the source to a position, building the line table on first use.</summary>
    internal TomlPosition GetPosition(int offset) {
        int[] lines = _lineStarts ??= BuildLineStarts(_source);
        int low = 0, high = lines.Length - 1;
        while (low < high) {
            int mid = (low + high + 1) >> 1;
            if (lines[mid] <= offset) low = mid;
            else high = mid - 1;
        }
        return new TomlPosition(offset, low + 1, offset - lines[low] + 1);
    }

    private static int[] BuildLineStarts(string text) {
        int count = 1;
        for (int i = 0; i < text.Length; i++) {
            if (text[i] == '\n') count++;
        }
        var starts = new int[count];
        int line = 1;
        for (int i = 0; i < text.Length; i++) {
            if (text[i] == '\n') starts[line++] = i + 1;
        }
        return starts;
    }
}
