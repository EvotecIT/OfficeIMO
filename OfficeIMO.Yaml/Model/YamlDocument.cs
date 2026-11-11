namespace OfficeIMO.Yaml;

/// <summary>
/// A YAML file or stream loaded into a lossless model: every document, node, comment, blank line and spelling of the
/// input is kept, and saving an unchanged document reproduces the input byte for byte, including its encoding and
/// byte-order mark.
/// </summary>
/// <remarks>
/// <para>The document retains its decoded source text once; nodes are offsets into it, so a node costs about one object
/// and comments and whitespace cost nothing until asked for. Values are decoded lazily and cached.</para>
/// <para>Loading enforces the limits in <see cref="YamlLoadOptions"/>, including the alias-expansion bound that stops
/// "billion laughs" inputs, and resolves every alias to its anchor. Malformed input throws <see cref="YamlException"/>.</para>
/// </remarks>
public sealed partial class YamlDocument : IDisposable {
    internal readonly string _source;
    internal readonly YamlTextEncoding _encoding;
    internal readonly bool _byteOrderMark;
    private readonly List<YamlDocumentPart> _parts = new List<YamlDocumentPart>();
    private List<OfficeDiagnostic>? _diagnostics;
    private int[]? _lineStarts;
    private YamlLayoutProfile? _layout;
    private YamlWriterOptions? _writerOptions;
    private HashSet<YamlNode>? _sharedEdits;
    /// <summary>Set by any edit; an unedited document saves as one copy of its source.</summary>
    internal bool _partsDirty;
    private string? _path;
    /// <summary>The loaded nodes; null for a document created in memory.</summary>
    internal YamlNodeTable? _table;

    private YamlDocument(string source, YamlTextEncoding encoding, bool byteOrderMark, YamlDuplicateKeyPolicy duplicateKeys, bool strict) {
        _source = source;
        _encoding = encoding;
        _byteOrderMark = byteOrderMark;
        DuplicateKeys = duplicateKeys;
        Strict = strict;
    }

    /// <summary>
    /// Creates an empty document in memory, with no YAML documents in its stream. Set <see cref="Root"/>, call
    /// <see cref="Set"/> with a path, or <see cref="AddDocument"/> to give it content; saving writes UTF-8 without a
    /// byte-order mark.
    /// </summary>
    /// <example>
    /// <code>
    /// using var doc = YamlDocument.Create();
    /// doc.Set("name", "web");
    /// doc.Set("spec.replicas", 3);
    /// doc.Add("spec.ports", 80);
    /// doc.Save("service.yaml");
    /// </code>
    /// </example>
    public static YamlDocument Create() => new YamlDocument(string.Empty, YamlTextEncoding.Utf8, byteOrderMark: false, YamlDuplicateKeyPolicy.LastWins, strict: false);

    /// <summary>
    /// Creates an empty document in memory associated with a path, without writing a file. <see cref="Save()"/> writes
    /// it there.
    /// </summary>
    /// <param name="path">The file the document will be saved to.</param>
    public static YamlDocument Create(string path) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        YamlDocument document = Create();
        document._path = System.IO.Path.GetFullPath(path);
        return document;
    }

    /// <summary>
    /// Gets the full path that <see cref="Save()"/> writes to: the path of <see cref="Create(string)"/>, or of the last
    /// file load or save. Documents parsed from text or loaded from a stream have none until they are saved to a path.
    /// </summary>
    public string? FilePath => _path;

    /// <summary>Gets the policy the document was loaded with; lookups among equal keys follow it.</summary>
    internal YamlDuplicateKeyPolicy DuplicateKeys { get; }

    /// <summary>Gets whether the document was loaded with <see cref="YamlLoadOptions.Strict"/>; without it, merge keys also apply under the core schema.</summary>
    internal bool Strict { get; }

    /// <summary>Gets the documents of the stream in source order. An empty or comment-only stream has none.</summary>
    public IReadOnlyList<YamlDocumentPart> Documents => _parts;

    /// <summary>
    /// Gets the root node of the first document, or <see langword="null"/> when the stream has no documents. Setting it
    /// replaces that root, or adds a first document to an empty stream.
    /// </summary>
    public YamlNode? Root {
        get => _parts.Count == 0 ? null : _parts[0].Root;
        set {
            if (_parts.Count == 0) AddDocument(value ?? new YamlScalar((string?)null));
            else _parts[0].Root = value!;
        }
    }

    /// <summary>
    /// Gets findings recorded while loading and editing, with a line and column as the location: each duplicate key a
    /// non-strict load or a lenient <see cref="YamlDuplicateKeyPolicy"/> admitted (<see cref="OfficeDiagnosticCodes.YamlDuplicateKey"/>), each
    /// indentation the specification rejects that a non-strict load accepted (<see cref="OfficeDiagnosticCodes.YamlLenientIndentation"/>),
    /// one note per document part that uses <c>&lt;&lt;</c> merge keys without a <c>%YAML 1.1</c> directive
    /// (<see cref="OfficeDiagnosticCodes.YamlMergeKeyWithoutDirective"/>), and each edit of a node that aliases share
    /// (<see cref="OfficeDiagnosticCodes.YamlAliasedNodeEdited"/>).
    /// </summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics => (IReadOnlyList<OfficeDiagnostic>?)_diagnostics ?? Array.Empty<OfficeDiagnostic>();

    /// <summary>
    /// Returns the node a path addresses, or <see langword="null"/>. With one document the path starts at its root; with
    /// several, a leading <c>$[n]</c> picks the zero-based document, as the paths <c>OfficeIMO.Reader.Yaml</c> reports for
    /// multi-document streams do, and paths without it start at the first document. See <see cref="YamlNode.Select"/>
    /// for the grammar.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is malformed.</exception>
    public YamlNode? Select(string path) => YamlPath.Select(this, path);

    /// <summary>
    /// Sets the value at a path, creating missing mappings on the way (and a sequence where the next step is an index),
    /// in the layout of their neighbors. The last step sets a mapping value as <see cref="YamlMapping.Set"/> does, replaces
    /// a sequence item, or appends when the index equals the sequence's count. <c>$</c> or an empty path sets the root.
    /// </summary>
    /// <returns>The node now at the path.</returns>
    /// <exception cref="ArgumentException">The path is malformed.</exception>
    /// <exception cref="ArgumentOutOfRangeException">An index is past the end of its sequence.</exception>
    /// <exception cref="InvalidOperationException">The path goes through a scalar, or a key step meets a sequence or an index step a mapping.</exception>
    public YamlNode Set(string path, YamlNode? value) => YamlPath.Set(this, path, value ?? new YamlScalar((string?)null));

    /// <summary>
    /// Appends a value to the sequence at a path, creating the sequence and missing mappings on the way.
    /// </summary>
    /// <returns>The added node.</returns>
    /// <exception cref="InvalidOperationException">The path addresses something other than a sequence, or goes through a scalar.</exception>
    public YamlNode Add(string path, YamlNode? value) => YamlPath.Add(this, path, value ?? new YamlScalar((string?)null));

    /// <summary>
    /// Removes the mapping entry or sequence item a path addresses, with its comment lines and its line.
    /// </summary>
    /// <returns><see langword="true"/> when something was removed; <see langword="false"/> when nothing is at the path.</returns>
    /// <exception cref="ArgumentException">The path is malformed or addresses a document root.</exception>
    /// <exception cref="InvalidOperationException">The path goes through a scalar, or aliases refer to an anchor inside the removed node.</exception>
    public bool Remove(string path) => YamlPath.Remove(this, path);

    /// <summary>
    /// Deep-merges an overlay document into this one, document part by part in order, the way Helm combines values files.
    /// Mappings merge key by key; overlay scalars are assigned, so an equal value keeps this document's spelling;
    /// sequences follow <see cref="YamlMergeOptions.Sequences"/>; new keys are appended with the overlay's comments in this
    /// document's layout. When a value changes kind (a mapping over a scalar, for example) the overlay wins and the
    /// conflict is reported. The overlay is not changed, and saving writes only what the merge touched.
    /// </summary>
    /// <remarks>
    /// <para>An overlay value that lands on an alias edits the anchored node, as any edit through an alias does, and is
    /// reported. Overlay anchors and aliases are copied with their names; a name the document already uses is renamed and
    /// reported. <c>&lt;&lt;</c> merge keys follow the typed view: the overlay's merged entries count as its entries, and a
    /// key this document gets only from a merge receives an entry of its own; merge keys are never flattened.</para>
    /// <para>Overlay parts beyond this document's are added. An empty overlay part changes nothing. A merge is not
    /// transactional: when an edit fails, the edits before it stay.</para>
    /// </remarks>
    /// <example>
    /// <code>
    /// using var values = YamlDocument.Load("values.yaml");
    /// using var production = YamlDocument.Load("values.prod.yaml");
    /// YamlMergeReport report = values.Merge(production, new YamlMergeOptions { NullRemovesKey = true });
    /// values.Save("values.merged.yaml");
    /// </code>
    /// </example>
    /// <returns>The added, replaced and removed paths and the findings of the merge.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="overlay"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="overlay"/> is this document.</exception>
    public YamlMergeReport Merge(YamlDocument overlay, YamlMergeOptions? options = null) => YamlMerge.Merge(this, overlay, options);

    /// <summary>Adds a document with the given root at the end of the stream, after a <c>---</c> marker when it is not the first.</summary>
    /// <returns>The new document part.</returns>
    public YamlDocumentPart AddDocument(YamlNode root) {
        if (root is null) throw new ArgumentNullException(nameof(root));
        var part = new YamlDocumentPart(this, YamlSchema.Core, hasStartMarker: _parts.Count > 0);
        _parts.Add(part);
        try {
            YamlEdit.ReplaceRoot(part, root);
        } catch {
            _parts.Remove(part);
            throw;
        }
        _partsDirty = true;
        return part;
    }

    /// <summary>
    /// Returns the pooled storage of the loaded nodes when no node was read; the document holds only managed memory.
    /// Do not use the document after disposing it: nodes not reached before are no longer available.
    /// </summary>
    public void Dispose() {
        _table?.Release();
    }

    /// <summary>The layout conventions new nodes follow, read from the source on first use.</summary>
    internal YamlLayoutProfile Layout => _layout ??= YamlLayoutProfile.Infer(this);

    /// <summary>Writer settings that reproduce the document's layout for new text.</summary>
    internal YamlWriterOptions WriterOptions => _writerOptions ??= Layout.ToWriterOptions();

    /// <summary>Returns true the first time an edit of a shared anchored node is noted.</summary>
    internal bool NoteSharedEdit(YamlNode anchored) => (_sharedEdits ??= new HashSet<YamlNode>()).Add(anchored);

    internal void AddPart(YamlDocumentPart part) => _parts.Add(part);

    internal void AddDiagnostic(OfficeDiagnostic diagnostic) => (_diagnostics ??= new List<OfficeDiagnostic>()).Add(diagnostic);

    /// <summary>Converts an offset into the source to a position, building the line table on first use.</summary>
    internal YamlPosition GetPosition(int offset) {
        int[] lines = _lineStarts ??= BuildLineStarts(_source);
        int low = 0, high = lines.Length - 1;
        while (low < high) {
            int mid = (low + high + 1) >> 1;
            if (lines[mid] <= offset) low = mid;
            else high = mid - 1;
        }
        return new YamlPosition(offset, low + 1, offset - lines[low] + 1);
    }

    private static int[] BuildLineStarts(string text) {
        int count = 1;
        for (int i = 0; i < text.Length; i++) {
            char ch = text[i];
            if (ch == '\n' || (ch == '\r' && (i + 1 == text.Length || text[i + 1] != '\n'))) count++;
        }
        var starts = new int[count];
        int line = 1;
        for (int i = 0; i < text.Length; i++) {
            char ch = text[i];
            if (ch == '\n' || (ch == '\r' && (i + 1 == text.Length || text[i + 1] != '\n'))) starts[line++] = i + 1;
        }
        return starts;
    }
}
