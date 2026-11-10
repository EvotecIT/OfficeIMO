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
public sealed partial class YamlDocument {
    internal readonly string _source;
    internal readonly YamlTextEncoding _encoding;
    internal readonly bool _byteOrderMark;
    private readonly List<YamlDocumentPart> _parts = new List<YamlDocumentPart>();
    private List<OfficeDiagnostic>? _diagnostics;
    private int[]? _lineStarts;

    private YamlDocument(string source, YamlTextEncoding encoding, bool byteOrderMark, YamlDuplicateKeyPolicy duplicateKeys) {
        _source = source;
        _encoding = encoding;
        _byteOrderMark = byteOrderMark;
        DuplicateKeys = duplicateKeys;
    }

    /// <summary>Gets the policy the document was loaded with; lookups among equal keys follow it.</summary>
    internal YamlDuplicateKeyPolicy DuplicateKeys { get; }

    /// <summary>Gets the documents of the stream in source order. An empty or comment-only stream has none.</summary>
    public IReadOnlyList<YamlDocumentPart> Documents => _parts;

    /// <summary>Gets the root node of the first document, or <see langword="null"/> when the stream has no documents.</summary>
    public YamlNode? Root => _parts.Count == 0 ? null : _parts[0].Root;

    /// <summary>
    /// Gets findings recorded while loading, such as each duplicate key a lenient <see cref="YamlDuplicateKeyPolicy"/>
    /// admitted (code <see cref="YamlErrorCodes.DuplicateKey"/>, with its line and column as the location).
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
