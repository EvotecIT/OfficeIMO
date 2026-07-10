namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Parsed AsciiDoc document with a lossless syntax tree and typed editable top-level blocks.
/// </summary>
public sealed class AsciiDocDocument {
    private readonly IReadOnlyList<AsciiDocBlock> _blocks;
    private readonly IReadOnlyList<AsciiDocDiagnostic> _diagnostics;

    internal AsciiDocDocument(
        AsciiDocSourceText source,
        AsciiDocSyntaxTree syntaxTree,
        IReadOnlyList<AsciiDocBlock> blocks,
        IReadOnlyList<AsciiDocDiagnostic> diagnostics) {
        Source = source;
        SyntaxTree = syntaxTree;
        _blocks = blocks;
        _diagnostics = diagnostics;
    }

    /// <summary>Original source text and line mapping.</summary>
    public AsciiDocSourceText Source { get; }

    /// <summary>Lossless syntax tree.</summary>
    public AsciiDocSyntaxTree SyntaxTree { get; }

    /// <summary>Typed top-level source blocks, including trivia and comments.</summary>
    public IReadOnlyList<AsciiDocBlock> Blocks => _blocks;

    /// <summary>Parser and recovery diagnostics.</summary>
    public IReadOnlyList<AsciiDocDiagnostic> Diagnostics => _diagnostics;

    /// <summary>True when any editable block or list item has changed.</summary>
    public bool IsModified => Blocks.Any(static block => block.IsModified);

    /// <summary>Parses an AsciiDoc string.</summary>
    public static AsciiDocParseResult Parse(string source, AsciiDocParseOptions? options = null) =>
        AsciiDocParser.Parse(source, options);

    /// <summary>
    /// Loads and parses an AsciiDoc file using the runtime's UTF-8 BOM detection.
    /// Retains decoded characters and line endings, not original encoding or BOM bytes.
    /// </summary>
    public static AsciiDocParseResult Load(string path, AsciiDocParseOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return Parse(File.ReadAllText(path), options);
    }

    /// <summary>Enumerates blocks of a requested semantic type.</summary>
    public IEnumerable<TBlock> BlocksOfType<TBlock>() where TBlock : AsciiDocBlock => Blocks.OfType<TBlock>();

    /// <summary>Builds the effective document attribute set in source order.</summary>
    public AsciiDocDocumentAttributes GetAttributes(IReadOnlyDictionary<string, string>? initialValues = null) {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (initialValues != null) {
            foreach (KeyValuePair<string, string> value in initialValues) values[value.Key] = value.Value;
        }
        foreach (AsciiDocAttributeEntry entry in BlocksOfType<AsciiDocAttributeEntry>()) {
            if (entry.IsUnset) values.Remove(entry.Name);
            else values[entry.Name] = entry.Value;
        }
        return new AsciiDocDocumentAttributes(values);
    }

    /// <summary>Writes this document using preserve mode.</summary>
    public string ToAsciiDoc() => AsciiDocWriter.Write(this, null);

    /// <summary>Writes this document using the requested mode.</summary>
    public string ToAsciiDoc(AsciiDocWriterMode mode) =>
        AsciiDocWriter.Write(this, new AsciiDocWriterOptions { Mode = mode });

    /// <summary>Writes this document with explicit options.</summary>
    public string ToAsciiDoc(AsciiDocWriterOptions? options) => AsciiDocWriter.Write(this, options);

    /// <summary>
    /// Saves the current document text using the runtime's default UTF-8 behavior.
    /// This does not reproduce an original file encoding or BOM.
    /// </summary>
    public void Save(string path, AsciiDocWriterOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        File.WriteAllText(path, ToAsciiDoc(options));
    }
}
