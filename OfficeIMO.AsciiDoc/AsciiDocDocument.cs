namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Parsed AsciiDoc document with a lossless syntax tree and typed editable top-level blocks.
/// </summary>
public sealed partial class AsciiDocDocument {
    private readonly IReadOnlyList<AsciiDocBlock> _blocks;
    private readonly List<AsciiDocBlock> _editableBlocks;
    private bool _structureWasModified;
    private readonly IReadOnlyList<AsciiDocDiagnostic> _diagnostics;

    internal AsciiDocDocument(
        AsciiDocSourceText source,
        AsciiDocSyntaxTree syntaxTree,
        IReadOnlyList<AsciiDocBlock> blocks,
        IReadOnlyList<AsciiDocDiagnostic> diagnostics,
        AsciiDocDocumentProfile profile) {
        Source = source;
        SyntaxTree = syntaxTree;
        _editableBlocks = blocks.ToList();
        _blocks = _editableBlocks.AsReadOnly();
        _diagnostics = diagnostics;
        Profile = profile;
    }

    /// <summary>Original source text and line mapping.</summary>
    public AsciiDocSourceText Source { get; }

    /// <summary>Lossless syntax tree.</summary>
    public AsciiDocSyntaxTree SyntaxTree { get; }

    /// <summary>Selected bounded document profile.</summary>
    public AsciiDocDocumentProfile Profile { get; }

    /// <summary>Typed top-level source blocks, including trivia and comments.</summary>
    public IReadOnlyList<AsciiDocBlock> Blocks => _blocks;

    /// <summary>Parser and recovery diagnostics.</summary>
    public IReadOnlyList<AsciiDocDiagnostic> Diagnostics => _diagnostics;

    /// <summary>True when any editable block or list item has changed.</summary>
    public bool IsModified => _structureWasModified || Blocks.Any(static block => block.IsModified);

    internal bool IsStructureModified => _structureWasModified;

    /// <summary>Parses an AsciiDoc string into the typed document model.</summary>
    public static AsciiDocDocument Parse(string source, AsciiDocParseOptions? options = null) =>
        ParseResult(source, options).Document;

    /// <summary>Parses an AsciiDoc string with syntax and recovery diagnostics.</summary>
    public static AsciiDocParseResult ParseResult(string source, AsciiDocParseOptions? options = null) =>
        AsciiDocParser.Parse(source, options);

    /// <summary>Parses an AsciiDoc string with cooperative cancellation.</summary>
    public static AsciiDocDocument Parse(string source, AsciiDocParseOptions? options, System.Threading.CancellationToken cancellationToken) =>
        ParseResult(source, options, cancellationToken).Document;

    /// <summary>Parses source with recovery diagnostics and cooperative cancellation.</summary>
    public static AsciiDocParseResult ParseResult(string source, AsciiDocParseOptions? options, System.Threading.CancellationToken cancellationToken) =>
        AsciiDocParser.Parse(source, options, cancellationToken);

    /// <summary>
    /// Loads and parses an AsciiDoc file using the selected encoding or Unicode BOM detection with a UTF-8 default.
    /// Retains decoded characters and line endings, not original encoding or BOM bytes.
    /// </summary>
    public static AsciiDocDocument Load(string path, AsciiDocParseOptions? options = null, Encoding? encoding = null) =>
        LoadResult(path, options, encoding).Document;

    /// <summary>
    /// Loads an AsciiDoc file with its lossless syntax and recovery diagnostics.
    /// Retains decoded characters and line endings, not original encoding or BOM bytes.
    /// </summary>
    public static AsciiDocParseResult LoadResult(string path, AsciiDocParseOptions? options = null, Encoding? encoding = null) {
        return LoadResult(path, options, encoding, System.Threading.CancellationToken.None);
    }

    /// <summary>Loads and parses an AsciiDoc file with cooperative cancellation.</summary>
    public static AsciiDocDocument Load(string path, AsciiDocParseOptions? options, Encoding? encoding, System.Threading.CancellationToken cancellationToken) =>
        LoadResult(path, options, encoding, cancellationToken).Document;

    /// <summary>Loads a bounded AsciiDoc file and returns recovery diagnostics with cooperative cancellation.</summary>
    public static AsciiDocParseResult LoadResult(string path, AsciiDocParseOptions? options, Encoding? encoding, System.Threading.CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return LoadResult(stream, options, encoding, cancellationToken);
    }

    /// <summary>Enumerates blocks of a requested semantic type.</summary>
    public IEnumerable<TBlock> BlocksOfType<TBlock>() where TBlock : AsciiDocBlock => Blocks.OfType<TBlock>();

    /// <summary>Builds the effective document attribute set in source order.</summary>
    public AsciiDocDocumentAttributes GetAttributes(IReadOnlyDictionary<string, string>? initialValues = null) {
        AsciiDocDocumentAttributes attributes = AsciiDocDocumentAttributes.Create(initialValues);
        foreach (AsciiDocAttributeEntry entry in GetAttributeEntries()) {
            attributes = attributes.Apply(entry);
        }
        return attributes;
    }

    /// <summary>Enumerates source blocks with immutable attribute snapshots in document order.</summary>
    public IEnumerable<AsciiDocBlockContext> GetBlockContexts(IReadOnlyDictionary<string, string>? initialValues = null, bool expandAssignmentValues = true) {
        return GetBlockContexts(initialValues, expandAssignmentValues, 64, default);
    }

    /// <summary>Enumerates attribute snapshots with bounded compound traversal and cooperative cancellation.</summary>
    public IEnumerable<AsciiDocBlockContext> GetBlockContexts(IReadOnlyDictionary<string, string>? initialValues, bool expandAssignmentValues, int maximumNestingDepth, System.Threading.CancellationToken cancellationToken = default) {
        if (maximumNestingDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumNestingDepth));
        AsciiDocDocumentAttributes attributes = AsciiDocDocumentAttributes.Create(initialValues);
        foreach (AsciiDocBlock block in Blocks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (block is AsciiDocAttributeEntry entry) attributes = attributes.Apply(entry, expandAssignmentValues);
            yield return new AsciiDocBlockContext(block, attributes);
            if (block is AsciiDocDelimitedBlock compound && compound.GetBody(cancellationToken) is AsciiDocDocument body) {
                if (maximumNestingDepth == 1) throw new InvalidDataException("AsciiDoc attribute traversal exceeds MaximumNestingDepth.");
                foreach (AsciiDocAttributeEntry child in body.GetAttributeEntries(maximumNestingDepth - 1, cancellationToken)) attributes = attributes.Apply(child, expandAssignmentValues);
            }
        }
    }

    /// <summary>Enumerates attribute entries in source order, including compound bodies, with a bounded traversal depth.</summary>
    public IEnumerable<AsciiDocAttributeEntry> GetAttributeEntries(int maximumNestingDepth = 64, System.Threading.CancellationToken cancellationToken = default) {
        if (maximumNestingDepth < 1) throw new ArgumentOutOfRangeException(nameof(maximumNestingDepth));
        return EnumerateAttributeEntries(this, maximumNestingDepth, 0, cancellationToken);
    }

    private static IEnumerable<AsciiDocAttributeEntry> EnumerateAttributeEntries(AsciiDocDocument document, int maximumDepth, int depth, System.Threading.CancellationToken token) {
        if (depth >= maximumDepth) throw new InvalidDataException("AsciiDoc attribute traversal exceeds MaximumNestingDepth.");
        foreach (AsciiDocBlock block in document.Blocks) {
            token.ThrowIfCancellationRequested();
            if (block is AsciiDocAttributeEntry entry) yield return entry;
            if (block is AsciiDocDelimitedBlock compound && compound.GetBody(token) is AsciiDocDocument body)
                foreach (AsciiDocAttributeEntry child in EnumerateAttributeEntries(body, maximumDepth, depth + 1, token)) yield return child;
        }
    }

    /// <summary>Writes this document using preserve mode.</summary>
    public string ToAsciiDoc() => AsciiDocWriter.Write(this, null);

    /// <summary>Writes this document using the requested mode.</summary>
    public string ToAsciiDoc(AsciiDocWriterMode mode) =>
        AsciiDocWriter.Write(this, new AsciiDocWriterOptions { Mode = mode });

    /// <summary>Writes this document with explicit options.</summary>
    public string ToAsciiDoc(AsciiDocWriterOptions? options) => AsciiDocWriter.Write(this, options);

}
