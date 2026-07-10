using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Loss-preserving RTF syntax tree produced from tokens.
/// </summary>
public sealed class RtfSyntaxTree {
    internal RtfSyntaxTree(RtfGroup root, IReadOnlyList<RtfDiagnostic> diagnostics, string? sourcePrefix = null, string? sourceSuffix = null) {
        Root = root ?? throw new ArgumentNullException(nameof(root));
        IReadOnlyList<RtfDiagnostic> suppliedDiagnostics = diagnostics ?? Array.Empty<RtfDiagnostic>();
        if (suppliedDiagnostics is RtfSyntaxDiagnosticCollection carried) {
            SourcePrefix = sourcePrefix ?? carried.SourcePrefix;
            SourceSuffix = sourceSuffix ?? carried.SourceSuffix;
            Diagnostics = suppliedDiagnostics;
        } else {
            SourcePrefix = sourcePrefix ?? string.Empty;
            SourceSuffix = sourceSuffix ?? string.Empty;
            Diagnostics = new RtfSyntaxDiagnosticCollection(suppliedDiagnostics, SourcePrefix, SourceSuffix);
        }
    }

    /// <summary>Root RTF group.</summary>
    public RtfGroup Root { get; }

    /// <summary>Parser diagnostics.</summary>
    public IReadOnlyList<RtfDiagnostic> Diagnostics { get; }

    internal string SourcePrefix { get; }

    internal string SourceSuffix { get; }

    internal RtfSyntaxTree WithRoot(RtfGroup root) => new RtfSyntaxTree(root, Diagnostics, SourcePrefix, SourceSuffix);

    private sealed class RtfSyntaxDiagnosticCollection : ReadOnlyCollection<RtfDiagnostic> {
        internal RtfSyntaxDiagnosticCollection(IEnumerable<RtfDiagnostic> diagnostics, string sourcePrefix, string sourceSuffix)
            : base(diagnostics.ToList()) {
            SourcePrefix = sourcePrefix;
            SourceSuffix = sourceSuffix;
        }

        internal string SourcePrefix { get; }

        internal string SourceSuffix { get; }
    }

    /// <summary>
    /// Parses RTF content into a syntax tree.
    /// </summary>
    public static RtfSyntaxTree Parse(string rtf) => RtfSyntaxParser.Parse(rtf);

    /// <summary>
    /// Parses RTF content into a syntax tree while limiting nested group depth.
    /// </summary>
    public static RtfSyntaxTree Parse(string rtf, int maxDepth) => RtfSyntaxParser.Parse(rtf, maxDepth);

    /// <summary>
    /// Parses RTF content using configured resource limits and cancellation.
    /// </summary>
    public static RtfSyntaxTree Parse(string rtf, RtfReadOptions? options, CancellationToken cancellationToken = default) =>
        RtfSyntaxParser.Parse(rtf, options, cancellationToken);

    /// <summary>
    /// Serializes the original syntax tree without semantic normalization.
    /// </summary>
    public string ToRtf() => RtfSyntaxWriter.Write(this);

    /// <summary>
    /// Creates an editor for targeted syntax-preserving changes.
    /// </summary>
    public RtfLosslessEditor EditLossless() => new RtfLosslessEditor(this);
}
