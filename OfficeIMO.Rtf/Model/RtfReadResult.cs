using OfficeIMO.Rtf.Diagnostics;
using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

/// <summary>
/// Result of reading RTF into syntax and semantic models.
/// </summary>
public sealed class RtfReadResult {
    internal RtfReadResult(RtfDocument document, RtfSyntaxTree syntaxTree, IReadOnlyList<RtfDiagnostic> diagnostics) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        SyntaxTree = syntaxTree ?? throw new ArgumentNullException(nameof(syntaxTree));
        Diagnostics = diagnostics ?? Array.Empty<RtfDiagnostic>();
    }

    /// <summary>Semantic document model.</summary>
    public RtfDocument Document { get; }

    /// <summary>Loss-preserving syntax tree.</summary>
    public RtfSyntaxTree SyntaxTree { get; }

    /// <summary>Combined parser and binder diagnostics.</summary>
    public IReadOnlyList<RtfDiagnostic> Diagnostics { get; }

    /// <summary>
    /// Serializes the original syntax tree without semantic normalization.
    /// </summary>
    public string ToRtfLossless() => SyntaxTree.ToRtf();

    /// <summary>
    /// Creates an editor for targeted syntax-preserving changes.
    /// </summary>
    public RtfLosslessEditor EditLossless() => new RtfLosslessEditor(this);

    /// <summary>
    /// Saves the original RTF stream to a file without semantic normalization.
    /// </summary>
    public void SaveLossless(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        RtfBytePreservingEncoding.WriteAllText(path, ToRtfLossless());
    }

    /// <summary>
    /// Saves the original RTF stream to a stream without semantic normalization.
    /// </summary>
    public void SaveLossless(Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        RtfBytePreservingEncoding.WriteTo(stream, ToRtfLossless());
    }
}
