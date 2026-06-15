using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Loss-preserving RTF syntax tree produced from tokens.
/// </summary>
public sealed class RtfSyntaxTree {
    internal RtfSyntaxTree(RtfGroup root, IReadOnlyList<RtfDiagnostic> diagnostics) {
        Root = root ?? throw new ArgumentNullException(nameof(root));
        Diagnostics = diagnostics ?? Array.Empty<RtfDiagnostic>();
    }

    /// <summary>Root RTF group.</summary>
    public RtfGroup Root { get; }

    /// <summary>Parser diagnostics.</summary>
    public IReadOnlyList<RtfDiagnostic> Diagnostics { get; }

    /// <summary>
    /// Parses RTF content into a syntax tree.
    /// </summary>
    public static RtfSyntaxTree Parse(string rtf) => RtfSyntaxParser.Parse(rtf);

    /// <summary>
    /// Serializes the original syntax tree without semantic normalization.
    /// </summary>
    public string ToRtf() => RtfSyntaxWriter.Write(this);

    /// <summary>
    /// Creates an editor for targeted syntax-preserving changes.
    /// </summary>
    public RtfLosslessEditor EditLossless() => new RtfLosslessEditor(this);
}
