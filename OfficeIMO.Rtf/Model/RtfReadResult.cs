using OfficeIMO.Rtf.Diagnostics;
using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

/// <summary>
/// Result of reading RTF into syntax and semantic models.
/// </summary>
public sealed partial class RtfReadResult : IOfficeResult<RtfDocument> {
    private byte[]? _originalBytes;

    internal RtfReadResult(RtfDocument document, RtfSyntaxTree syntaxTree, IReadOnlyList<RtfDiagnostic> diagnostics) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        SyntaxTree = syntaxTree ?? throw new ArgumentNullException(nameof(syntaxTree));
        Diagnostics = diagnostics ?? Array.Empty<RtfDiagnostic>();
    }

    /// <summary>Semantic document model.</summary>
    public RtfDocument Document { get; }

    /// <inheritdoc />
    public RtfDocument Value => Document;

    /// <inheritdoc />
    public bool Succeeded => true;

    /// <inheritdoc />
    public RtfDocument RequireValue() => Document;

    /// <summary>Loss-preserving syntax tree.</summary>
    public RtfSyntaxTree SyntaxTree { get; }

    /// <summary>Combined parser and binder diagnostics.</summary>
    public IReadOnlyList<RtfDiagnostic> Diagnostics { get; }

    /// <summary>Gets whether the read API retained the exact original source bytes.</summary>
    public bool HasOriginalBytes => _originalBytes != null;

    /// <summary>Gets whether <see cref="ToBytesLossless"/> can return bytes without character transcoding.</summary>
    public bool CanWriteLosslessBytes => _originalBytes != null || RtfBytePreservingEncoding.CanEncode(ToRtfLossless());

    /// <summary>
    /// Serializes the original syntax tree without semantic normalization.
    /// </summary>
    public string ToRtfLossless() => SyntaxTree.ToRtf();

    /// <summary>
    /// Serializes the original syntax tree to source-preserving bytes without semantic normalization.
    /// </summary>
    public byte[] ToBytesLossless() => _originalBytes != null
        ? (byte[])_originalBytes.Clone()
        : RtfBytePreservingEncoding.ToBytes(ToRtfLossless());

    /// <summary>Attempts to return source-preserving bytes without throwing for character-only input.</summary>
    public bool TryGetLosslessBytes(out byte[] bytes) {
        if (_originalBytes != null) {
            bytes = (byte[])_originalBytes.Clone();
            return true;
        }
        string rtf = ToRtfLossless();
        if (!RtfBytePreservingEncoding.CanEncode(rtf)) {
            bytes = Array.Empty<byte>();
            return false;
        }
        bytes = RtfBytePreservingEncoding.ToBytes(rtf);
        return true;
    }

    /// <summary>
    /// Creates an editor for targeted syntax-preserving changes.
    /// </summary>
    public RtfLosslessEditor EditLossless() => new RtfLosslessEditor(this);

    /// <summary>
    /// Saves the original RTF stream to a file without semantic normalization.
    /// </summary>
    public void SaveLossless(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytes(path, ToBytesLossless());
    }

    /// <summary>
    /// Saves the original RTF stream to a stream without semantic normalization.
    /// </summary>
    public void SaveLossless(Stream stream) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        OfficeIMO.Core.Internal.OfficeStreamWriter.WriteAllBytes(stream, ToBytesLossless());
    }

    internal RtfReadResult AttachOriginalBytes(byte[] bytes) {
        _originalBytes = bytes == null ? throw new ArgumentNullException(nameof(bytes)) : (byte[])bytes.Clone();
        return this;
    }
}
