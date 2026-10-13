using OfficeIMO.Rtf;

namespace OfficeIMO.Reader.Rtf;

/// <summary>Projects caller-owned RTF read results through the canonical Reader adapter.</summary>
public static class RtfReadResultReaderExtensions {
    /// <summary>Reads a loaded RTF result into the shared document envelope, retaining original parser diagnostics.</summary>
    /// <remarks>The native result and its original-source payload inventory remain caller-owned. Reader projection
    /// and output budgets apply here; input and parser limits must be enforced during the preceding native load.
    /// This method does not reopen the source or defer data already decoded. The logical source name is not opened or hashed.</remarks>
    public static OfficeDocumentReadResult ReadDocument(this RtfReadResult readResult, string? sourceName = null,
        ReaderOptions? options = null, ReaderRtfOptions? formatOptions = null, CancellationToken cancellationToken = default) {
        if (readResult == null) throw new ArgumentNullException(nameof(readResult));
        cancellationToken.ThrowIfCancellationRequested();
        ReaderOptions effective = DocumentReaderEngine.NormalizeOptions(options);
        ReaderRtfOptions format = ReaderRtfOptionsCloner.CloneOrDefault(formatOptions);
        string name = string.IsNullOrWhiteSpace(sourceName) ? "document.rtf" : sourceName!.Trim();
        using var scope = ReaderReadScope.Enter(effective);
        OfficeDocumentReadResult result = RtfReaderAdapter.ProjectDocument(readResult, name, effective, format, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Source.Path ??= name;
        return ReaderReadScope.Complete(result);
    }
}
