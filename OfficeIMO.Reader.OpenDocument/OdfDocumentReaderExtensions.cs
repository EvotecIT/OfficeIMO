using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Projects caller-owned native OpenDocument models through the canonical Reader adapter.</summary>
public static class OdfDocumentReaderExtensions {
    /// <summary>Reads a loaded native document into the shared document envelope without reopening or disposing it.</summary>
    /// <remarks>Reader projection and output budgets apply to this operation. Source input and native load limits must
    /// be enforced when loading the document; this method cannot defer binary data already read by an earlier load.
    /// The source name supplies logical locations and is not opened or hashed as a file.</remarks>
    public static OfficeDocumentReadResult ReadDocument(this OdfDocument document, string? sourceName = null,
        ReaderOptions? options = null, ReaderOpenDocumentOptions? formatOptions = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        ReaderOptions effective = DocumentReaderEngine.NormalizeOptions(options);
        ReaderOpenDocumentOptions format = (formatOptions ?? new ReaderOpenDocumentOptions()).Clone();
        string name = string.IsNullOrWhiteSpace(sourceName) ? "document.odf" : sourceName!.Trim();
        using var scope = ReaderReadScope.Enter(effective);
        OfficeDocumentReadResult result = OpenDocumentReaderAdapter.ProjectDocument(document, name, effective, format, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Source.Path ??= name;
        return ReaderReadScope.Complete(result);
    }
}
