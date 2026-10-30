using OfficeIMO.Publisher;

namespace OfficeIMO.Reader.Publisher;

/// <summary>Projects a loaded Publisher document into the shared Reader model.</summary>
public static class PublisherDocumentReaderExtensions {
    /// <summary>Reads the document's current stories, tables, page inventory and requested image payloads.</summary>
    /// <param name="document">The caller-owned document, including any in-memory edits.</param>
    /// <param name="sourceName">Logical citation name. Omitted names use document.pub.</param>
    /// <param name="readerOptions">Projection and output budgets, copied for this operation.</param>
    /// <param name="publisherOptions">Publisher semantic projection settings, copied for this operation.</param>
    /// <param name="cancellationToken">Cancellation for the projection.</param>
    /// <remarks>
    /// The logical source name is not opened, and the document is not serialized or reloaded.
    /// Input admission limits must be enforced when loading the document. Output projection limits
    /// apply here. Source byte hashes, lengths and timestamps are not inferred from an in-memory model;
    /// requested chunk hashes describe the current semantic content.
    /// </remarks>
    public static OfficeDocumentReadResult ToOfficeDocumentReadResult(this PublisherDocument document,
        string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPublisherOptions? publisherOptions = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        ReaderOptions effective = DocumentReaderEngine.NormalizeOptions(readerOptions);
        ReaderPublisherOptions configured = (publisherOptions ?? new ReaderPublisherOptions()).Clone();
        if (effective.Projection == ReaderProjection.TextAndMetadata)
            throw new NotSupportedException("The Publisher Reader adapter does not support the TextAndMetadata projection.");
        string name = string.IsNullOrWhiteSpace(sourceName) ? "document.pub" : sourceName!.Trim();
        using var scope = ReaderReadScope.Enter(effective);
        OfficeDocumentReadResult result = new PublisherReadProjection(document, name, effective, configured, cancellationToken).Build();
        cancellationToken.ThrowIfCancellationRequested();
        DocumentReaderEngine.ApplyExternalSourceMetadata(result, DocumentReaderEngine.BuildSourceId(name),
            lastWriteUtc: null, lengthBytes: null, computeHashes: effective.ComputeHashes);
        cancellationToken.ThrowIfCancellationRequested();
        return ReaderReadScope.Complete(result);
    }
}
