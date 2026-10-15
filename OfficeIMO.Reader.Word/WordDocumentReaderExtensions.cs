using OfficeIMO.Word;

namespace OfficeIMO.Reader.Word;

/// <summary>Projects caller-owned Word models through the canonical Reader adapter.</summary>
public static class WordDocumentReaderExtensions {
    /// <summary>Reads text, structures and requested inspection evidence from an already loaded document.</summary>
    /// <param name="document">Caller-owned model, which remains open after projection.</param>
    /// <param name="sourceName">Logical citation name; omitted names use document.docx. No source file is reopened.</param>
    /// <param name="options">Snapshotted projection and output budgets. The default projection is Full.</param>
    /// <param name="formatOptions">Snapshotted Word projection policy.</param>
    /// <param name="cancellationToken">Cancellation for this projection.</param>
    /// <remarks>Input and package admission options cannot retroactively bound the model's earlier load. Source byte hashes, lengths and timestamps are not inferred from the logical name. Keep the model open to request subsequent native payload reads.</remarks>
    public static OfficeDocumentReadResult ReadDocument(this WordDocument document, string? sourceName = null,
        ReaderOptions? options = null, ReaderWordOptions? formatOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        ReaderOptions effective = DocumentReaderEngine.NormalizeOptions(options);
        ReaderWordOptions configured = WordReaderAdapter.Clone(formatOptions);
        using var scope = ReaderReadScope.Enter(effective);
        if (effective.Projection == ReaderProjection.TextAndMetadata && document.SourceFormatDescriptor.Encoding != OfficeFormatEncoding.OpenXml)
            throw new NotSupportedException("TextAndMetadata projection requires a modern XML Word source model.");
        string name = string.IsNullOrWhiteSpace(sourceName) ? "document.docx" : sourceName!.Trim();
        OfficeDocumentReadResult result = WordReaderAdapter.Project(document, name, effective, configured, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Source ??= new OfficeDocumentSource();
        result.Source.Path ??= name;
        return ReaderReadScope.Complete(result);
    }
}
