using OfficeIMO.PowerPoint;

namespace OfficeIMO.Reader.PowerPoint;

/// <summary>Projects caller-owned PowerPoint models through the canonical Reader adapter.</summary>
public static class PowerPointDocumentReaderExtensions {
    /// <summary>Reads slide text, structures and requested inspection evidence from an already loaded presentation.</summary>
    /// <param name="presentation">Caller-owned model, which remains open after projection.</param>
    /// <param name="sourceName">Logical citation name; omitted names use presentation.pptx. No source file is reopened.</param>
    /// <param name="options">Snapshotted projection and output budgets. The default projection is Full.</param>
    /// <param name="formatOptions">Snapshotted PowerPoint projection policy.</param>
    /// <param name="cancellationToken">Cancellation for this projection.</param>
    /// <remarks>Input and package admission options cannot retroactively bound the model's earlier load. Source byte hashes, lengths and timestamps are not inferred from the logical name. Keep the model open to request subsequent native payload reads.</remarks>
    public static OfficeDocumentReadResult ReadDocument(this PowerPointDocument presentation, string? sourceName = null,
        ReaderOptions? options = null, ReaderPowerPointOptions? formatOptions = null, CancellationToken cancellationToken = default) {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        cancellationToken.ThrowIfCancellationRequested();
        ReaderOptions effective = DocumentReaderEngine.NormalizeOptions(options);
        ReaderPowerPointOptions configured = PowerPointReaderAdapter.Clone(formatOptions);
        using var scope = ReaderReadScope.Enter(effective);
        if (effective.Projection == ReaderProjection.TextAndMetadata && presentation.SourceFormatDescriptor.Encoding != OfficeFormatEncoding.OpenXml)
            throw new NotSupportedException("TextAndMetadata projection requires a modern XML presentation source model.");
        string name = string.IsNullOrWhiteSpace(sourceName) ? "presentation.pptx" : sourceName!.Trim();
        OfficeDocumentReadResult result = PowerPointReaderAdapter.Project(presentation, name, effective, configured, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Source ??= new OfficeDocumentSource();
        result.Source.Path ??= name;
        return ReaderReadScope.Complete(result);
    }
}
