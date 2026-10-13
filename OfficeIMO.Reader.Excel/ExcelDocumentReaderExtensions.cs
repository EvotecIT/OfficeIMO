using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

/// <summary>Projects caller-owned Excel models through the canonical Reader adapter.</summary>
public static class ExcelDocumentReaderExtensions {
    /// <summary>Reads table text, structures and requested inspection evidence from an already loaded workbook.</summary>
    /// <param name="document">Caller-owned model, which remains open after projection.</param>
    /// <param name="sourceName">Logical citation name; omitted names use workbook.xlsx. No source file is reopened.</param>
    /// <param name="options">Snapshotted projection and output budgets. The default projection is Full.</param>
    /// <param name="formatOptions">Snapshotted Excel projection policy, including its native read options.</param>
    /// <param name="cancellationToken">Cancellation for this projection.</param>
    /// <remarks>Input and package admission options cannot retroactively bound the model's earlier load. Source byte hashes, lengths and timestamps are not inferred from the logical name. Keep the model open to request subsequent native payload reads.</remarks>
    public static OfficeDocumentReadResult ReadDocument(this ExcelDocument document, string? sourceName = null,
        ReaderOptions? options = null, ReaderExcelOptions? formatOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        ReaderOptions effective = DocumentReaderEngine.NormalizeOptions(options);
        ReaderExcelOptions configured = ExcelReaderAdapter.Clone(formatOptions);
        using var scope = ReaderReadScope.Enter(effective);
        if (effective.Projection == ReaderProjection.TextAndMetadata && document.SourceFormatDescriptor.Encoding != OfficeFormatEncoding.OpenXml)
            throw new NotSupportedException("TextAndMetadata projection requires a modern XML workbook source model.");
        string name = string.IsNullOrWhiteSpace(sourceName) ? "workbook.xlsx" : sourceName!.Trim();
        OfficeDocumentReadResult result = ExcelReaderAdapter.Project(document, name, effective, configured, cancellationToken,
            capabilityId: OfficeDocumentReaderBuilderExcelExtensions.HandlerId);
        cancellationToken.ThrowIfCancellationRequested();
        result.Source ??= new OfficeDocumentSource();
        result.Source.Path ??= name;
        return ReaderReadScope.Complete(result);
    }
}
