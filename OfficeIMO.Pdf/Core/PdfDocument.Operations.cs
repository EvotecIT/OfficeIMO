namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Inspects metadata, pages, annotations, fields, and catalog-level state.
    /// </summary>
    public PdfDocumentInfo Inspect(PdfReadOptions? options = null) {
        return PdfInspector.Inspect(Snapshot(), options);
    }

    /// <summary>
    /// Reports read and rewrite capabilities for this PDF.
    /// </summary>
    public PdfDocumentPreflight Preflight(PdfReadOptions? options = null) {
        return PdfInspector.Preflight(Snapshot(), options);
    }

    internal PdfOperationResult<T> TryOperation<T>(
        string operationName,
        PdfPreflightCapability capability,
        Func<T> operation,
        PdfReadOptions? options = null) where T : class {
        Guard.NotNullOrWhiteSpace(operationName, nameof(operationName));
        Guard.NotNull(operation, nameof(operation));

        PdfDocumentPreflight preflight = Preflight(options);
        if (!preflight.Can(capability)) {
            return PdfOperationResult<T>.Blocked(operationName, capability, preflight);
        }

        try {
            return PdfOperationResult<T>.Success(operationName, capability, preflight, operation());
        } catch (Exception ex) {
            return PdfOperationResult<T>.Failed(operationName, capability, preflight, ex);
        }
    }

    /// <summary>
    /// Creates a new PDF by merging this PDF with another loaded or generated PDF.
    /// </summary>
    public PdfDocument MergeWith(PdfDocument document) {
        Guard.NotNull(document, nameof(document));
        return FromBytes(PdfMerger.Merge(Snapshot(), document.Snapshot()));
    }

    /// <summary>
    /// Attempts to merge this PDF with another loaded or generated PDF, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryMergeWith(PdfDocument document, PdfReadOptions? options = null) {
        Guard.NotNull(document, nameof(document));
        return TryOperation("Merge documents", PdfPreflightCapability.ManipulatePages, () => MergeWith(document), options);
    }

    /// <summary>
    /// Creates a new PDF by merging this PDF with another PDF byte payload.
    /// </summary>
    public PdfDocument MergeWith(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return FromBytes(PdfMerger.Merge(Snapshot(), pdf));
    }

    /// <summary>
    /// Attempts to merge this PDF with another PDF byte payload, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryMergeWith(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return TryOperation("Merge documents", PdfPreflightCapability.ManipulatePages, () => MergeWith(pdf), options);
    }

    /// <summary>
    /// Creates a new PDF by merging this PDF with another PDF file.
    /// </summary>
    public PdfDocument MergeWith(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return MergeWith(File.ReadAllBytes(path));
    }

    /// <summary>
    /// Attempts to merge this PDF with another PDF file, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryMergeWith(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return TryOperation("Merge documents", PdfPreflightCapability.ManipulatePages, () => MergeWith(path), options);
    }

    /// <summary>
    /// Creates a new PDF by merging this PDF with another readable PDF stream.
    /// </summary>
    public PdfDocument MergeWith(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return MergeWith(buffer.ToArray());
    }

    /// <summary>
    /// Attempts to merge this PDF with another readable PDF stream, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryMergeWith(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        return TryOperation("Merge documents", PdfPreflightCapability.ManipulatePages, () => MergeWith(stream), options);
    }

    /// <summary>
    /// Creates a new PDF with updated metadata. Null values preserve existing fields; empty strings clear fields.
    /// </summary>
    public PdfDocument UpdateMetadata(string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        return FromBytes(PdfMetadataEditor.UpdateMetadata(Snapshot(), title, author, subject, keywords));
    }

    /// <summary>
    /// Attempts to create a new PDF with updated metadata, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryUpdateMetadata(string? title = null, string? author = null, string? subject = null, string? keywords = null, PdfReadOptions? options = null) {
        return TryOperation("Update metadata", PdfPreflightCapability.ManipulatePages, () => UpdateMetadata(title, author, subject, keywords), options);
    }

    /// <summary>
    /// Creates a new PDF with exactly the supplied metadata.
    /// </summary>
    public PdfDocument ReplaceMetadata(PdfMetadata metadata) {
        Guard.NotNull(metadata, nameof(metadata));
        return FromBytes(PdfMetadataEditor.ReplaceMetadata(Snapshot(), metadata));
    }

    /// <summary>
    /// Attempts to create a new PDF with exactly the supplied metadata, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryReplaceMetadata(PdfMetadata metadata, PdfReadOptions? options = null) {
        Guard.NotNull(metadata, nameof(metadata));
        return TryOperation("Replace metadata", PdfPreflightCapability.ManipulatePages, () => ReplaceMetadata(metadata), options);
    }
}
