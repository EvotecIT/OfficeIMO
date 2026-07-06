namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Inspects metadata, pages, annotations, fields, and catalog-level state.
    /// </summary>
    public PdfDocumentInfo Inspect(PdfReadOptions? options = null) {
        return PdfInspector.Inspect(Snapshot(), options ?? ReadOptions);
    }

    /// <summary>
    /// Reports read and rewrite capabilities for this PDF.
    /// </summary>
    public PdfDocumentPreflight Preflight(PdfReadOptions? options = null) {
        return PdfInspector.Preflight(Snapshot(), options ?? ReadOptions);
    }

    /// <summary>
    /// Validates signature structure, byte ranges, and preservation markers for this PDF.
    /// </summary>
    public PdfSignatureValidationReport ValidateSignatures(PdfReadOptions? options = null) {
        return PdfSignatureValidator.Validate(Snapshot(), options ?? ReadOptions);
    }

    /// <summary>
    /// Analyzes which append-only mutation actions OfficeIMO.Pdf can safely attempt for this PDF.
    /// </summary>
    public PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(PdfReadOptions? options = null) {
        return PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(Inspect(options).Security);
    }

    /// <summary>
    /// Builds a combined PDF diagnostic report for this document.
    /// </summary>
    public PdfDiagnosticReport Diagnostics(PdfReadOptions? options = null) {
        return PdfDiagnostics.Analyze(Snapshot(), options ?? ReadOptions);
    }

    /// <summary>
    /// Builds an optimization opportunity report for this document without modifying it.
    /// </summary>
    public PdfOptimizationReport AnalyzeOptimization(PdfReadOptions? options = null) {
        return PdfDiagnostics.AnalyzeOptimization(Snapshot(), options ?? ReadOptions);
    }

    /// <summary>
    /// Plans rectangle-based redaction impact without modifying the PDF.
    /// </summary>
    public PdfRedactionPlan PlanRedactions(IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return PdfRedactionPlanner.Plan(Snapshot(), areas, layoutOptions, options ?? ReadOptions);
    }

    /// <summary>
    /// Creates a new PDF with matching text objects and annotations removed from the supplied redaction areas.
    /// </summary>
    public PdfDocument ApplyRedactions(IEnumerable<PdfRedactionArea> areas, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return FromBytes(PdfRedactionApplier.Apply(Snapshot(), areas, applyOptions, layoutOptions, options ?? ReadOptions));
    }

    /// <summary>
    /// Attempts to apply rectangle-based redactions, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryApplyRedactions(IEnumerable<PdfRedactionArea> areas, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        Guard.NotNull(areas, nameof(areas));
        return TryOperation("Apply redactions", PdfPreflightCapability.ManipulatePages, () => ApplyRedactions(areas, applyOptions, layoutOptions, options), options ?? ReadOptions);
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
    /// Creates a new PDF with visual annotation appearance streams painted into page content where supported.
    /// </summary>
    public PdfDocument FlattenVisualAnnotations() {
        return FromBytes(PdfAnnotationFlattener.FlattenVisualAnnotations(Snapshot()));
    }

    /// <summary>
    /// Attempts to flatten visual annotation appearance streams, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryFlattenVisualAnnotations(PdfReadOptions? options = null) {
        return TryOperation("Flatten visual annotations", PdfPreflightCapability.ManipulatePages, FlattenVisualAnnotations, options);
    }

    /// <summary>
    /// Appends a metadata-only incremental revision without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendMetadataRevision(string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        return FromBytes(PdfIncrementalUpdater.UpdateMetadata(Snapshot(), title, author, subject, keywords));
    }

    /// <summary>
    /// Attempts to append a metadata-only incremental revision, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppendMetadataRevision(string? title = null, string? author = null, string? subject = null, string? keywords = null, PdfReadOptions? options = null) {
        return TryOperation("Append metadata revision", PdfPreflightCapability.AppendMetadataRevision, () => AppendMetadataRevision(title, author, subject, keywords), options);
    }

    /// <summary>
    /// Appends an external-signature placeholder as an incremental revision for a later CMS, CAdES, or timestamp signature.
    /// </summary>
    public PdfExternalSignaturePreparation PrepareExternalSignature(PdfExternalSignatureOptions? signatureOptions = null) {
        return PdfIncrementalUpdater.PrepareExternalSignature(Snapshot(), signatureOptions);
    }

    /// <summary>
    /// Attempts to append an external-signature placeholder revision, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfExternalSignaturePreparation> TryPrepareExternalSignature(PdfExternalSignatureOptions? signatureOptions = null, PdfReadOptions? options = null) {
        return TryOperation("Prepare external signature", PdfPreflightCapability.PrepareExternalSignatureRevision, () => PrepareExternalSignature(signatureOptions), options);
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
