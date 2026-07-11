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

    /// <summary>Chooses a full-rewrite, append-only, or blocked path for an existing-document mutation.</summary>
    public PdfMutationPlan PlanMutation(
        PdfMutationOperation operation,
        IEnumerable<string>? fieldNames = null,
        PdfReadOptions? options = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) {
        return PdfMutationPlanner.Plan(Preflight(options), operation, fieldNames, executionPreference);
    }

    /// <summary>
    /// Validates signature structure, byte ranges, and preservation markers for this PDF.
    /// </summary>
    public PdfSignatureValidationReport ValidateSignatures(PdfReadOptions? options = null) {
        return PdfSignatureValidator.Validate(Snapshot(), options ?? ReadOptions);
    }

    /// <summary>Validates signature structure and delegates CMS, trust, timestamp, and revocation policy to an optional provider.</summary>
    public PdfSignatureValidationReport ValidateSignatures(
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfReadOptions? options = null) {
        Guard.NotNull(cryptographyProvider, nameof(cryptographyProvider));
        return PdfSignatureValidator.Validate(Snapshot(), cryptographyProvider, options ?? ReadOptions);
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
        return TryMutationOperation(
            "Apply redactions",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.Redact,
            _ => ApplyRedactions(areas, applyOptions, layoutOptions, options),
            options: options ?? ReadOptions);
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

    internal PdfOperationResult<T> TryMutationOperation<T>(
        string operationName,
        PdfPreflightCapability capability,
        PdfMutationOperation mutationOperation,
        Func<PdfMutationExecutionMode, T> operation,
        IEnumerable<string>? fieldNames = null,
        PdfReadOptions? options = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) where T : class {
        Guard.NotNullOrWhiteSpace(operationName, nameof(operationName));
        Guard.NotNull(operation, nameof(operation));

        PdfMutationPlan plan = PlanMutation(mutationOperation, fieldNames, options, executionPreference);
        if (!plan.CanExecute) {
            return PdfOperationResult<T>.MutationBlocked(operationName, capability, plan);
        }

        try {
            return PdfOperationResult<T>.MutationSuccess(operationName, capability, plan, operation(plan.ExecutionMode));
        } catch (Exception ex) {
            return PdfOperationResult<T>.MutationFailed(operationName, capability, plan, ex);
        }
    }

    internal PdfOperationResult<T> TryMutationOperation<T>(
        string operationName,
        PdfPreflightCapability capability,
        PdfMutationOperation mutationOperation,
        Func<T> operation,
        PdfReadOptions? options = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) where T : class {
        Guard.NotNull(operation, nameof(operation));
        return TryMutationOperation(
            operationName,
            capability,
            mutationOperation,
            _ => operation(),
            options: options,
            executionPreference: executionPreference);
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
        return TryMutationOperation("Merge documents", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, _ => MergeWith(document), options: options);
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
        return TryMutationOperation("Merge documents", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, _ => MergeWith(pdf), options: options);
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
        return TryMutationOperation("Merge documents", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, _ => MergeWith(path), options: options);
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
        return TryMutationOperation("Merge documents", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, _ => MergeWith(stream), options: options);
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
        return TryMutationOperation("Flatten visual annotations", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyAnnotations, _ => FlattenVisualAnnotations(), options: options);
    }

    /// <summary>
    /// Appends a metadata-only incremental revision without rewriting the existing PDF bytes.
    /// </summary>
    public PdfDocument AppendMetadataRevision(
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        bool createXmpMetadata = false) {
        return AppendMetadataRevision(title, author, subject, keywords, ReadOptions, createXmpMetadata);
    }

    /// <summary>
    /// Attempts to append a metadata-only incremental revision, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryAppendMetadataRevision(
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        PdfReadOptions? options = null,
        bool createXmpMetadata = false) {
        return TryMutationOperation(
            "Append metadata revision",
            PdfPreflightCapability.AppendMetadataRevision,
            PdfMutationOperation.UpdateMetadata,
            _ => AppendMetadataRevision(title, author, subject, keywords, options ?? ReadOptions, createXmpMetadata),
            options: options,
            executionPreference: PdfMutationExecutionPreference.RequireAppendOnly);
    }

    private PdfDocument AppendMetadataRevision(
        string? title,
        string? author,
        string? subject,
        string? keywords,
        PdfReadOptions? readOptions,
        bool createXmpMetadata = false) {
        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(Snapshot(), title, author, subject, keywords, readOptions, createXmpMetadata);
        return FromBytes(updated, readOptions);
    }

    /// <summary>
    /// Appends an external-signature placeholder as an incremental revision for a later CMS, CAdES, or timestamp signature.
    /// </summary>
    public PdfExternalSignaturePreparation PrepareExternalSignature(PdfExternalSignatureOptions? signatureOptions = null) {
        return PdfIncrementalUpdater.PrepareExternalSignature(Snapshot(), signatureOptions);
    }

    /// <summary>Prepares, externally signs, and applies a PDF signature without placing key-storage logic in OfficeIMO.Pdf.</summary>
    public PdfExternalSignatureCompletion SignExternal(
        IPdfExternalSigner signer,
        PdfExternalSignatureOptions? signatureOptions = null) {
        Guard.NotNull(signer, nameof(signer));
        return PdfIncrementalUpdater.SignExternal(Snapshot(), signer, signatureOptions);
    }

    /// <summary>
    /// Attempts to append an external-signature placeholder revision, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfExternalSignaturePreparation> TryPrepareExternalSignature(PdfExternalSignatureOptions? signatureOptions = null, PdfReadOptions? options = null) {
        return TryMutationOperation(
            "Prepare external signature",
            PdfPreflightCapability.PrepareExternalSignatureRevision,
            PdfMutationOperation.PrepareExternalSignature,
            _ => PrepareExternalSignature(signatureOptions),
            options: options);
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
        return TryMutationOperation(
            "Update metadata",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.UpdateMetadata,
            mode => mode == PdfMutationExecutionMode.AppendOnly
                ? AppendMetadataRevision(title, author, subject, keywords, options ?? ReadOptions)
                : UpdateMetadata(title, author, subject, keywords),
            options: options);
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
        return TryMutationOperation(
            "Replace metadata",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.UpdateMetadata,
            _ => ReplaceMetadata(metadata),
            options: options,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite);
    }
}
