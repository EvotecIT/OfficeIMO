namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Reports read and rewrite capabilities for a PDF byte array without requiring the document to open successfully.
    /// This is useful for encrypted, malformed, or otherwise unsupported input that still needs a diagnostic report.
    /// </summary>
    public static PdfDocumentPreflight Preflight(byte[] pdf, PdfReadOptions? options = null) =>
        PdfInspector.Preflight(pdf, options);

    /// <summary>
    /// Reports read and rewrite capabilities for a PDF file without requiring the document to open successfully.
    /// This is useful for encrypted, malformed, or otherwise unsupported input that still needs a diagnostic report.
    /// </summary>
    public static PdfDocumentPreflight Preflight(string path, PdfReadOptions? options = null) =>
        PdfInspector.Preflight(path, options);

    /// <summary>
    /// Reports read and rewrite capabilities for a readable PDF stream without requiring the document to open successfully.
    /// The stream is consumed from its current position.
    /// </summary>
    public static PdfDocumentPreflight Preflight(Stream stream, PdfReadOptions? options = null) =>
        PdfInspector.Preflight(stream, options);

    /// <summary>
    /// Produces one consolidated health and capability report.
    /// Supply a compliance profile to include artifact readback readiness.
    /// </summary>
    public PdfAnalysisReport Analyze(PdfComplianceProfile complianceProfile = PdfComplianceProfile.None) {
        var snapshot = GetReadSnapshot();
        PdfDocumentInfo info = PdfInspector.Inspect(snapshot.Bytes, snapshot.Document);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(
            snapshot.Bytes,
            snapshot.Options,
            () => snapshot.Document);
        PdfDiagnosticReport diagnostics = PdfDiagnostics.Analyze(
            snapshot.Bytes,
            snapshot.Document,
            info,
            preflight);
        PdfOptimizationReport optimization = PdfDiagnostics.BuildOptimizationReport(diagnostics);
        PdfSignatureValidationReport signatures = PdfSignatureValidator.Validate(
            snapshot.Bytes,
            info.Security);
        PdfAppendOnlyMutationReport appendOnlyMutation = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(info.Security);
        PdfComplianceReadinessReport? compliance = complianceProfile == PdfComplianceProfile.None
            ? null
            : PdfComplianceAnalyzer.AssessReadback(complianceProfile, snapshot.Document, info);

        return new PdfAnalysisReport(
            info,
            preflight,
            diagnostics,
            optimization,
            signatures,
            appendOnlyMutation,
            snapshot.Document.RepairReport,
            compliance);
    }

    /// <summary>
    /// Inspects metadata, pages, annotations, fields, and catalog-level state.
    /// </summary>
    public PdfDocumentInfo Inspect(PdfReadOptions? options = null) {
        var snapshot = GetReadSnapshot(options);
        return PdfInspector.Inspect(snapshot.Bytes, snapshot.Document);
    }

    /// <summary>
    /// Reports read and rewrite capabilities for this PDF.
    /// </summary>
    public PdfDocumentPreflight Preflight(PdfReadOptions? options = null) {
        var snapshot = GetReadSnapshot(options);
        return PdfInspector.Preflight(
            snapshot.Bytes,
            snapshot.Options,
            () => snapshot.Document);
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
        return PdfSignatureValidator.Validate(GetBytesForOperation(), options ?? ReadOptions);
    }

    /// <summary>Validates signature structure and delegates CMS, trust, timestamp, and revocation policy to an optional provider.</summary>
    public PdfSignatureValidationReport ValidateSignatures(
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfReadOptions? options = null) {
        Guard.NotNull(cryptographyProvider, nameof(cryptographyProvider));
        return PdfSignatureValidator.Validate(GetBytesForOperation(), cryptographyProvider, options ?? ReadOptions);
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
        var snapshot = GetReadSnapshot(options);
        PdfDocumentInfo info = PdfInspector.Inspect(snapshot.Bytes, snapshot.Document);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(
            snapshot.Bytes,
            snapshot.Options,
            () => snapshot.Document);
        return PdfDiagnostics.Analyze(snapshot.Bytes, snapshot.Document, info, preflight);
    }

    /// <summary>Creates a bounded debugger projection of objects, revisions, pages, resources, and content operators.</summary>
    public PdfDebuggerReport Debug(PdfDebuggerOptions? options = null, PdfReadOptions? readOptions = null) {
        return PdfDebugger.Dump(GetBytesForOperation(), options, readOptions ?? ReadOptions);
    }

    /// <summary>
    /// Builds an optimization opportunity report for this document without modifying it.
    /// </summary>
    public PdfOptimizationReport AnalyzeOptimization(PdfReadOptions? options = null) {
        return PdfDiagnostics.BuildOptimizationReport(Diagnostics(options));
    }

    /// <summary>Applies dependency-free lossless optimization and returns the candidate with action and preservation reports.</summary>
    public PdfOptimizationActionResult Optimize(PdfOptimizationOptions? options = null) =>
        PdfOptimizer.Optimize(GetBytesForOperation(), options).WithReadOptions(ReadOptions);

    /// <summary>Applies a named deterministic lossless optimization profile.</summary>
    public PdfOptimizationActionResult Optimize(PdfOptimizationProfile profile) =>
        PdfOptimizer.Optimize(GetBytesForOperation(), profile).WithReadOptions(ReadOptions);

    /// <summary>
    /// Plans rectangle-based redaction impact without modifying the PDF.
    /// </summary>
    public PdfRedactionPlan PlanRedactions(IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return PdfRedactionPlanner.Plan(GetBytesForOperation(), areas, layoutOptions, options ?? ReadOptions);
    }

    /// <summary>Derives a reviewable redaction plan from literal text, regex, logical kinds, and form-field names.</summary>
    public PdfRedactionPlan SearchRedactions(PdfRedactionSearchOptions search, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) => PdfRedactionPlanner.Search(GetBytesForOperation(), search, layoutOptions, options ?? ReadOptions);

    /// <summary>
    /// Creates a new PDF with matching text objects and annotations removed from the supplied redaction areas.
    /// </summary>
    public PdfDocument ApplyRedactions(IEnumerable<PdfRedactionArea> areas, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) {
        return ApplyMutation(input => PdfRedactionApplier.Apply(input, areas, applyOptions, layoutOptions, options ?? ReadOptions), options);
    }

    /// <summary>Applies a reviewed redaction plan, including exact field removal for field-derived areas.</summary>
    public PdfDocument ApplyRedactions(PdfRedactionPlan plan, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) => ApplyMutation(input => PdfRedactionApplier.Apply(input, plan, applyOptions, layoutOptions, options ?? ReadOptions), options);

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

    /// <summary>Creates one PDF by merging all supplied documents in order through a single merge pass.</summary>
    public static PdfDocument Merge(params PdfDocument[] documents) =>
        Merge((IEnumerable<PdfDocument>)documents);

    /// <summary>Creates one PDF by merging all supplied documents in order through a single merge pass.</summary>
    public static PdfDocument Merge(IEnumerable<PdfDocument> documents) {
        Guard.NotNull(documents, nameof(documents));
        PdfDocument[] sources = documents.ToArray();
        if (sources.Length == 0) {
            throw new ArgumentException("At least one PDF document must be supplied.", nameof(documents));
        }

        if (sources.Any(static document => document is null)) {
            throw new ArgumentException("PDF documents cannot contain null entries.", nameof(documents));
        }

        byte[][] bytes = sources.Select(static document => document.GetBytesForOperation()).ToArray();
        PdfReadOptions[] readOptions = sources.Select(static document => document.ReadOptions).ToArray();
        return Open(PdfMerger.Merge(bytes, readOptions), sources[0].ReadOptions);
    }

    /// <summary>
    /// Creates a new PDF by merging this PDF with another loaded or generated PDF.
    /// </summary>
    public PdfDocument MergeWith(PdfDocument document) {
        Guard.NotNull(document, nameof(document));
        return ApplyMutation(input => PdfMerger.Merge(
            new[] { input, document.GetBytesForOperation() },
            new[] { ReadOptions, document.ReadOptions }));
    }

    /// <summary>
    /// Attempts to merge this PDF with another loaded or generated PDF, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryMergeWith(PdfDocument document, PdfReadOptions? options = null) {
        Guard.NotNull(document, nameof(document));
        return TryMutationOperation("Merge documents", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, _ => MergeWith(document), options: options);
    }

    /// <summary>
    /// Creates a new PDF by merging this PDF with another PDF byte payload.
    /// </summary>
    public PdfDocument MergeWith(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return ApplyMutation(input => PdfMerger.Merge(input, pdf));
    }

    /// <summary>
    /// Attempts to merge this PDF with another PDF byte payload, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryMergeWith(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return TryMutationOperation("Merge documents", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, _ => MergeWith(pdf), options: options);
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
        return TryMutationOperation("Merge documents", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, _ => MergeWith(path), options: options);
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
        return TryMutationOperation("Merge documents", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.MergeDocuments, _ => MergeWith(stream), options: options);
    }

    /// <summary>
    /// Creates a new PDF with visual annotation appearance streams painted into page content where supported.
    /// </summary>
    public PdfDocument FlattenVisualAnnotations() {
        return ApplyMutation(input => PdfAnnotationFlattener.FlattenVisualAnnotations(input, options: null, readOptions: ReadOptions));
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
        return ApplyMutation(
            input => PdfIncrementalUpdater.UpdateMetadata(
                input,
                title,
                author,
                subject,
                keywords,
                readOptions,
                createXmpMetadata),
            readOptions);
    }

    /// <summary>
    /// Appends an external-signature placeholder as an incremental revision for a later CMS, CAdES, or timestamp signature.
    /// </summary>
    public PdfExternalSignaturePreparation PrepareExternalSignature(PdfExternalSignatureOptions? signatureOptions = null) {
        return PdfIncrementalUpdater.PrepareExternalSignature(GetBytesForOperation(), signatureOptions);
    }

    /// <summary>Completes a persisted external-signature placeholder with detached CMS or timestamp bytes.</summary>
    public PdfDocument CompleteExternalSignature(byte[] signatureContents) {
        Guard.NotNull(signatureContents, nameof(signatureContents));
        return ApplyMutation(
            input => PdfIncrementalUpdater.ApplyExternalSignature(input, signatureContents),
            operationName: "CompleteExternalSignature");
    }

    /// <summary>Prepares, externally signs, and applies a PDF signature without placing key-storage logic in OfficeIMO.Pdf.</summary>
    public PdfExternalSignatureCompletion SignExternal(
        IPdfExternalSigner signer,
        PdfExternalSignatureOptions? signatureOptions = null) {
        Guard.NotNull(signer, nameof(signer));
        return PdfIncrementalUpdater.SignExternal(GetBytesForOperation(), signer, signatureOptions);
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
        return ApplyMutation(input => PdfMetadataEditor.UpdateMetadata(input, title, author, subject, keywords));
    }

    /// <summary>
    /// Creates a normalized full-rewrite PDF whose Info dictionary and XMP packet share the supplied common fields.
    /// Null values preserve existing values and empty strings clear them.
    /// </summary>
    public PdfDocument SynchronizeMetadata(
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        bool createXmpMetadata = true) {
        return ApplyMutation(input => PdfMetadataEditor.SynchronizeMetadata(
            input, title, author, subject, keywords, createXmpMetadata));
    }

    /// <summary>Attempts a full-rewrite Info/XMP synchronization and returns planner diagnostics when blocked.</summary>
    public PdfOperationResult<PdfDocument> TrySynchronizeMetadata(
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        bool createXmpMetadata = true,
        PdfReadOptions? options = null) {
        return TryMutationOperation(
            "Synchronize Info and XMP metadata",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.SynchronizeMetadata,
            _ => SynchronizeMetadata(title, author, subject, keywords, createXmpMetadata),
            options: options,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite);
    }

    /// <summary>Removes or quarantines active content and embedded payloads through a proven full rewrite.</summary>
    public PdfSanitizationResult Sanitize(PdfSanitizationOptions? options = null) {
        return PdfSanitizer.Sanitize(GetBytesForOperation(), options);
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
        return ApplyMutation(input => PdfMetadataEditor.ReplaceMetadata(input, metadata));
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
