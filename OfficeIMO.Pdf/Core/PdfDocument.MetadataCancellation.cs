using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Updates metadata with cooperative cancellation, preserving the source document and its captured read limits.</summary>
    public PdfDocument UpdateMetadata(CancellationToken cancellationToken, string? title = null, string? author = null,
        string? subject = null, string? keywords = null) =>
        UpdateMetadata(title, author, subject, keywords, ReadOptions, cancellationToken);

    /// <summary>Replaces metadata with cooperative cancellation through parsing, rewriting and output adoption.</summary>
    public PdfDocument ReplaceMetadata(PdfMetadata metadata, CancellationToken cancellationToken) {
        Guard.NotNull(metadata, nameof(metadata));
        return ReplaceMetadata(metadata, ReadOptions, cancellationToken);
    }

    /// <summary>Synchronizes Info and XMP with cooperative cancellation, retaining the source document unchanged.</summary>
    public PdfDocument SynchronizeMetadata(CancellationToken cancellationToken, string? title = null, string? author = null,
        string? subject = null, string? keywords = null, bool createXmpMetadata = true) =>
        SynchronizeMetadata(title, author, subject, keywords, createXmpMetadata, ReadOptions, cancellationToken);

    /// <summary>Appends a metadata revision with cancellation through planning, serialization and output adoption.</summary>
    public PdfDocument AppendMetadataRevision(CancellationToken cancellationToken, string? title = null, string? author = null,
        string? subject = null, string? keywords = null, bool createXmpMetadata = false) =>
        AppendMetadataRevision(title, author, subject, keywords, ReadOptions, createXmpMetadata, cancellationToken);

    /// <summary>Attempts metadata updating through the planner-selected path; requested cancellation propagates to the caller.</summary>
    public PdfOperationResult<PdfDocument> UpdateMetadataResult(CancellationToken cancellationToken, string? title = null,
        string? author = null, string? subject = null, string? keywords = null, PdfLoadOptions? options = null) =>
        TryMutationOperation("Update metadata", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.UpdateMetadata,
            mode => mode == PdfMutationExecutionMode.AppendOnly
                ? AppendMetadataRevision(title, author, subject, keywords, options ?? ReadOptions, cancellationToken: cancellationToken)
                : UpdateMetadata(title, author, subject, keywords, options ?? ReadOptions, cancellationToken),
            options: options, cancellationToken: cancellationToken);

    /// <summary>Attempts metadata replacement with cancellation through planning, full rewriting and output adoption.</summary>
    public PdfOperationResult<PdfDocument> ReplaceMetadataResult(PdfMetadata metadata, CancellationToken cancellationToken,
        PdfLoadOptions? options = null) {
        Guard.NotNull(metadata, nameof(metadata));
        return TryMutationOperation("Replace metadata", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.UpdateMetadata,
            _ => ReplaceMetadata(metadata, options ?? ReadOptions, cancellationToken), options: options,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite, cancellationToken: cancellationToken);
    }

    /// <summary>Attempts Info/XMP synchronization with cancellation across the planner-selected rewrite or append-only path.</summary>
    public PdfOperationResult<PdfDocument> SynchronizeMetadataResult(CancellationToken cancellationToken, string? title = null,
        string? author = null, string? subject = null, string? keywords = null, bool createXmpMetadata = true,
        PdfLoadOptions? options = null) =>
        TryMutationOperation("Synchronize Info and XMP metadata", PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.SynchronizeMetadata,
            mode => mode == PdfMutationExecutionMode.AppendOnly
                ? ApplyMutation(input => PdfIncrementalUpdater.SynchronizeMetadata(input, title, author, subject, keywords,
                    options ?? ReadOptions, createXmpMetadata, cancellationToken), options ?? ReadOptions, cancellationToken: cancellationToken)
                : SynchronizeMetadata(title, author, subject, keywords, createXmpMetadata, options ?? ReadOptions, cancellationToken),
            options: options, cancellationToken: cancellationToken);

    /// <summary>Attempts a metadata-only incremental revision; requested cancellation propagates without changing the source.</summary>
    public PdfOperationResult<PdfDocument> AppendMetadataRevisionResult(CancellationToken cancellationToken, string? title = null,
        string? author = null, string? subject = null, string? keywords = null, PdfLoadOptions? options = null,
        bool createXmpMetadata = false) =>
        TryMutationOperation("Append metadata revision", PdfPreflightCapability.AppendMetadataRevision,
            PdfMutationOperation.UpdateMetadata,
            _ => AppendMetadataRevision(title, author, subject, keywords, options ?? ReadOptions, createXmpMetadata, cancellationToken),
            options: options, executionPreference: PdfMutationExecutionPreference.RequireAppendOnly, cancellationToken: cancellationToken);
}
