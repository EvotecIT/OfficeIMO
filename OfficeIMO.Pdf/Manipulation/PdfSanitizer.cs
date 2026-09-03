namespace OfficeIMO.Pdf;

/// <summary>Removes or quarantines active content and embedded payloads through a proven full rewrite.</summary>
internal static partial class PdfSanitizer {
    internal static PdfSanitizationReport Inspect(byte[] pdf, PdfSanitizationOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfSanitizationOptions policy = options ?? new PdfSanitizationOptions();
        IReadOnlyList<PdfSanitizationFinding> findings = Analyze(pdf, policy, readOptions);
        return BuildReport(pdf, policy, readOptions, findings);
    }

    /// <summary>Returns the forbidden-content inventory that the supplied policy would remove.</summary>
    public static IReadOnlyList<PdfSanitizationFinding> Analyze(byte[] pdf, PdfSanitizationOptions? options = null) {
        return Analyze(pdf, options, readOptions: null);
    }

    internal static IReadOnlyList<PdfSanitizationFinding> Analyze(byte[] pdf, PdfSanitizationOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfSanitizationOptions policy = options ?? new PdfSanitizationOptions();
        System.Threading.CancellationToken cancellationToken = policy.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        var parsed = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadLimits readLimits = readOptions?.Limits ?? new PdfReadLimits();
        return Scan(parsed.Map, policy, readLimits);
    }

    /// <summary>
    /// Produces a normalized PDF with forbidden actions, unsafe URI targets, rich media, and embedded payloads removed.
    /// Quarantine mode returns decoded attachments to the caller but never writes them to disk.
    /// </summary>
    public static PdfSanitizationResult Sanitize(byte[] pdf, PdfSanitizationOptions? options = null) {
        return Sanitize(pdf, options, readOptions: null);
    }

    internal static PdfSanitizationResult Sanitize(byte[] pdf, PdfSanitizationOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfSanitizationOptions policy = options ?? new PdfSanitizationOptions();
        if (policy.MaximumOutputBytes <= 0L) {
            throw new ArgumentOutOfRangeException(nameof(options), "Maximum sanitized output bytes must be positive.");
        }
        System.Threading.CancellationToken cancellationToken = policy.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        var (plan, sourceDocument) = PdfMutationPlanner.RequireFullRewriteDocument(
            pdf,
            PdfMutationOperation.Sanitize,
            readOptions,
            cancellationToken: cancellationToken);
        PdfSanitizationReport before = BuildReport(pdf, policy, readOptions, Analyze(pdf, policy, readOptions));
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<PdfExtractedAttachment> quarantined;
        if (policy.ShouldRemoveEmbeddedFiles && policy.EmbeddedFiles == PdfEmbeddedFileSanitizationMode.Quarantine) {
            quarantined = sourceDocument.ExtractAttachments(cancellationToken);
        } else {
            quarantined = Array.Empty<PdfExtractedAttachment>();
        }
        cancellationToken.ThrowIfCancellationRequested();

        PdfReadLimits readLimits = readOptions?.Limits ?? new PdfReadLimits();
        int maximumActionDepth = readLimits.MaxObjectNestingDepth;
        int maximumActionNodes = readLimits.MaxIndirectObjects;
        byte[] sanitized;
        try {
            sanitized = PdfDocumentObjectGraphRewriter.Rewrite(
                pdf,
                sourceReadOptions: readOptions,
                outputEncryption: null,
                (objects, security) => {
                    cancellationToken.ThrowIfCancellationRequested();
                    SanitizeObjectGraph(objects, security, policy, maximumActionDepth, maximumActionNodes);
                    return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value)
                        ? security.InfoObjectNumber
                        : null;
                },
                maximumOutputBytes: policy.MaximumOutputBytes,
                cancellationToken: cancellationToken);
        } catch (InvalidDataException exception) when (PdfDocumentObjectGraphRewriter.IsOutputLimitExceeded(exception)) {
            throw new InvalidOperationException(
                $"The sanitized PDF exceeded the configured {policy.MaximumOutputBytes:N0}-byte output limit while it was being serialized.",
                exception);
        }
        cancellationToken.ThrowIfCancellationRequested();
        PdfLoadOptions rewrittenReadOptions = PdfLoadOptions.WithMinimumInputBytes(readOptions, sanitized.LongLength);
        PdfSanitizationReport remaining = BuildReport(
            sanitized,
            policy,
            rewrittenReadOptions,
            Analyze(sanitized, policy, rewrittenReadOptions));
        cancellationToken.ThrowIfCancellationRequested();
        int remainingCount = Math.Max(remaining.TotalCount, remaining.CategoryCounts.Total);
        if (remainingCount > 0) {
            throw new InvalidOperationException(
                "PDF sanitization post-save validation found " + remainingCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " forbidden item(s); the artifact was not returned.");
        }

        PdfDocumentInfo originalInfo = PdfInspector.Inspect(pdf, sourceDocument, cancellationToken);
        var preservationOptions = new PdfRewritePreservationOptions {
            OriginalReadOptions = readOptions,
            RewrittenReadOptions = rewrittenReadOptions,
            PreserveMetadata = !policy.ShouldRemoveUserMetadata,
            PreserveOutlines = !policy.ShouldRemoveBookmarks,
            PreserveLinkAnnotations = true,
            PreserveAnnotations = true,
            PreserveEmbeddedFiles = !policy.ShouldRemoveEmbeddedFiles,
            PreserveXmpMetadata = !policy.ShouldRemoveUserMetadata,
            PreserveOptionalContent = !policy.ShouldRemoveOptionalContent,
            PreserveCatalogViewSettings = !policy.ShouldRemoveBookmarks &&
                !policy.ShouldRemoveEmbeddedFiles &&
                !policy.ShouldRemoveOptionalContent,
            PreserveCatalogActions = true,
            PreservePageActions = true,
            PreserveOpenAction = true,
            PreserveFormWidgetActions = true,
            FilterActionsByPreservedTypes = true,
            PreserveRevisionStructure = false,
            PreserveSecurityState = !sourceDocument.Security.HasEncryption
        };
        if (policy.ShouldRemoveEmbeddedFiles && (policy.ContentKindsToRemove.HasValue || policy.RemoveRichMedia)) {
            preservationOptions.ExcludedAnnotationSubtypes.Add("FileAttachment");
        }
        if (policy.ShouldRemoveCommentsAndMarkup) {
            foreach (string subtype in originalInfo.AnnotationSubtypeCounts.Keys) {
                if (policy.ShouldRemoveCommentAnnotation(subtype)) preservationOptions.ExcludedAnnotationSubtypes.Add(subtype);
            }
        } else if (!policy.ContentKindsToRemove.HasValue && policy.RemoveRichMedia) {
            foreach (string subtype in RichAnnotationSubtypes) preservationOptions.ExcludedAnnotationSubtypes.Add(subtype);
        }
        for (int i = 0; i < originalInfo.LinkAnnotations.Count; i++) {
            string? uri = originalInfo.LinkAnnotations[i].Uri;
            if (uri is not null && policy.ShouldRemoveAction("URI", uri)) preservationOptions.ExcludedLinkAnnotationUris.Add(uri);
        }
        AddPolicyRetainedActionTypes(originalInfo, policy, preservationOptions.PreservedActionTypes);
        for (int i = 0; i < before.Findings.Count; i++) {
            if (before.Findings[i].ActionKind == PdfSanitizationActionKind.Uri) {
                preservationOptions.ExcludedActionUris.Add(before.Findings[i].Detail);
            }
        }
        PdfRewritePreservationReport preservation = PdfRewritePreservation.AssertPreserved(
            pdf,
            sanitized,
            preservationOptions,
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        return new PdfSanitizationResult(sanitized, plan, preservation, before, remaining, quarantined, rewrittenReadOptions);
    }

    internal static void AddPolicyRetainedActionTypes(PdfDocumentInfo info, PdfSanitizationOptions policy, ISet<string> preservedActionTypes) {
        for (int i = 0; i < info.CatalogActions.Count; i++) AddPolicyRetainedActionType(info.CatalogActions[i].ActionType, policy, preservedActionTypes);
        for (int i = 0; i < info.Pages.Count; i++) {
            IReadOnlyList<PdfPageAction> actions = info.Pages[i].PageActions;
            for (int j = 0; j < actions.Count; j++) AddPolicyRetainedActionType(actions[j].ActionType, policy, preservedActionTypes);
        }
        for (int fieldIndex = 0; fieldIndex < info.FormFields.Count; fieldIndex++) {
            IReadOnlyList<PdfFormWidget> widgets = info.FormFields[fieldIndex].Widgets;
            for (int widgetIndex = 0; widgetIndex < widgets.Count; widgetIndex++) {
                IReadOnlyList<PdfFormWidgetAction> actions = widgets[widgetIndex].Actions;
                for (int actionIndex = 0; actionIndex < actions.Count; actionIndex++) {
                    AddPolicyRetainedActionType(actions[actionIndex].ActionType, policy, preservedActionTypes);
                }
            }
        }
        if (info.OpenAction is not null) AddPolicyRetainedActionType(info.OpenAction.ActionType, policy, preservedActionTypes);
        foreach (string actionType in policy.AllowedActionTypes) preservedActionTypes.Add(actionType);
    }

    private static void AddPolicyRetainedActionType(string actionType, PdfSanitizationOptions policy, ISet<string> preservedActionTypes) {
        if (!policy.ShouldRemoveAction(actionType)) preservedActionTypes.Add(actionType);
    }

    /// <summary>Sanitizes a PDF from the current position of a readable stream.</summary>
    public static PdfSanitizationResult Sanitize(Stream stream, PdfSanitizationOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Sanitize(buffer.ToArray(), options);
    }

    /// <summary>Sanitizes a PDF file and returns the result without writing output automatically.</summary>
    public static PdfSanitizationResult Sanitize(string inputPath, PdfSanitizationOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return Sanitize(File.ReadAllBytes(inputPath), options);
    }
}
