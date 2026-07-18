namespace OfficeIMO.Pdf;

/// <summary>
/// Result of a source-document to PDF conversion, pairing the generated PDF document with a snapshot of conversion diagnostics.
/// </summary>
public sealed partial class PdfDocumentConversionResult {
    private readonly PdfConversionReport _sourceReport;

    /// <summary>
    /// Creates a conversion result from a generated PDF document and the conversion report populated during export.
    /// </summary>
    public PdfDocumentConversionResult(PdfDocument document, PdfConversionReport conversionReport) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(conversionReport, nameof(conversionReport));

        _sourceReport = conversionReport;
        Value = document;
        Report = SnapshotReport(conversionReport);
    }

    private PdfDocumentConversionResult(PdfDocument document, PdfConversionReport sourceReport, PdfConversionReport conversionReportSnapshot) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(sourceReport, nameof(sourceReport));
        Guard.NotNull(conversionReportSnapshot, nameof(conversionReportSnapshot));

        _sourceReport = sourceReport;
        Value = document;
        Report = SnapshotReport(conversionReportSnapshot);
    }

    /// <summary>The generated PDF document, ready for fluent OfficeIMO.Pdf processing.</summary>
    public PdfDocument Value { get; }

    /// <summary>Create/open and post-processing evidence accumulated by the generated PDF document.</summary>
    public PdfPipelineReport Pipeline => Value.Pipeline;

    /// <summary>Snapshot of conversion warnings captured when the result was created and refreshed after save-time diagnostics run.</summary>
    public PdfConversionReport Report { get; }

    /// <summary>Warnings captured during conversion.</summary>
    public IReadOnlyList<PdfConversionWarning> Warnings => Report.Warnings;

    /// <summary>True when conversion produced at least one warning.</summary>
    public bool HasWarnings => Report.HasWarnings;

    /// <summary>True when conversion reported an approximation, omission, or error.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Counts grouped conversion diagnostics captured with this result.</summary>
    public PdfConversionReportSummary Summary => Report.Summarize();

    /// <summary>Returns the generated PDF document.</summary>
    public PdfDocument RequireValue() => Value;

    /// <summary>Returns the generated PDF document only when conversion reported no possible content loss.</summary>
    public PdfDocument RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }

    /// <summary>
    /// Returns a new conversion result with the supplied PDF document while preserving the captured conversion diagnostics.
    /// </summary>
    public PdfDocumentConversionResult WithValue(PdfDocument value) {
        return new PdfDocumentConversionResult(value, _sourceReport, Report);
    }

    /// <summary>
    /// Returns a new conversion result that prepends adapter or source-projection warnings while retaining
    /// the original converter report as a live linked source for save-time diagnostics.
    /// </summary>
    public PdfDocumentConversionResult WithAdditionalWarnings(IEnumerable<PdfConversionWarning> warnings) {
        Guard.NotNull(warnings, nameof(warnings));
        var combinedSource = new PdfConversionReport();
        combinedSource.AddRange(warnings);
        combinedSource.LinkReport(_sourceReport);
        return new PdfDocumentConversionResult(Value, combinedSource, combinedSource);
    }

    /// <summary>
    /// Applies PDF post-processing while preserving the captured conversion diagnostics in the returned result.
    /// </summary>
    public PdfDocumentConversionResult Process(Func<PdfDocument, PdfDocument> process) {
        Guard.NotNull(process, nameof(process));
        return WithValue(process(Value));
    }

    /// <summary>
    /// Builds a reusable proof snapshot for the generated PDF and captured conversion diagnostics.
    /// </summary>
    public PdfConversionProofReport AssessProof(PdfConversionProofOptions? options = null) {
        options ??= new PdfConversionProofOptions();

        var issues = new List<PdfConversionProofIssue>();
        PdfDocumentInfo? documentInfo = null;
        PdfLogicalDocument? logicalDocument = null;
        string extractedText = string.Empty;
        IReadOnlyList<string> logicalSignals = Array.Empty<string>();
        long artifactByteCount = 0;
        string artifactSha256 = string.Empty;
        byte[]? pdfBytes = null;

        bool shouldReadGeneratedPdf = ShouldReadGeneratedPdf(options);
        bool shouldCaptureArtifactHash = options.IncludeArtifactHash || !string.IsNullOrWhiteSpace(options.RequiredArtifactSha256);
        if (shouldReadGeneratedPdf || shouldCaptureArtifactHash) {
            try {
                pdfBytes = ToBytes();
                if (shouldCaptureArtifactHash) {
                    artifactByteCount = pdfBytes.LongLength;
                    artifactSha256 = ComputeSha256Hex(pdfBytes);
                }

                if (shouldReadGeneratedPdf) {
                    PdfReadDocument readDocument = PdfReadDocument.Open(pdfBytes);
                    extractedText = readDocument.ExtractText();
                    documentInfo = PdfInspector.Inspect(pdfBytes);
                    if (options.RequiredLogicalSignals.Count > 0) {
                        logicalDocument = PdfLogicalDocument.From(readDocument);
                        logicalSignals = BuildLogicalSignals(documentInfo, logicalDocument, extractedText);
                    }
                }
            } catch (Exception ex) {
                issues.Add(new PdfConversionProofIssue(
                    shouldReadGeneratedPdf ? "ReadablePdf" : "ArtifactSha256",
                    shouldReadGeneratedPdf ? "OfficeIMO.Pdf reader can inspect and extract text" : "generated PDF artifact hash",
                    ex.GetType().Name + ": " + ex.Message));
            }
        }

        if (pdfBytes is null && ShouldCaptureSaveDiagnostics(options)) {
            try {
                pdfBytes = ToBytes();
            } catch (Exception ex) {
                issues.Add(new PdfConversionProofIssue(
                    "SaveDiagnostics",
                    "generated PDF save diagnostics",
                    ex.GetType().Name + ": " + ex.Message));
            }
        }

        for (int i = 0; i < options.RequiredTextMarkers.Count; i++) {
            string marker = options.RequiredTextMarkers[i];
            if (extractedText.IndexOf(marker, StringComparison.Ordinal) < 0) {
                issues.Add(new PdfConversionProofIssue("TextMarker", marker, "missing"));
            }
        }

        for (int i = 0; i < options.RequiredLogicalSignals.Count; i++) {
            string signal = options.RequiredLogicalSignals[i];
            if (!ContainsLogicalSignal(logicalSignals, signal)) {
                issues.Add(new PdfConversionProofIssue("LogicalSignal", signal, "missing"));
            }
        }

        if (options.RequiredPageCount.HasValue && documentInfo is not null && documentInfo.PageCount != options.RequiredPageCount.Value) {
            issues.Add(new PdfConversionProofIssue(
                "PageCount",
                options.RequiredPageCount.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                documentInfo.PageCount.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }

        if (options.RequiredPageWidth.HasValue && options.RequiredPageHeight.HasValue && documentInfo is not null) {
            AddPageSizeIssues(issues, documentInfo, options);
        }

        if (documentInfo is not null) {
            AddMetadataIssues(issues, documentInfo.Metadata, options);
            AddOutlineTitleIssues(issues, documentInfo, options);
            AddLinkUriIssues(issues, documentInfo, options);
            AddFormFieldNameIssues(issues, documentInfo, options);
            AddNamedDestinationNameIssues(issues, documentInfo, options);
            AddPageLabelRangeIssues(issues, documentInfo, options);
            AddAttachmentFileNameIssues(issues, documentInfo, options);
            AddOutputIntentIssues(issues, documentInfo, options);
            AddCatalogMetadataIssues(issues, documentInfo, options);
            AddOpenActionIssues(issues, documentInfo, options);
            AddViewerPreferenceIssues(issues, documentInfo, options);
            AddOptionalContentIssues(issues, documentInfo, options);
            AddXmpMetadataIssues(issues, documentInfo, options);
            AddTaggedContentIssues(issues, documentInfo, options);
        }

        if (!string.IsNullOrWhiteSpace(options.RequiredArtifactSha256) &&
            !string.Equals(artifactSha256, NormalizeSha256(options.RequiredArtifactSha256), StringComparison.Ordinal)) {
            issues.Add(new PdfConversionProofIssue("ArtifactSha256", NormalizeSha256(options.RequiredArtifactSha256), artifactSha256));
        }

        for (int i = 0; i < options.RequiredWarningCodes.Count; i++) {
            string code = options.RequiredWarningCodes[i];
            if (!Warnings.Any(warning => string.Equals(warning.Code, code, StringComparison.Ordinal))) {
                issues.Add(new PdfConversionProofIssue("WarningCode", code, "missing"));
            }
        }

        for (int i = 0; i < options.RequiredWarningSources.Count; i++) {
            string source = options.RequiredWarningSources[i];
            if (!Warnings.Any(warning => string.Equals(warning.Source, source, StringComparison.Ordinal))) {
                issues.Add(new PdfConversionProofIssue("WarningSource", source, "missing"));
            }
        }

        PdfConversionReportSummary warningSummary = Summary;
        if (options.RequireNoUnexpectedWarnings) {
            AddUnexpectedWarningIssues(issues, options);
        }

        if (options.RequireNoErrorWarnings && warningSummary.HasErrors) {
            issues.Add(new PdfConversionProofIssue(
                "WarningSeverity",
                "0 error warnings",
                warningSummary.ErrorCount.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }

        return new PdfConversionProofReport(documentInfo, logicalDocument, extractedText, logicalSignals, artifactByteCount, artifactSha256, warningSummary, issues);
    }

    /// <summary>
    /// Builds a conversion proof snapshot and throws when requested proof checks fail.
    /// </summary>
    public PdfConversionProofReport AssertProof(PdfConversionProofOptions? options = null) {
        PdfConversionProofReport proof = AssessProof(options);
        proof.ThrowIfFailed();
        return proof;
    }

    /// <summary>Returns the generated PDF bytes.</summary>
    public byte[] ToBytes() {
        try {
            return Value.ToBytes();
        } finally {
            RefreshConversionReport();
        }
    }

    /// <summary>
    /// Writes the generated PDF document to the supplied stream and returns conversion plus output evidence.
    /// </summary>
    public PdfSaveResult Save(Stream stream) {
        PdfSaveResult result;
        try {
            result = Value.Save(stream);
        } finally {
            RefreshConversionReport();
        }
        return result.WithReport(Report);
    }

    /// <summary>
    /// Writes the generated PDF document to the supplied file path and returns conversion plus output evidence.
    /// </summary>
    public PdfSaveResult Save(string path) {
        PdfSaveResult result;
        try {
            result = Value.Save(path);
        } finally {
            RefreshConversionReport();
        }
        return result.WithReport(Report);
    }

    /// <summary>
    /// Attempts to write the generated PDF document to the supplied stream and returns output diagnostics instead of throwing.
    /// </summary>
    public PdfSaveResult TrySave(Stream stream) {
        PdfSaveResult result;
        try {
            result = Value.TrySave(stream);
        } finally {
            RefreshConversionReport();
        }
        return result.WithReport(Report);
    }

    /// <summary>
    /// Attempts to write the generated PDF document to the supplied file path and returns output diagnostics instead of throwing.
    /// </summary>
    public PdfSaveResult TrySave(string path) {
        PdfSaveResult result;
        try {
            result = Value.TrySave(path);
        } finally {
            RefreshConversionReport();
        }
        return result.WithReport(Report);
    }

    /// <summary>
    /// Asynchronously writes the generated PDF document to the supplied stream and returns conversion plus output evidence.
    /// </summary>
    public async System.Threading.Tasks.Task<PdfSaveResult> SaveAsync(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        PdfSaveResult result;
        try {
            result = await Value.SaveAsync(stream, cancellationToken).ConfigureAwait(false);
        } finally {
            RefreshConversionReport();
        }
        return result.WithReport(Report);
    }

    /// <summary>
    /// Asynchronously writes the generated PDF document to the supplied file path and returns conversion plus output evidence.
    /// </summary>
    public async System.Threading.Tasks.Task<PdfSaveResult> SaveAsync(string path, System.Threading.CancellationToken cancellationToken = default) {
        PdfSaveResult result;
        try {
            result = await Value.SaveAsync(path, cancellationToken).ConfigureAwait(false);
        } finally {
            RefreshConversionReport();
        }
        return result.WithReport(Report);
    }

    /// <summary>
    /// Attempts to asynchronously write the generated PDF document to the supplied stream and returns output diagnostics instead of throwing.
    /// </summary>
    public System.Threading.Tasks.Task<PdfSaveResult> TrySaveAsync(Stream stream, System.Threading.CancellationToken cancellationToken = default) {
        return TrySaveAsyncCore(stream, cancellationToken);
    }

    /// <summary>
    /// Attempts to asynchronously write the generated PDF document to the supplied file path and returns output diagnostics instead of throwing.
    /// </summary>
    public System.Threading.Tasks.Task<PdfSaveResult> TrySaveAsync(string path, System.Threading.CancellationToken cancellationToken = default) {
        return TrySaveAsyncCore(path, cancellationToken);
    }

    private async System.Threading.Tasks.Task<PdfSaveResult> TrySaveAsyncCore(Stream stream, System.Threading.CancellationToken cancellationToken) {
        PdfSaveResult result;
        try {
            result = await Value.TrySaveAsync(stream, cancellationToken).ConfigureAwait(false);
        } finally {
            RefreshConversionReport();
        }
        return result.WithReport(Report);
    }

    private async System.Threading.Tasks.Task<PdfSaveResult> TrySaveAsyncCore(string path, System.Threading.CancellationToken cancellationToken) {
        PdfSaveResult result;
        try {
            result = await Value.TrySaveAsync(path, cancellationToken).ConfigureAwait(false);
        } finally {
            RefreshConversionReport();
        }
        return result.WithReport(Report);
    }

    private static PdfConversionReport SnapshotReport(PdfConversionReport conversionReport) {
        var snapshot = new PdfConversionReport();
        snapshot.AddRange(conversionReport.Warnings);
        return snapshot;
    }

    private void RefreshConversionReport() {
        IReadOnlyList<PdfConversionWarning> sourceWarnings = _sourceReport.Warnings;
        for (int i = 0; i < sourceWarnings.Count; i++) {
            PdfConversionWarning warning = sourceWarnings[i];
            if (!ContainsEquivalentWarning(Report.Warnings, warning)) {
                Report.Add(warning);
            }
        }
    }

    private static bool ContainsEquivalentWarning(IReadOnlyList<PdfConversionWarning> warnings, PdfConversionWarning candidate) {
        for (int i = 0; i < warnings.Count; i++) {
            PdfConversionWarning warning = warnings[i];
            if (string.Equals(warning.Converter, candidate.Converter, StringComparison.Ordinal) &&
                string.Equals(warning.Code, candidate.Code, StringComparison.Ordinal) &&
                string.Equals(warning.Source, candidate.Source, StringComparison.Ordinal) &&
                string.Equals(warning.Message, candidate.Message, StringComparison.Ordinal) &&
                warning.Severity == candidate.Severity) {
                return true;
            }
        }

        return false;
    }

    private static bool ShouldCaptureSaveDiagnostics(PdfConversionProofOptions options) =>
        options.RequiredWarningCodes.Count > 0 ||
        options.RequiredWarningSources.Count > 0 ||
        options.RequireNoUnexpectedWarnings ||
        options.RequireNoErrorWarnings;

    private void AddUnexpectedWarningIssues(List<PdfConversionProofIssue> issues, PdfConversionProofOptions options) {
        var unexpectedCodes = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < Warnings.Count; i++) {
            PdfConversionWarning warning = Warnings[i];
            if (!options.AcceptedWarningCodes.Contains(warning.Code)) {
                unexpectedCodes.Add(warning.Code);
            }
        }

        foreach (string code in unexpectedCodes.OrderBy(value => value, StringComparer.Ordinal)) {
            issues.Add(new PdfConversionProofIssue("UnexpectedWarningCode", "accepted warning code", code));
        }
    }

    private static void AddPageSizeIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        double expectedWidth = options.RequiredPageWidth!.Value;
        double expectedHeight = options.RequiredPageHeight!.Value;
        double tolerance = options.RequiredPageSizeTolerance;
        for (int i = 0; i < documentInfo.Pages.Count; i++) {
            PdfPageInfo page = documentInfo.Pages[i];
            if (Math.Abs(page.Width - expectedWidth) > tolerance || Math.Abs(page.Height - expectedHeight) > tolerance) {
                issues.Add(new PdfConversionProofIssue(
                    "PageSize",
                    FormatPageSize(expectedWidth, expectedHeight),
                    "page " + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + FormatPageSize(page.Width, page.Height)));
                return;
            }
        }
    }

    private static bool ShouldReadGeneratedPdf(PdfConversionProofOptions options) {
        return options.RequireReadablePdf ||
            options.RequiredTextMarkers.Count > 0 ||
            options.RequiredLogicalSignals.Count > 0 ||
            options.RequiredOutlineTitles.Count > 0 ||
            options.RequiredLinkUris.Count > 0 ||
            options.RequiredFormFieldNames.Count > 0 ||
            options.RequiredNamedDestinationNames.Count > 0 ||
            options.RequiredPageLabelRanges.Count > 0 ||
            options.RequiredAttachmentFileNames.Count > 0 ||
            options.RequiredOutputIntentSubtypes.Count > 0 ||
            options.RequiredOutputConditionIdentifiers.Count > 0 ||
            HasRequiredCatalogMetadata(options) ||
            options.RequiredOpenActionPageNumber.HasValue ||
            options.RequiredOpenActionDestinationMode.HasValue ||
            options.RequiredViewerPreferences.Count > 0 ||
            HasRequiredOptionalContent(options) ||
            HasRequiredXmpMetadata(options) ||
            HasRequiredTaggedContent(options) ||
            options.RequiredPageCount.HasValue ||
            (options.RequiredPageWidth.HasValue && options.RequiredPageHeight.HasValue) ||
            HasRequiredMetadata(options);
    }

    private static void AddMetadataIssues(List<PdfConversionProofIssue> issues, PdfMetadata metadata, PdfConversionProofOptions options) {
        AddMetadataIssue(issues, "Metadata.Title", options.RequiredMetadataTitle, metadata.Title);
        AddMetadataIssue(issues, "Metadata.Author", options.RequiredMetadataAuthor, metadata.Author);
        AddMetadataIssue(issues, "Metadata.Subject", options.RequiredMetadataSubject, metadata.Subject);
        AddMetadataIssue(issues, "Metadata.Keywords", options.RequiredMetadataKeywords, metadata.Keywords);
    }

    private static void AddMetadataIssue(List<PdfConversionProofIssue> issues, string feature, string? expected, string? actual) {
        if (expected is null) {
            return;
        }

        string actualValue = actual ?? string.Empty;
        if (!string.Equals(actualValue, expected, StringComparison.Ordinal)) {
            issues.Add(new PdfConversionProofIssue(feature, expected, actualValue));
        }
    }

    private static bool HasRequiredMetadata(PdfConversionProofOptions options) {
        return options.RequiredMetadataTitle is not null ||
            options.RequiredMetadataAuthor is not null ||
            options.RequiredMetadataSubject is not null ||
            options.RequiredMetadataKeywords is not null;
    }

    private static void AddOutlineTitleIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        for (int i = 0; i < options.RequiredOutlineTitles.Count; i++) {
            string title = options.RequiredOutlineTitles[i];
            if (!ContainsOutlineTitle(documentInfo.Outlines, title)) {
                issues.Add(new PdfConversionProofIssue("OutlineTitle", title, "missing"));
            }
        }
    }

    private static bool ContainsOutlineTitle(IReadOnlyList<PdfOutlineItem> outlines, string title) {
        for (int i = 0; i < outlines.Count; i++) {
            PdfOutlineItem outline = outlines[i];
            if (string.Equals(outline.Title, title, StringComparison.Ordinal)) {
                return true;
            }

            if (ContainsOutlineTitle(outline.Children, title)) {
                return true;
            }
        }

        return false;
    }

    private static void AddLinkUriIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        for (int i = 0; i < options.RequiredLinkUris.Count; i++) {
            string uri = options.RequiredLinkUris[i];
            if (!ContainsExact(documentInfo.LinkUris, uri)) {
                issues.Add(new PdfConversionProofIssue("LinkUri", uri, "missing"));
            }
        }
    }

    private static void AddFormFieldNameIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        for (int i = 0; i < options.RequiredFormFieldNames.Count; i++) {
            string name = options.RequiredFormFieldNames[i];
            if (!ContainsExact(documentInfo.FormFieldNames, name)) {
                issues.Add(new PdfConversionProofIssue("FormFieldName", name, "missing"));
            }
        }
    }

    private static void AddNamedDestinationNameIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        for (int i = 0; i < options.RequiredNamedDestinationNames.Count; i++) {
            string name = options.RequiredNamedDestinationNames[i];
            if (!ContainsExact(documentInfo.NamedDestinationNames, name)) {
                issues.Add(new PdfConversionProofIssue("NamedDestination", name, "missing"));
            }
        }
    }

    private static void AddPageLabelRangeIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        for (int i = 0; i < options.RequiredPageLabelRanges.Count; i++) {
            PdfPageLabelRange range = options.RequiredPageLabelRanges[i];
            if (!ContainsPageLabelRange(documentInfo.PageLabels, range)) {
                issues.Add(new PdfConversionProofIssue("PageLabel", FormatPageLabelRange(range), "missing"));
            }
        }
    }

    private static bool ContainsPageLabelRange(IReadOnlyList<PdfPageLabel> labels, PdfPageLabelRange expected) {
        string expectedStyle = PdfPageLabelDictionaryBuilder.GetStyleName(expected.Style);
        for (int i = 0; i < labels.Count; i++) {
            PdfPageLabel label = labels[i];
            int actualStartNumber = label.StartNumber ?? 1;
            if (label.StartPageNumber == expected.StartPageNumber &&
                actualStartNumber == expected.StartNumber &&
                string.Equals(label.Style, expectedStyle, StringComparison.Ordinal) &&
                string.Equals(label.Prefix ?? string.Empty, expected.Prefix ?? string.Empty, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static void AddAttachmentFileNameIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        for (int i = 0; i < options.RequiredAttachmentFileNames.Count; i++) {
            string fileName = options.RequiredAttachmentFileNames[i];
            if (!ContainsExact(documentInfo.AttachmentFileNames, fileName)) {
                issues.Add(new PdfConversionProofIssue("AttachmentFileName", fileName, "missing"));
            }
        }
    }

    private static void AddOutputIntentIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        for (int i = 0; i < options.RequiredOutputIntentSubtypes.Count; i++) {
            string subtype = options.RequiredOutputIntentSubtypes[i];
            if (!ContainsExact(documentInfo.OutputIntentSubtypes, subtype)) {
                issues.Add(new PdfConversionProofIssue("OutputIntentSubtype", subtype, "missing"));
            }
        }

        for (int i = 0; i < options.RequiredOutputConditionIdentifiers.Count; i++) {
            string identifier = options.RequiredOutputConditionIdentifiers[i];
            if (!ContainsExact(documentInfo.OutputConditionIdentifiers, identifier)) {
                issues.Add(new PdfConversionProofIssue("OutputConditionIdentifier", identifier, "missing"));
            }
        }
    }

    private static void AddCatalogMetadataIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        AddStringIssue(issues, "CatalogLanguage", options.RequiredCatalogLanguage, documentInfo.CatalogLanguage);
        AddStringIssue(issues, "CatalogPageMode", options.RequiredCatalogPageMode, documentInfo.CatalogPageMode);
        AddStringIssue(issues, "CatalogPageLayout", options.RequiredCatalogPageLayout, documentInfo.CatalogPageLayout);
    }

    private static void AddOpenActionIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        if (!options.RequiredOpenActionPageNumber.HasValue && !options.RequiredOpenActionDestinationMode.HasValue) {
            return;
        }

        PdfDocumentOpenAction? openAction = documentInfo.OpenAction;
        if (openAction is null) {
            if (options.RequiredOpenActionPageNumber.HasValue) {
                issues.Add(new PdfConversionProofIssue(
                    "OpenAction.PageNumber",
                    options.RequiredOpenActionPageNumber.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    "missing"));
            }

            if (options.RequiredOpenActionDestinationMode.HasValue) {
                issues.Add(new PdfConversionProofIssue(
                    "OpenAction.DestinationMode",
                    options.RequiredOpenActionDestinationMode.Value.ToString(),
                    "missing"));
            }

            return;
        }

        if (options.RequiredOpenActionPageNumber.HasValue && openAction.PageNumber != options.RequiredOpenActionPageNumber.Value) {
            issues.Add(new PdfConversionProofIssue(
                "OpenAction.PageNumber",
                options.RequiredOpenActionPageNumber.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                openAction.PageNumber.HasValue ? openAction.PageNumber.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : string.Empty));
        }

        if (options.RequiredOpenActionDestinationMode.HasValue && openAction.DestinationMode != options.RequiredOpenActionDestinationMode.Value) {
            issues.Add(new PdfConversionProofIssue(
                "OpenAction.DestinationMode",
                options.RequiredOpenActionDestinationMode.Value.ToString(),
                openAction.DestinationMode.HasValue ? openAction.DestinationMode.Value.ToString() : string.Empty));
        }
    }

    private static void AddViewerPreferenceIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        if (options.RequiredViewerPreferences.Count == 0) {
            return;
        }

        PdfViewerPreferences? viewerPreferences = documentInfo.ViewerPreferences;
        foreach (KeyValuePair<string, string> requirement in options.RequiredViewerPreferences) {
            string actual = viewerPreferences?.GetValue(requirement.Key) ?? string.Empty;
            if (!string.Equals(actual, requirement.Value, StringComparison.Ordinal)) {
                issues.Add(new PdfConversionProofIssue("ViewerPreference." + requirement.Key, requirement.Value, string.IsNullOrEmpty(actual) ? "missing" : actual));
            }
        }
    }

    private static void AddStringIssue(List<PdfConversionProofIssue> issues, string feature, string? expected, string? actual) {
        if (expected is null) {
            return;
        }

        string actualValue = actual ?? string.Empty;
        if (!string.Equals(actualValue, expected, StringComparison.Ordinal)) {
            issues.Add(new PdfConversionProofIssue(feature, expected, string.IsNullOrEmpty(actualValue) ? "missing" : actualValue));
        }
    }

    private static bool HasRequiredCatalogMetadata(PdfConversionProofOptions options) {
        return options.RequiredCatalogLanguage is not null ||
            options.RequiredCatalogPageMode is not null ||
            options.RequiredCatalogPageLayout is not null;
    }

    private static bool ContainsExact(IReadOnlyList<string> values, string expected) {
        for (int i = 0; i < values.Count; i++) {
            if (string.Equals(values[i], expected, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static string ComputeSha256Hex(byte[] bytes) {
#if NET8_0_OR_GREATER
        return ToLowerHex(System.Security.Cryptography.SHA256.HashData(bytes));
#else
        using (System.Security.Cryptography.SHA256 sha256 = System.Security.Cryptography.SHA256.Create()) {
            return ToLowerHex(sha256.ComputeHash(bytes));
        }
#endif
    }

    private static string ToLowerHex(byte[] bytes) {
        const string hex = "0123456789abcdef";
        char[] chars = new char[bytes.Length * 2];
        for (int i = 0; i < bytes.Length; i++) {
            chars[i * 2] = hex[bytes[i] >> 4];
            chars[(i * 2) + 1] = hex[bytes[i] & 0x0F];
        }

        return new string(chars);
    }

    private static string NormalizeSha256(string? sha256) {
        if (string.IsNullOrWhiteSpace(sha256)) {
            return string.Empty;
        }

        return sha256!.Trim().ToLowerInvariant();
    }

    private static string FormatPageSize(double width, double height) {
        return width.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) +
            "x" +
            height.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string FormatPageLabelRange(PdfPageLabelRange range) {
        string value = "page " +
            range.StartPageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " " +
            PdfPageLabelDictionaryBuilder.GetStyleName(range.Style) +
            " start " +
            range.StartNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (range.Prefix is not null) {
            value += " prefix " + range.Prefix;
        }

        return value;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> BuildLogicalSignals(PdfDocumentInfo documentInfo, PdfLogicalDocument logicalDocument, string extractedText) {
        var signals = new List<string>();

        AddSignal(signals, documentInfo.PageCount > 0, "page-count");
        AddSignal(signals, HasReadablePageGeometry(documentInfo), "page-geometry");
        AddSignal(signals, HasDocumentMetadata(documentInfo.Metadata), "metadata", "document-metadata");
        AddSignal(signals, !string.IsNullOrWhiteSpace(extractedText), "text", "text-blocks", "worksheet-text", "positioned-text", "logical-readback");
        AddSignal(signals, logicalDocument.Headings.Count > 0 || logicalDocument.Outlines.Count > 0 || documentInfo.Outlines.Count > 0, "headings");
        AddSignal(signals, logicalDocument.Paragraphs.Count > 0, "paragraphs");
        AddSignal(signals, logicalDocument.ListItems.Count > 0, "lists", "list-items");
        AddSignal(signals, logicalDocument.Tables.Count > 0, "tables", "logical-table-reconstruction", "table-confidence", "numeric-column-profiles");
        AddSignal(signals, logicalDocument.Images.Count > 0, "images", "image-extraction", "image-geometry", "image-placeholders");
        AddSignal(signals, logicalDocument.Links.Count > 0 || documentInfo.LinkAnnotationCount > 0, "links", "safe-links");
        AddSignal(signals, logicalDocument.Outlines.Count > 0 || documentInfo.Outlines.Count > 0, "outlines", "bookmarks");
        AddSignal(signals, logicalDocument.FormFields.Count > 0 || documentInfo.FormFieldCount > 0, "form-fields");
        AddSignal(signals, logicalDocument.FormWidgets.Count > 0 || documentInfo.FormWidgetCount > 0, "form-widgets");
        AddSignal(signals, logicalDocument.Attachments.Count > 0 || documentInfo.Attachments.Count > 0, "attachments", "embedded-files");
        AddSignal(signals, logicalDocument.NamedDestinations.Count > 0 || documentInfo.NamedDestinationCount > 0, "named-destinations");
        AddSignal(signals, logicalDocument.PageLabels.Count > 0 || documentInfo.PageLabelCount > 0, "page-labels");
        AddSignal(signals, logicalDocument.HasReadableOutputIntents || documentInfo.OutputIntents.Count > 0, "output-intents");
        AddSignal(signals, logicalDocument.HasReadableTaggedContent || documentInfo.TaggedContent is not null, "tagged-content");
        AddSignal(signals, logicalDocument.HasReadableOptionalContent || documentInfo.OptionalContent is not null, "optional-content", "layers");
        AddSignal(signals, logicalDocument.XmpMetadata is not null || documentInfo.XmpMetadata is not null, "xmp", "xmp-metadata");
        AddSignal(signals, logicalDocument.HasSecurityState || documentInfo.Security.HasEncryption || documentInfo.Security.HasSignatures, "security-metadata");
        AddSignal(signals, HasCatalogViewMetadata(documentInfo), "catalog-view", "viewer-initial-view");
        AddSignal(signals, logicalDocument.HasReadableOpenAction || documentInfo.OpenAction is not null, "open-action", "document-open-action");
        AddSignal(signals, logicalDocument.HasCatalogActions || documentInfo.CatalogActions.Count > 0, "catalog-actions");
        AddSignal(signals, logicalDocument.HasPageActions || documentInfo.Pages.Any(page => page.HasPageActions), "page-actions");
        AddSignal(signals, logicalDocument.ViewerPreferences is not null || documentInfo.ViewerPreferences is not null, "viewer-preferences");

        return signals.AsReadOnly();
    }

    private static bool HasReadablePageGeometry(PdfDocumentInfo documentInfo) {
        for (int i = 0; i < documentInfo.Pages.Count; i++) {
            PdfPageInfo page = documentInfo.Pages[i];
            if (page.Width <= 0D || page.Height <= 0D) {
                return false;
            }
        }

        return documentInfo.Pages.Count > 0;
    }

    private static bool HasDocumentMetadata(PdfMetadata metadata) {
        return !string.IsNullOrEmpty(metadata.Title) ||
            !string.IsNullOrEmpty(metadata.Author) ||
            !string.IsNullOrEmpty(metadata.Subject) ||
            !string.IsNullOrEmpty(metadata.Keywords);
    }

    private static bool HasCatalogViewMetadata(PdfDocumentInfo documentInfo) {
        return !string.IsNullOrEmpty(documentInfo.CatalogPageMode) ||
            !string.IsNullOrEmpty(documentInfo.CatalogPageLayout) ||
            !string.IsNullOrEmpty(documentInfo.CatalogLanguage);
    }

    private static bool ContainsLogicalSignal(IReadOnlyList<string> signals, string requiredSignal) {
        for (int i = 0; i < signals.Count; i++) {
            if (string.Equals(signals[i], requiredSignal, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static void AddSignal(List<string> signals, bool condition, params string[] names) {
        if (!condition) {
            return;
        }

        for (int i = 0; i < names.Length; i++) {
            if (!ContainsLogicalSignal(signals, names[i])) {
                signals.Add(names[i]);
            }
        }
    }

}
