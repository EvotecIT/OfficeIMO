using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

/// <summary>
/// PDF ingestion adapter for <see cref="OfficeDocumentReader"/>.
/// </summary>
internal static partial class PdfReaderAdapter {
    private const double HighConfidenceTableThreshold = 0.95D;

    /// <summary>
    /// Reads a PDF file and emits normalized page-aware chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(string pdfPath, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (pdfPath == null) throw new ArgumentNullException(nameof(pdfPath));
        if (pdfPath.Length == 0) throw new ArgumentException("PDF path cannot be empty.", nameof(pdfPath));
        if (!File.Exists(pdfPath)) throw new FileNotFoundException($"PDF file '{pdfPath}' doesn't exist.", pdfPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectivePdfOptions = ReaderPdfOptionsCloner.CloneOrDefault(pdfOptions);
        ReaderInputLimits.EnforceFileSize(pdfPath, effectiveReaderOptions.MaxInputBytes);
        var source = BuildSourceMetadataFromPath(pdfPath, effectiveReaderOptions.ComputeHashes);

        PdfDocument pdf = PdfDocument.Open(pdfPath, CreatePdfReadOptions(effectiveReaderOptions));
        PdfLogicalDocument document = LoadDocument(pdf, effectivePdfOptions);
        foreach (var chunk in Read(document, source, effectiveReaderOptions, effectivePdfOptions, applyPageRanges: false, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads a PDF stream and emits normalized page-aware chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(Stream pdfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (pdfStream == null) throw new ArgumentNullException(nameof(pdfStream));
        if (!pdfStream.CanRead) throw new ArgumentException("PDF stream must be readable.", nameof(pdfStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectivePdfOptions = ReaderPdfOptionsCloner.CloneOrDefault(pdfOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.pdf");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        cancellationToken.ThrowIfCancellationRequested();
        PdfDocument pdf = OpenReaderPdf(pdfStream, effectiveReaderOptions);
        UpdateSourceMetadataFromPdfDocument(source, pdf, effectiveReaderOptions.ComputeHashes);
        PdfLogicalDocument document = LoadDocument(pdf, effectivePdfOptions);
        foreach (var chunk in Read(document, source, effectiveReaderOptions, effectivePdfOptions, applyPageRanges: false, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads PDF bytes and emits normalized page-aware chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(byte[] pdfBytes, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (pdfBytes == null) throw new ArgumentNullException(nameof(pdfBytes));

        using var stream = new MemoryStream(pdfBytes, writable: false);
        foreach (var chunk in Read(stream, sourceName, readerOptions, pdfOptions, cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Converts an already loaded logical PDF model into normalized Reader chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(PdfLogicalDocument document, string sourceName = "document.pdf", ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectivePdfOptions = ReaderPdfOptionsCloner.CloneOrDefault(pdfOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.pdf");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        foreach (var chunk in Read(document, source, effectiveReaderOptions, effectivePdfOptions, applyPageRanges: true, cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> Read(PdfLogicalDocument document, SourceMetadata source, ReaderOptions readerOptions, ReaderPdfOptions pdfOptions, bool applyPageRanges, CancellationToken cancellationToken) {
        int maxChars = readerOptions.MaxChars > 0 ? readerOptions.MaxChars : 8_000;
        var markdownOptions = ReaderPdfOptions.CloneMarkdownOptions(pdfOptions.MarkdownOptions) ?? ReaderPdfOptions.CreateOfficeIMOProfile().MarkdownOptions!;
        markdownOptions.IncludePageSeparators = false;
        IReadOnlyList<PdfLogicalPage> pages = applyPageRanges ? GetReaderPages(document, pdfOptions) : document.Pages;

        if (!pdfOptions.ChunkByPage) {
            string markdown = BuildMarkdown(pages, markdownOptions);
            IReadOnlyList<PdfLogicalTableExtraction> documentTableExtractions = PdfLogicalTableAnalysis.ExtractTables(pages, GetMaxTableRows(readerOptions));
            var documentTables = BuildTables(documentTableExtractions);
            var documentVisuals = BuildVisuals(pages);
            var documentFormFields = BuildFormFields(document, pages, page: null);
            var documentActions = BuildActions(document, pages, page: null);
            ReaderChunkDiagnostics documentDiagnostics = BuildChunkDiagnostics(document, pages, page: null, documentTableExtractions, documentActions);
            foreach (var chunk in BuildChunksFromText(
                markdown,
                source,
                readerOptions,
                page: null,
                sourceBlockIndex: 0,
                blockKind: "document",
                blockAnchor: "document",
                tables: documentTables,
                visuals: documentVisuals,
                formFields: documentFormFields,
                actions: documentActions,
                diagnostics: documentDiagnostics,
                idPrefix: "pdf-document",
                maxChars: maxChars,
                cancellationToken: cancellationToken)) {
                yield return chunk;
            }

            yield break;
        }

        int emittedIndex = 0;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();

            PdfLogicalPage page = pages[pageIndex];
            string pageOccurrence = pageIndex.ToString("D4", CultureInfo.InvariantCulture);
            string pageAnchor = "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + pageOccurrence;
            string idPrefix = "pdf-page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture) + "-selection-" + pageOccurrence;
            string markdown = page.ToMarkdown(markdownOptions);
            IReadOnlyList<PdfLogicalTableExtraction> pageTableExtractions = PdfLogicalTableAnalysis.ExtractTables(page, GetMaxTableRows(readerOptions));
            var pageTables = BuildTables(pageTableExtractions, pageIndex);
            var pageVisuals = BuildVisuals(new[] { page }, pageIndex);
            var pageFormFields = BuildFormFields(document, pages, page);
            var pageActions = BuildActions(document, pages, page);
            ReaderChunkDiagnostics pageDiagnostics = BuildChunkDiagnostics(document, pages, page, pageTableExtractions, pageActions);
            foreach (var chunk in BuildChunksFromText(
                markdown,
                source,
                readerOptions,
                page.PageNumber,
                pageIndex,
                "page",
                pageAnchor,
                pageTables,
                pageVisuals,
                pageFormFields,
                pageActions,
                pageDiagnostics,
                idPrefix,
                maxChars,
                cancellationToken)) {
                chunk.Location.BlockIndex = emittedIndex++;
                yield return chunk;
            }
        }

        if (emittedIndex == 0) {
            const string warning = "PDF content produced no readable logical pages.";
            yield return EnrichChunk(new ReaderChunk {
                Id = "pdf-warning-0000",
                Kind = ReaderInputKind.Pdf,
                Location = new ReaderLocation {
                    Path = source.Path,
                    BlockIndex = 0,
                    SourceBlockKind = "warning"
                },
                Text = warning,
                Diagnostics = BuildChunkDiagnostics(document, pages, page: null, Array.Empty<PdfLogicalTableExtraction>(), actions: null),
                Warnings = new[] { warning }
            }, source, readerOptions.ComputeHashes);
        }
    }

    private static IReadOnlyList<PdfLogicalPage> GetReaderPages(PdfLogicalDocument document, ReaderPdfOptions options) {
        IReadOnlyList<PdfPageRange>? ranges = options.PageRanges;
        if (ranges == null || ranges.Count == 0) {
            return document.Pages;
        }

        int maxSourcePageNumber = 0;
        for (int i = 0; i < document.Pages.Count; i++) {
            maxSourcePageNumber = Math.Max(maxSourcePageNumber, document.Pages[i].PageNumber);
        }

        if (maxSourcePageNumber == 0) {
            return Array.Empty<PdfLogicalPage>();
        }

        var pages = new List<PdfLogicalPage>();
        for (int rangeIndex = 0; rangeIndex < ranges.Count; rangeIndex++) {
            PdfPageRange range = ranges[rangeIndex];
            if (range.FirstPage < 1 || range.LastPage < range.FirstPage) {
                throw new ArgumentOutOfRangeException(nameof(ReaderPdfOptions.PageRanges), "Page ranges must be inclusive one-based ranges.");
            }

            if (range.LastPage > maxSourcePageNumber) {
                throw new ArgumentOutOfRangeException(nameof(ReaderPdfOptions.PageRanges), "Page range cannot exceed the document page count.");
            }

            for (int pageNumber = range.FirstPage; pageNumber <= range.LastPage; pageNumber++) {
                IReadOnlyList<PdfLogicalPage> sourcePages = document.GetPages(pageNumber);
                for (int sourceIndex = 0; sourceIndex < sourcePages.Count; sourceIndex++) {
                    pages.Add(sourcePages[sourceIndex]);
                }
            }
        }

        return pages.AsReadOnly();
    }

    private static string BuildMarkdown(IReadOnlyList<PdfLogicalPage> pages, PdfLogicalMarkdownOptions markdownOptions) {
        if (pages.Count == 0) {
            return string.Empty;
        }

        return string.Join(Environment.NewLine + Environment.NewLine, pages.Select(page => page.ToMarkdown(markdownOptions)).Where(text => !string.IsNullOrWhiteSpace(text)));
    }

    private static IEnumerable<ReaderChunk> BuildChunksFromText(string markdown, SourceMetadata source, ReaderOptions readerOptions, int? page, int sourceBlockIndex, string blockKind, string blockAnchor, IReadOnlyList<ReaderTable>? tables, IReadOnlyList<ReaderVisual>? visuals, IReadOnlyList<ReaderFormField>? formFields, IReadOnlyList<ReaderActionSummary>? actions, ReaderChunkDiagnostics diagnostics, string idPrefix, int maxChars, CancellationToken cancellationToken) {
        var parts = SplitText(markdown, maxChars);
        if (parts.Count == 0) {
            string warning = page.HasValue
                ? "PDF page " + page.Value.ToString(CultureInfo.InvariantCulture) + " produced no readable text."
                : "PDF content produced no readable text.";
            parts = new[] { new TextPart(warning, new[] { warning }) };
        }

        for (int i = 0; i < parts.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            TextPart part = parts[i];
            yield return EnrichChunk(new ReaderChunk {
                Id = parts.Count == 1
                    ? idPrefix
                    : idPrefix + "-part-" + i.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Pdf,
                Location = new ReaderLocation {
                    Path = source.Path,
                    SourceBlockIndex = sourceBlockIndex,
                    SourceBlockKind = blockKind,
                    BlockAnchor = parts.Count == 1 ? blockAnchor : blockAnchor + "-part-" + i.ToString("D4", CultureInfo.InvariantCulture),
                    Page = page
                },
                Text = part.Text,
                Markdown = part.Text,
                Tables = i == 0 ? tables : null,
                Visuals = i == 0 ? visuals : null,
                FormFields = i == 0 ? formFields : null,
                Actions = i == 0 ? actions : null,
                Diagnostics = diagnostics,
                Warnings = part.Warnings
            }, source, readerOptions.ComputeHashes);
        }
    }

    private static ReaderChunkDiagnostics BuildChunkDiagnostics(PdfLogicalDocument document, IReadOnlyList<PdfLogicalPage> selectedPages, PdfLogicalPage? page, IReadOnlyList<PdfLogicalTableExtraction> tableExtractions, IReadOnlyList<ReaderActionSummary>? actions) {
        IReadOnlyList<PdfLogicalPage> scope = page is null ? selectedPages : new[] { page };
        PdfDocumentSecurityInfo security = document.Security;
        bool hasScopedOpenAction = GetScopedOpenAction(document.OpenAction, scope) is not null;
        int selectedCatalogActionCount = GetScopedCatalogActions(document, selectedPages, page: null).Count;
        int selectedPageActionCount = CountPageActions(scope);
        int selectedAnnotationActionCount = CountAnnotationActions(scope);
        int imageCount = CountImages(scope);
        int imageGeometryCount = CountImageGeometry(scope);
        int imageNonAxisAlignedCount = CountNonAxisAlignedImages(scope);
        int selectedFormWidgetCount = CountFormWidgets(scope);
        int selectedFormWidgetAppearanceStateCount = CountFormWidgetAppearanceStates(scope);
        TableDiagnosticSummary tableSummary = SummarizeTables(tableExtractions);
        ActionDiagnosticSummary actionSummary = SummarizeActions(actions);
        PdfTaggedContentInfo? taggedContent = document.TaggedContent;
        PdfOptionalContentProperties? optionalContent = document.OptionalContent;
        return new ReaderChunkDiagnostics {
            SourceKind = "pdf",
            PageCount = document.PageCount,
            SelectedPageCount = selectedPages.Count,
            PageNumber = page?.PageNumber,
            TableCount = tableExtractions.Count,
            TableGeometryCount = tableSummary.GeometryCount,
            TableGeometryCoverage = GetCoverage(tableSummary.GeometryCount, tableExtractions.Count),
            MinTableConfidence = tableSummary.MinConfidence,
            AverageTableConfidence = tableSummary.AverageConfidence,
            LowConfidenceTableCount = tableSummary.LowConfidenceCount,
            NumericTableColumnCount = tableSummary.NumericColumnCount,
            FallbackTableColumnNameCount = tableSummary.FallbackColumnNameCount,
            MissingTableCellCount = tableSummary.MissingCellCount,
            ImageCount = imageCount,
            ImageGeometryCount = imageGeometryCount,
            ImageGeometryCoverage = GetCoverage(imageGeometryCount, imageCount),
            ImageNonAxisAlignedCount = imageNonAxisAlignedCount,
            ImageNonAxisAlignedCoverage = GetCoverage(imageNonAxisAlignedCount, imageGeometryCount),
            LinkCount = CountLinks(scope),
            HasXmpMetadata = document.XmpMetadata != null,
            OutputIntentCount = document.OutputIntents.Count,
            AttachmentCount = document.Attachments.Count,
            HasTaggedContent = taggedContent != null,
            TaggedStructureElementCount = taggedContent?.StructureElementCount ?? 0,
            TaggedMarkedContentReferenceCount = taggedContent?.MarkedContentReferenceCount ?? 0,
            OptionalContentGroupCount = optionalContent?.GroupCount ?? 0,
            OptionalContentInitiallyHiddenCount = CountInitiallyHiddenOptionalContentGroups(optionalContent),
            OptionalContentLockedCount = CountLockedOptionalContentGroups(optionalContent),
            HasOpenAction = hasScopedOpenAction,
            HasCatalogActions = selectedCatalogActionCount > 0,
            HasPageActions = selectedPageActionCount > 0,
            HasAnnotationActions = selectedAnnotationActionCount > 0,
            HasActiveContent = selectedCatalogActionCount > 0 || selectedPageActionCount > 0 || selectedAnnotationActionCount > 0,
            PotentiallyUnsafeActionCount = actionSummary.PotentiallyUnsafeCount,
            JavaScriptActionCount = actionSummary.JavaScriptCount,
            LaunchActionCount = actionSummary.LaunchCount,
            SubmitFormActionCount = actionSummary.SubmitFormCount,
            ImportDataActionCount = actionSummary.ImportDataCount,
            CatalogActionCount = selectedCatalogActionCount,
            PageActionCount = document.PageActionCount,
            SelectedPageActionCount = selectedPageActionCount,
            AnnotationActionCount = CountAnnotationActions(document.Pages),
            SelectedAnnotationActionCount = selectedAnnotationActionCount,
            FormFieldCount = document.FormFields.Count,
            FormWidgetCount = document.FormWidgets.Count,
            SelectedFormWidgetCount = selectedFormWidgetCount,
            SelectedFormWidgetAppearanceStateCount = selectedFormWidgetAppearanceStateCount,
            SelectedFormWidgetAppearanceStateCoverage = GetCoverage(selectedFormWidgetAppearanceStateCount, selectedFormWidgetCount),
            SelectedFormWidgetNormalAppearanceStateCount = CountFormWidgetNormalAppearanceStates(scope),
            HasSecurityState = document.HasSecurityState,
            HasEncryption = security.HasEncryption,
            HasSignatures = security.HasSignatures,
            HasIncrementalUpdates = security.HasIncrementalUpdates,
            RevisionCount = security.RevisionCount,
            RequiresAppendOnlyMutation = security.RequiresAppendOnlyMutation
        };
    }

    private static int CountInitiallyHiddenOptionalContentGroups(PdfOptionalContentProperties? optionalContent) {
        if (optionalContent == null) {
            return 0;
        }

        int count = 0;
        for (int i = 0; i < optionalContent.Groups.Count; i++) {
            if (optionalContent.Groups[i].IsInitiallyVisible == false) {
                count++;
            }
        }

        return count;
    }

    private static int CountLockedOptionalContentGroups(PdfOptionalContentProperties? optionalContent) {
        if (optionalContent == null) {
            return 0;
        }

        int count = 0;
        for (int i = 0; i < optionalContent.Groups.Count; i++) {
            if (optionalContent.Groups[i].IsLocked == true) {
                count++;
            }
        }

        return count;
    }

    private readonly struct ActionDiagnosticSummary {
        public ActionDiagnosticSummary(int potentiallyUnsafeCount, int javaScriptCount, int launchCount, int submitFormCount, int importDataCount) {
            PotentiallyUnsafeCount = potentiallyUnsafeCount;
            JavaScriptCount = javaScriptCount;
            LaunchCount = launchCount;
            SubmitFormCount = submitFormCount;
            ImportDataCount = importDataCount;
        }

        public int PotentiallyUnsafeCount { get; }

        public int JavaScriptCount { get; }

        public int LaunchCount { get; }

        public int SubmitFormCount { get; }

        public int ImportDataCount { get; }
    }

    private static ActionDiagnosticSummary SummarizeActions(IReadOnlyList<ReaderActionSummary>? actions) {
        if (actions == null || actions.Count == 0) {
            return new ActionDiagnosticSummary(0, 0, 0, 0, 0);
        }

        int potentiallyUnsafeCount = 0;
        int javaScriptCount = 0;
        int launchCount = 0;
        int submitFormCount = 0;
        int importDataCount = 0;
        for (int i = 0; i < actions.Count; i++) {
            ReaderActionSummary action = actions[i];
            if (action.IsPotentiallyUnsafe) {
                potentiallyUnsafeCount++;
            }

            if (string.Equals(action.ActionType, "JavaScript", StringComparison.Ordinal)) {
                javaScriptCount++;
            } else if (string.Equals(action.ActionType, "Launch", StringComparison.Ordinal)) {
                launchCount++;
            } else if (string.Equals(action.ActionType, "SubmitForm", StringComparison.Ordinal)) {
                submitFormCount++;
            } else if (string.Equals(action.ActionType, "ImportData", StringComparison.Ordinal)) {
                importDataCount++;
            }
        }

        return new ActionDiagnosticSummary(potentiallyUnsafeCount, javaScriptCount, launchCount, submitFormCount, importDataCount);
    }

    private readonly struct TableDiagnosticSummary {
        public TableDiagnosticSummary(int geometryCount, double? minConfidence, double? averageConfidence, int lowConfidenceCount, int numericColumnCount, int fallbackColumnNameCount, int missingCellCount) {
            GeometryCount = geometryCount;
            MinConfidence = minConfidence;
            AverageConfidence = averageConfidence;
            LowConfidenceCount = lowConfidenceCount;
            NumericColumnCount = numericColumnCount;
            FallbackColumnNameCount = fallbackColumnNameCount;
            MissingCellCount = missingCellCount;
        }

        public int GeometryCount { get; }

        public double? MinConfidence { get; }

        public double? AverageConfidence { get; }

        public int LowConfidenceCount { get; }

        public int NumericColumnCount { get; }

        public int FallbackColumnNameCount { get; }

        public int MissingCellCount { get; }
    }

    private static TableDiagnosticSummary SummarizeTables(IReadOnlyList<PdfLogicalTableExtraction> tables) {
        if (tables.Count == 0) {
            return new TableDiagnosticSummary(0, null, null, 0, 0, 0, 0);
        }

        int geometryCount = 0;
        int lowConfidenceCount = 0;
        int numericColumnCount = 0;
        int fallbackColumnNameCount = 0;
        int missingCellCount = 0;
        double minConfidence = double.MaxValue;
        double totalConfidence = 0D;
        for (int i = 0; i < tables.Count; i++) {
            PdfLogicalTableData data = tables[i].Data;
            PdfLogicalTableDiagnostics diagnostics = data.Diagnostics;
            if (diagnostics.HasGeometry) {
                geometryCount++;
            }

            double confidence = diagnostics.Confidence;
            if (confidence < HighConfidenceTableThreshold) {
                lowConfidenceCount++;
            }

            if (confidence < minConfidence) {
                minConfidence = confidence;
            }

            totalConfidence += confidence;

            missingCellCount += diagnostics.MissingCellCount;
            for (int profileIndex = 0; profileIndex < data.ColumnProfiles.Count; profileIndex++) {
                PdfLogicalTableColumnProfile profile = data.ColumnProfiles[profileIndex];
                if (profile.IsNumeric) {
                    numericColumnCount++;
                }

                if (IsFallbackColumnName(profile.Name, profile.Index)) {
                    fallbackColumnNameCount++;
                }
            }
        }

        return new TableDiagnosticSummary(
            geometryCount,
            minConfidence,
            totalConfidence / tables.Count,
            lowConfidenceCount,
            numericColumnCount,
            fallbackColumnNameCount,
            missingCellCount);
    }

    private static bool IsFallbackColumnName(string? name, int columnIndex) {
        if (string.IsNullOrWhiteSpace(name)) {
            return false;
        }

        string trimmed = name!.Trim();
        return string.Equals(
            trimmed,
            "Column " + (columnIndex + 1).ToString(CultureInfo.InvariantCulture),
            StringComparison.Ordinal);
    }

    private static double GetCoverage(int countWithSignal, int totalCount) {
        return totalCount == 0 ? 0D : (double)countWithSignal / totalCount;
    }

    private static int CountImageGeometry(IReadOnlyList<PdfLogicalPage> pages) {
        int count = 0;
        for (int i = 0; i < pages.Count; i++) {
            IReadOnlyList<PdfLogicalImage> images = pages[i].Images;
            for (int j = 0; j < images.Count; j++) {
                if (images[j].PrimaryPlacement is not null) {
                    count++;
                }
            }
        }

        return count;
    }

    private static int CountNonAxisAlignedImages(IReadOnlyList<PdfLogicalPage> pages) {
        int count = 0;
        for (int i = 0; i < pages.Count; i++) {
            IReadOnlyList<PdfLogicalImage> images = pages[i].Images;
            for (int j = 0; j < images.Count; j++) {
                PdfImagePlacement? placement = images[j].PrimaryPlacement;
                if (placement is not null && !placement.IsAxisAligned) {
                    count++;
                }
            }
        }

        return count;
    }

    private static int CountImages(IReadOnlyList<PdfLogicalPage> pages) {
        int count = 0;
        for (int i = 0; i < pages.Count; i++) {
            count += pages[i].Images.Count;
        }

        return count;
    }

    private static int CountLinks(IReadOnlyList<PdfLogicalPage> pages) {
        int count = 0;
        for (int i = 0; i < pages.Count; i++) {
            count += pages[i].Links.Count;
        }

        return count;
    }

    private static int CountPageActions(IReadOnlyList<PdfLogicalPage> pages) {
        int count = 0;
        for (int i = 0; i < pages.Count; i++) {
            count += pages[i].PageActionCount;
        }

        return count;
    }

    private static int CountAnnotationActions(IReadOnlyList<PdfLogicalPage> pages) {
        int count = 0;
        for (int i = 0; i < pages.Count; i++) {
            IReadOnlyList<PdfAnnotation> annotations = pages[i].Annotations;
            for (int j = 0; j < annotations.Count; j++) {
                PdfAnnotation annotation = annotations[j];
                if (annotation.HasAction) {
                    count++;
                }

                count += annotation.AdditionalActions.Count;
                count += annotation.ChainedActions.Count;
            }
        }

        return count;
    }

    private static int CountFormWidgets(IReadOnlyList<PdfLogicalPage> pages) {
        int count = 0;
        for (int i = 0; i < pages.Count; i++) {
            count += pages[i].FormWidgets.Count;
        }

        return count;
    }

    private static int CountFormWidgetAppearanceStates(IReadOnlyList<PdfLogicalPage> pages) {
        int count = 0;
        for (int i = 0; i < pages.Count; i++) {
            IReadOnlyList<PdfLogicalFormWidget> widgets = pages[i].FormWidgets;
            for (int j = 0; j < widgets.Count; j++) {
                if (!string.IsNullOrEmpty(widgets[j].AppearanceState)) {
                    count++;
                }
            }
        }

        return count;
    }

    private static int CountFormWidgetNormalAppearanceStates(IReadOnlyList<PdfLogicalPage> pages) {
        int count = 0;
        for (int i = 0; i < pages.Count; i++) {
            IReadOnlyList<PdfLogicalFormWidget> widgets = pages[i].FormWidgets;
            for (int j = 0; j < widgets.Count; j++) {
                count += widgets[j].NormalAppearanceStateCount;
            }
        }

        return count;
    }

    private static IReadOnlyList<ReaderFormField>? BuildFormFields(PdfLogicalDocument document, IReadOnlyList<PdfLogicalPage> selectedPages, PdfLogicalPage? page) {
        if (document.FormFields.Count == 0) return null;

        int[] pageNumbers = GetSelectedPageNumbers(selectedPages, page);
        if (pageNumbers.Length == 0) return null;

        var result = new List<ReaderFormField>();
        for (int i = 0; i < document.FormFields.Count; i++) {
            PdfFormField field = document.FormFields[i];
            IReadOnlyList<PdfFormWidget> widgets = GetWidgetsForPages(field, pageNumbers);
            if (widgets.Count == 0) {
                continue;
            }

            result.Add(BuildFormField(field, widgets));
        }

        return result.Count == 0 ? null : result.AsReadOnly();
    }

    private static int[] GetSelectedPageNumbers(IReadOnlyList<PdfLogicalPage> selectedPages, PdfLogicalPage? page) {
        if (page is not null) {
            return new[] { page.PageNumber };
        }

        var pageNumbers = new List<int>();
        for (int i = 0; i < selectedPages.Count; i++) {
            int pageNumber = selectedPages[i].PageNumber;
            if (!pageNumbers.Contains(pageNumber)) {
                pageNumbers.Add(pageNumber);
            }
        }

        return pageNumbers.ToArray();
    }

    private static IReadOnlyList<PdfFormWidget> GetWidgetsForPages(PdfFormField field, int[] pageNumbers) {
        if (field.Widgets.Count == 0 || pageNumbers.Length == 0) {
            return Array.Empty<PdfFormWidget>();
        }

        var widgets = new List<PdfFormWidget>();
        for (int i = 0; i < field.Widgets.Count; i++) {
            PdfFormWidget widget = field.Widgets[i];
            if (!widget.PageNumber.HasValue) {
                continue;
            }

            for (int j = 0; j < pageNumbers.Length; j++) {
                if (widget.PageNumber.Value == pageNumbers[j]) {
                    widgets.Add(widget);
                    break;
                }
            }
        }

        return widgets.Count == 0 ? Array.Empty<PdfFormWidget>() : widgets.AsReadOnly();
    }

    private static ReaderFormField BuildFormField(PdfFormField field, IReadOnlyList<PdfFormWidget> widgets) {
        return new ReaderFormField {
            Name = field.Name,
            PartialName = field.PartialName,
            AlternateName = field.AlternateName,
            MappingName = field.MappingName,
            FieldType = field.FieldType,
            Kind = ToReaderFormFieldKind(field.Kind),
            Value = field.Value,
            Values = field.Values,
            DefaultValue = field.DefaultValue,
            DefaultValues = field.DefaultValues,
            MaxLength = field.MaxLength,
            IsReadOnly = field.IsReadOnly,
            IsRequired = field.IsRequired,
            IsNoExport = field.IsNoExport,
            IsMultiline = field.IsMultiline,
            IsPassword = field.IsPassword,
            IsComb = field.IsComb,
            OptionCount = field.OptionCount,
            SelectedOptionCount = field.SelectedOptionCount,
            WidgetCount = widgets.Count,
            PageNumbers = GetWidgetPageNumbers(widgets),
            Widgets = BuildFormWidgets(widgets)
        };
    }

    private static IReadOnlyList<int> GetWidgetPageNumbers(IReadOnlyList<PdfFormWidget> widgets) {
        var pageNumbers = new List<int>();
        for (int i = 0; i < widgets.Count; i++) {
            int? pageNumber = widgets[i].PageNumber;
            if (pageNumber.HasValue && !pageNumbers.Contains(pageNumber.Value)) {
                pageNumbers.Add(pageNumber.Value);
            }
        }

        return pageNumbers.Count == 0 ? Array.Empty<int>() : pageNumbers.AsReadOnly();
    }

    private static IReadOnlyList<ReaderFormWidget> BuildFormWidgets(IReadOnlyList<PdfFormWidget> widgets) {
        if (widgets.Count == 0) return Array.Empty<ReaderFormWidget>();

        var result = new ReaderFormWidget[widgets.Count];
        for (int i = 0; i < widgets.Count; i++) {
            PdfFormWidget widget = widgets[i];
            result[i] = new ReaderFormWidget {
                FieldName = widget.FieldName,
                PageNumber = widget.PageNumber,
                X1 = widget.X1,
                Y1 = widget.Y1,
                X2 = widget.X2,
                Y2 = widget.Y2,
                Width = widget.Width,
                Height = widget.Height,
                AppearanceState = widget.AppearanceState,
                IsHidden = widget.IsHidden,
                IsPrint = widget.IsPrint,
                IsReadOnly = widget.IsReadOnly,
                NormalAppearanceStateCount = widget.NormalAppearanceStateCount,
                NormalAppearanceStates = widget.NormalAppearanceStates
            };
        }

        return Array.AsReadOnly(result);
    }

    private static ReaderFormFieldKind ToReaderFormFieldKind(PdfFormFieldKind kind) {
        switch (kind) {
            case PdfFormFieldKind.Text:
                return ReaderFormFieldKind.Text;
            case PdfFormFieldKind.Button:
                return ReaderFormFieldKind.Button;
            case PdfFormFieldKind.Choice:
                return ReaderFormFieldKind.Choice;
            case PdfFormFieldKind.Signature:
                return ReaderFormFieldKind.Signature;
            default:
                return ReaderFormFieldKind.Unknown;
        }
    }

    private static IReadOnlyList<TextPart> SplitText(string text, int maxChars) {
        if (string.IsNullOrWhiteSpace(text)) return Array.Empty<TextPart>();
        if (maxChars <= 0 || text.Length <= maxChars) return new[] { new TextPart(text.Trim(), null) };

        var parts = new List<TextPart>();
        int index = 0;
        while (index < text.Length) {
            int remaining = text.Length - index;
            int take = Math.Min(maxChars, remaining);
            if (take < remaining) {
                int splitAt = FindSplit(text, index, take);
                if (splitAt > index) {
                    take = splitAt - index;
                }
            }

            string segment = text.Substring(index, take).Trim();
            if (segment.Length > 0) {
                parts.Add(new TextPart(segment, new[] { "PDF content was split due to MaxChars." }));
            }

            index += take;
            while (index < text.Length && char.IsWhiteSpace(text[index])) {
                index++;
            }
        }

        return parts;
    }

    private static int FindSplit(string text, int index, int take) {
        int end = Math.Min(text.Length, index + take);
        for (int i = end - 1; i > index; i--) {
            char ch = text[i];
            if (ch == '\n' || ch == '\r' || char.IsWhiteSpace(ch)) {
                return i;
            }
        }

        return end;
    }

    private static int GetMaxTableRows(ReaderOptions readerOptions) {
        return readerOptions.MaxTableRows > 0 ? readerOptions.MaxTableRows : 0;
    }

    private static IReadOnlyList<ReaderTable>? BuildTables(IReadOnlyList<PdfLogicalTableExtraction> tables, int? pageSelectionIndex = null) {
        if (tables.Count == 0) return null;

        var result = new List<ReaderTable>(tables.Count);
        for (int i = 0; i < tables.Count; i++) {
            PdfLogicalTableExtraction table = tables[i];
            PdfLogicalTableData data = table.Data;
            int selectionIndex = pageSelectionIndex ?? table.PageIndex;

            result.Add(new ReaderTable {
                Kind = table.DetectionKind,
                Location = new ReaderLocation {
                    Page = table.PageNumber,
                    TableIndex = table.TableIndex,
                    SourceBlockIndex = selectionIndex,
                    SourceBlockKind = "table",
                    BlockAnchor = "page-" + table.PageNumber.ToString(CultureInfo.InvariantCulture)
                        + "-selection-" + selectionIndex.ToString("D4", CultureInfo.InvariantCulture)
                        + "-table-" + table.TableIndex.ToString(CultureInfo.InvariantCulture)
                },
                Columns = data.Columns,
                ColumnProfiles = BuildColumnProfiles(data),
                Diagnostics = BuildTableDiagnostics(data.Diagnostics),
                Rows = data.Rows,
                TotalRowCount = data.TotalRowCount,
                Truncated = data.Truncated
            });
        }

        return result;
    }

    private static ReaderTableDiagnostics BuildTableDiagnostics(PdfLogicalTableDiagnostics source) {
        return new ReaderTableDiagnostics {
            Confidence = source.Confidence,
            SchemaConfidence = source.SchemaConfidence,
            CellCompleteness = source.CellCompleteness,
            ColumnGeometryConfidence = source.ColumnGeometryConfidence,
            SourceRowCount = source.SourceRowCount,
            ExpectedCellCount = source.ExpectedCellCount,
            FilledCellCount = source.FilledCellCount,
            MissingCellCount = source.MissingCellCount,
            XStart = source.XStart,
            XEnd = source.XEnd,
            YTop = source.YTop,
            YBottom = source.YBottom,
            Width = source.Width,
            Height = source.Height,
            HasGeometry = source.HasGeometry
        };
    }

    private static IReadOnlyList<ReaderTableColumnProfile> BuildColumnProfiles(PdfLogicalTableData data) {
        if (data.ColumnProfiles.Count == 0) {
            return Array.Empty<ReaderTableColumnProfile>();
        }

        var profiles = new ReaderTableColumnProfile[data.ColumnProfiles.Count];
        for (int i = 0; i < data.ColumnProfiles.Count; i++) {
            PdfLogicalTableColumnProfile source = data.ColumnProfiles[i];
            profiles[i] = new ReaderTableColumnProfile {
                Index = source.Index,
                Name = source.Name,
                Kind = ToReaderTableColumnKind(source.Kind),
                NonEmptyCellCount = source.NonEmptyCellCount,
                NumericCellCount = source.NumericCellCount,
                Confidence = source.Confidence
            };
        }

        return Array.AsReadOnly(profiles);
    }

    private static ReaderTableColumnKind ToReaderTableColumnKind(PdfLogicalTableColumnKind kind) {
        switch (kind) {
            case PdfLogicalTableColumnKind.Empty:
                return ReaderTableColumnKind.Empty;
            case PdfLogicalTableColumnKind.Numeric:
                return ReaderTableColumnKind.Numeric;
            case PdfLogicalTableColumnKind.Text:
                return ReaderTableColumnKind.Text;
            case PdfLogicalTableColumnKind.Mixed:
                return ReaderTableColumnKind.Mixed;
            default:
                return ReaderTableColumnKind.Mixed;
        }
    }

    private static PdfLogicalDocument LoadDocument(PdfDocument document, ReaderPdfOptions options) {
        if (document is null) throw new ArgumentNullException(nameof(document));
        var ranges = options.PageRanges?.ToArray();
        return ranges is { Length: > 0 }
            ? document.Read.Logical(PdfPageSelection.FromRanges(ranges), options.LayoutOptions)
            : document.Read.Logical(options.LayoutOptions);
    }

    private static PdfReadOptions? CreatePdfReadOptions(ReaderOptions options) {
        return options.MaxInputBytes.HasValue
            ? new PdfReadOptions {
                Limits = new PdfReadLimits {
                    MaxInputBytes = options.MaxInputBytes.Value
                }
            }
            : null;
    }

    private static PdfDocument OpenReaderPdf(Stream stream, ReaderOptions options) {
        try {
            PdfReadOptions? readOptions = CreatePdfReadOptions(options);
            return ReaderInputLimits.TryGetOwnedSnapshotBytes(stream, out byte[] ownedBytes)
                ? PdfDocument.OpenOwned(ownedBytes, readOptions)
                : PdfDocument.Open(stream, readOptions);
        } catch (PdfReadLimitException exception) when (exception.Kind == PdfReadLimitKind.InputBytes) {
            throw new IOException(
                "Input exceeds MaxInputBytes (" +
                exception.Actual.ToString(CultureInfo.InvariantCulture) + " > " +
                exception.Limit.ToString(CultureInfo.InvariantCulture) + ").",
                exception);
        }
    }

    private static ReaderChunk EnrichChunk(ReaderChunk chunk, SourceMetadata source, bool computeHashes) {
        chunk.SourceId ??= source.SourceId;
        chunk.SourceHash ??= source.SourceHash;
        chunk.SourceLastWriteUtc ??= source.LastWriteUtc;
        chunk.SourceLengthBytes ??= source.LengthBytes;
        chunk.TokenEstimate ??= EstimateTokenCount(chunk.Markdown ?? chunk.Text);
        if (computeHashes && string.IsNullOrWhiteSpace(chunk.ChunkHash)) {
            chunk.ChunkHash = ComputeChunkHash(chunk);
        }

        return chunk;
    }

    private static int EstimateTokenCount(string? text) {
        var safeText = text ?? string.Empty;
        if (safeText.Length == 0) return 0;
        return Math.Max(1, (safeText.Length + 3) / 4);
    }

    private static string ComputeChunkHash(ReaderChunk chunk) {
        var data = string.Join("|",
            chunk.Kind.ToString(),
            chunk.SourceId ?? string.Empty,
            chunk.Location.Path ?? string.Empty,
            chunk.Location.SourceBlockKind ?? string.Empty,
            chunk.Location.BlockAnchor ?? string.Empty,
            chunk.Location.Page?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Text ?? string.Empty,
            chunk.Markdown ?? string.Empty,
            BuildChunkMetadataHashInput(chunk));

        return ComputeSha256Hex(data);
    }

    private static SourceMetadata BuildSourceMetadataFromPath(string path, bool computeHash) {
        var normalizedPath = NormalizePathForId(path);
        var sourceId = BuildSourceId(normalizedPath);

        DateTime? lastWriteUtc = null;
        long? lengthBytes = null;
        try {
            var fileInfo = new FileInfo(path);
            if (fileInfo.Exists) {
                lastWriteUtc = fileInfo.LastWriteTimeUtc;
                lengthBytes = fileInfo.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        return new SourceMetadata {
            Path = path,
            SourceId = sourceId,
            SourceHash = computeHash ? TryComputeFileSha256(path) : null,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    private static void UpdateSourceMetadataFromPdfDocument(SourceMetadata source, PdfDocument document, bool computeHash) {
        byte[] bytes = document.GetBytesForOperation();
        source.LengthBytes = bytes.LongLength;
        if (computeHash) {
            source.SourceHash ??= ComputeSha256Hex(bytes);
        }
    }

    private static string NormalizeLogicalSourceName(string? sourceName, string fallback) {
        if (sourceName is null) return fallback;
        string trimmed = sourceName.Trim();
        return trimmed.Length == 0 ? fallback : trimmed;
    }

    private static string? TryComputeFileSha256(string path) {
        try {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return ComputeSha256Hex(fs);
        } catch {
            return null;
        }
    }

    private static string ComputeSha256Hex(string value) {
        using var sha = System.Security.Cryptography.SHA256.Create();
        var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
        var hash = sha.ComputeHash(bytes);
        return ConvertToHexLower(hash);
    }

    private static string ComputeSha256Hex(byte[] bytes) {
        using var sha = System.Security.Cryptography.SHA256.Create();
        var hash = sha.ComputeHash(bytes ?? Array.Empty<byte>());
        return ConvertToHexLower(hash);
    }

    private static string ComputeSha256Hex(Stream stream) {
        using var sha = System.Security.Cryptography.SHA256.Create();
        var hash = sha.ComputeHash(stream);
        return ConvertToHexLower(hash);
    }

    private static string ConvertToHexLower(byte[] bytes) {
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }

    private static string BuildSourceId(string sourceKey) {
        var normalized = sourceKey ?? string.Empty;
        if (Path.DirectorySeparatorChar == '\\') {
            normalized = normalized.ToLowerInvariant();
        }

        return "src:" + ComputeSha256Hex(normalized);
    }

    private static string NormalizePathForId(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;

        string fullPath;
        try {
            fullPath = Path.GetFullPath(path);
        } catch {
            fullPath = path;
        }

        return fullPath.Replace('\\', '/');
    }

    private sealed class SourceMetadata {
        public string Path { get; set; } = string.Empty;
        public string SourceId { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }

    private sealed class TextPart {
        public TextPart(string text, IReadOnlyList<string>? warnings) {
            Text = text;
            Warnings = warnings;
        }

        public string Text { get; }
        public IReadOnlyList<string>? Warnings { get; }
    }
}
