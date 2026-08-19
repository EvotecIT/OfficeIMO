using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>Describes one PDF page projected onto a PowerPoint slide.</summary>
public sealed class PdfPowerPointVisualPageEntry {
    internal PdfPowerPointVisualPageEntry(OfficeIMO.Pdf.PdfPageRenderResult render, int slideIndex) {
        PageNumber = render.PageNumber;
        SlideIndex = slideIndex;
        Succeeded = render.Succeeded;
        PixelWidth = render.Width;
        PixelHeight = render.Height;
        Diagnostics = Array.AsReadOnly(render.Diagnostics.ToArray());
        CapabilityDiagnostics = Array.AsReadOnly(render.CapabilityDiagnostics.ToArray());
    }

    /// <summary>One-based source PDF page number.</summary>
    public int PageNumber { get; }
    /// <summary>Zero-based destination slide index.</summary>
    public int SlideIndex { get; }
    /// <summary>Whether the managed renderer produced page image bytes.</summary>
    public bool Succeeded { get; }
    /// <summary>Rendered page width in pixels.</summary>
    public int PixelWidth { get; }
    /// <summary>Rendered page height in pixels.</summary>
    public int PixelHeight { get; }
    /// <summary>Human-readable render diagnostics.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Typed managed-renderer capability diagnostics.</summary>
    public IReadOnlyList<OfficeIMO.Pdf.PdfRenderCapabilityDiagnostic> CapabilityDiagnostics { get; }
}

/// <summary>
/// Describes one logical PDF table imported into a PowerPoint slide.
/// </summary>
public sealed class PdfPowerPointTableImportEntry {
    internal PdfPowerPointTableImportEntry(
        int pageIndex,
        int pageNumber,
        int tableIndex,
        string detectionKind,
        int slideIndex,
        int segmentIndex,
        int segmentCount,
        int rowStartIndex,
        int columnStartIndex,
        int sourceColumnCount,
        int columnCount,
        int rowCount,
        int totalRowCount,
        bool truncated,
        bool headerRowIncluded,
        IReadOnlyList<int>? sourcePageNumbers = null,
        int sourceTableCount = 1,
        int suppressedRepeatedHeaderRows = 0,
        int additionalHeaderRowCount = 0) {
        PageIndex = pageIndex;
        PageNumber = pageNumber;
        TableIndex = tableIndex;
        DetectionKind = detectionKind ?? string.Empty;
        SlideIndex = slideIndex;
        SegmentIndex = segmentIndex;
        SegmentCount = segmentCount;
        RowStartIndex = rowStartIndex;
        ColumnStartIndex = columnStartIndex;
        SourceColumnCount = sourceColumnCount;
        ColumnCount = columnCount;
        RowCount = rowCount;
        TotalRowCount = totalRowCount;
        Truncated = truncated;
        HeaderRowIncluded = headerRowIncluded;
        SourcePageNumbers = Array.AsReadOnly((sourcePageNumbers ?? new[] { pageNumber }).ToArray());
        SourceTableCount = sourceTableCount;
        SuppressedRepeatedHeaderRows = suppressedRepeatedHeaderRows;
        AdditionalHeaderRowCount = additionalHeaderRowCount;
    }

    /// <summary>Zero-based page index within the selected logical page collection.</summary>
    public int PageIndex { get; }

    /// <summary>One-based source page number from the PDF document.</summary>
    public int PageNumber { get; }

    /// <summary>Zero-based table index within the source logical PDF page.</summary>
    public int TableIndex { get; }

    /// <summary>Detection heuristic that produced the imported table.</summary>
    public string DetectionKind { get; }

    /// <summary>Zero-based slide index where the table was written.</summary>
    public int SlideIndex { get; }

    /// <summary>Zero-based segment index for this source PDF table.</summary>
    public int SegmentIndex { get; }

    /// <summary>Total number of PowerPoint slide segments produced for this source PDF table.</summary>
    public int SegmentCount { get; }

    /// <summary>Zero-based body row index where this slide segment starts in the normalized PDF table.</summary>
    public int RowStartIndex { get; }

    /// <summary>Zero-based column index where this slide segment starts in the normalized PDF table.</summary>
    public int ColumnStartIndex { get; }

    /// <summary>Total detected source columns before any per-slide column split was applied.</summary>
    public int SourceColumnCount { get; }

    /// <summary>Number of imported columns.</summary>
    public int ColumnCount { get; }

    /// <summary>Number of body rows written to PowerPoint.</summary>
    public int RowCount { get; }

    /// <summary>Total body rows detected before any row cap was applied.</summary>
    public int TotalRowCount { get; }

    /// <summary>True when imported rows were truncated by the configured row cap.</summary>
    public bool Truncated { get; }

    /// <summary>True when a column-header row was written above the imported body rows.</summary>
    public bool HeaderRowIncluded { get; }

    /// <summary>One-based PDF page numbers contributing rows to this imported table.</summary>
    public IReadOnlyList<int> SourcePageNumbers { get; }

    /// <summary>Number of page-level table segments combined into this logical table.</summary>
    public int SourceTableCount { get; }

    /// <summary>Number of repeated continuation header rows omitted from body data.</summary>
    public int SuppressedRepeatedHeaderRows { get; }

    /// <summary>Number of repeated header rows appended to primary header labels.</summary>
    public int AdditionalHeaderRowCount { get; }
}

/// <summary>Describes editable objects reconstructed for one source PDF page.</summary>
public sealed class PdfPowerPointEditablePageEntry {
    internal PdfPowerPointEditablePageEntry(
        int pageNumber,
        int slideIndex,
        int textBoxCount,
        int tableCount,
        int shapeCount,
        int imageCount,
        int omittedTextCount,
        int omittedTableCount,
        int omittedVectorCount,
        int omittedImageCount) {
        PageNumber = pageNumber;
        SlideIndex = slideIndex;
        TextBoxCount = textBoxCount;
        TableCount = tableCount;
        ShapeCount = shapeCount;
        ImageCount = imageCount;
        OmittedTextCount = omittedTextCount;
        OmittedTableCount = omittedTableCount;
        OmittedVectorCount = omittedVectorCount;
        OmittedImageCount = omittedImageCount;
    }

    /// <summary>One-based source PDF page number.</summary>
    public int PageNumber { get; }
    /// <summary>Zero-based primary destination slide index.</summary>
    public int SlideIndex { get; }
    /// <summary>Editable text boxes created from non-table PDF text blocks.</summary>
    public int TextBoxCount { get; }
    /// <summary>Native PowerPoint tables created on the primary slide.</summary>
    public int TableCount { get; }
    /// <summary>Safe vector primitives reconstructed as PowerPoint shapes.</summary>
    public int ShapeCount { get; }
    /// <summary>Supported source image placements retained as separate pictures.</summary>
    public int ImageCount { get; }
    /// <summary>Text blocks omitted because the configured editable object limit was reached.</summary>
    public int OmittedTextCount { get; }
    /// <summary>Table segments omitted because the configured editable object limit was reached.</summary>
    public int OmittedTableCount { get; }
    /// <summary>Vector primitives that could not be represented safely as PowerPoint shapes.</summary>
    public int OmittedVectorCount { get; }
    /// <summary>Image resources or placements that could not be represented safely as PowerPoint pictures.</summary>
    public int OmittedImageCount { get; }
    /// <summary>Whether this page omitted supported semantic object categories.</summary>
    public bool HasOmittedContent =>
        OmittedTextCount > 0 ||
        OmittedTableCount > 0 ||
        OmittedVectorCount > 0 ||
        OmittedImageCount > 0;
}

/// <summary>Reports a PDF-to-PowerPoint conversion in visual, hybrid, table, or editable-content mode.</summary>
public sealed class PdfPowerPointConversionReport {
    private readonly bool _hasOmittedPageContent;

    internal PdfPowerPointConversionReport(
        IReadOnlyList<PdfPowerPointTableImportEntry> entries,
        OfficeIMO.Pdf.PdfTableExtractionScopeReport sourceScope) {
        Mode = PdfPowerPointImportMode.EditableTables;
        TableEntries = Array.AsReadOnly((entries ?? throw new ArgumentNullException(nameof(entries))).ToArray());
        SourceScope = sourceScope ?? throw new ArgumentNullException(nameof(sourceScope));
        VisualPages = Array.Empty<PdfPowerPointVisualPageEntry>();
        EditablePages = Array.Empty<PdfPowerPointEditablePageEntry>();
        _hasOmittedPageContent = SourceScope.HasOmittedPageContent;
        Warnings = CreateProjectionWarnings(SourceScope, failedVisualScope: null, hasFailedVisualPages: false);
    }

    internal PdfPowerPointConversionReport(IReadOnlyList<PdfPowerPointVisualPageEntry> visualPages) {
        Mode = PdfPowerPointImportMode.VisualPages;
        TableEntries = Array.Empty<PdfPowerPointTableImportEntry>();
        VisualPages = Array.AsReadOnly((visualPages ?? throw new ArgumentNullException(nameof(visualPages))).ToArray());
        EditablePages = Array.Empty<PdfPowerPointEditablePageEntry>();
        _hasOmittedPageContent = false;
        Warnings = CreateVisualPageWarnings(VisualPages);
    }

    internal PdfPowerPointConversionReport(
        IReadOnlyList<PdfPowerPointTableImportEntry> entries,
        IReadOnlyList<PdfPowerPointVisualPageEntry> visualPages,
        OfficeIMO.Pdf.PdfTableExtractionScopeReport sourceScope,
        OfficeIMO.Pdf.PdfTableExtractionScopeReport failedVisualScope) {
        Mode = PdfPowerPointImportMode.HybridVisualAndEditableTables;
        TableEntries = Array.AsReadOnly((entries ?? throw new ArgumentNullException(nameof(entries))).ToArray());
        VisualPages = Array.AsReadOnly((visualPages ?? throw new ArgumentNullException(nameof(visualPages))).ToArray());
        EditablePages = Array.Empty<PdfPowerPointEditablePageEntry>();
        SourceScope = sourceScope ?? throw new ArgumentNullException(nameof(sourceScope));
        if (failedVisualScope == null) throw new ArgumentNullException(nameof(failedVisualScope));
        bool hasFailedVisualPages = VisualPages.Any(static page => !page.Succeeded);
        _hasOmittedPageContent = hasFailedVisualPages &&
            (failedVisualScope.HasOmittedPageContent || SourceScope.OptionalContentGroupCount > 0);
        var warnings = new List<OfficeIMO.Pdf.PdfConversionWarning>(CreateProjectionWarnings(
            SourceScope,
            failedVisualScope,
            hasFailedVisualPages));
        AddRendererWarnings(warnings, VisualPages);
        Warnings = warnings.AsReadOnly();
    }

    internal PdfPowerPointConversionReport(
        IReadOnlyList<PdfPowerPointEditablePageEntry> editablePages,
        IReadOnlyList<PdfPowerPointTableImportEntry> tableEntries,
        OfficeIMO.Pdf.PdfTableExtractionScopeReport sourceScope,
        IReadOnlyList<OfficeIMO.Pdf.PdfConversionWarning> warnings) {
        Mode = PdfPowerPointImportMode.EditableContent;
        EditablePages = Array.AsReadOnly((editablePages ?? throw new ArgumentNullException(nameof(editablePages))).ToArray());
        TableEntries = Array.AsReadOnly((tableEntries ?? throw new ArgumentNullException(nameof(tableEntries))).ToArray());
        SourceScope = sourceScope ?? throw new ArgumentNullException(nameof(sourceScope));
        VisualPages = Array.Empty<PdfPowerPointVisualPageEntry>();
        Warnings = Array.AsReadOnly((warnings ?? throw new ArgumentNullException(nameof(warnings))).ToArray());
        _hasOmittedPageContent = EditablePages.Any(static page => page.HasOmittedContent)
            || Warnings.Any(static warning =>
                warning.Severity != OfficeIMO.Pdf.PdfConversionWarningSeverity.Information
                && warning.Details.TryGetValue("Disposition", out string? disposition)
                && string.Equals(disposition, "Omitted", StringComparison.Ordinal));
    }

    /// <summary>Gets the conversion strategy used for this operation.</summary>
    public PdfPowerPointImportMode Mode { get; }

    /// <summary>Gets a snapshot of imported table segment metadata.</summary>
    public IReadOnlyList<PdfPowerPointTableImportEntry> TableEntries { get; }

    /// <summary>Gets a snapshot of visual page-to-slide mappings.</summary>
    public IReadOnlyList<PdfPowerPointVisualPageEntry> VisualPages { get; }

    /// <summary>Gets editable object counts for semantic page projections.</summary>
    public IReadOnlyList<PdfPowerPointEditablePageEntry> EditablePages { get; }

    /// <summary>Gets source-page content that was outside this table-only import.</summary>
    public OfficeIMO.Pdf.PdfTableExtractionScopeReport? SourceScope { get; }

    /// <summary>Gets typed warnings for source content that is not editable in the selected projection.</summary>
    public IReadOnlyList<OfficeIMO.Pdf.PdfConversionWarning> Warnings { get; }

    /// <summary>Gets whether the source contained page content outside the imported tables.</summary>
    public bool HasOmittedPageContent => _hasOmittedPageContent;

    /// <summary>Gets whether source content exists outside editable table overlays, even when retained in the hybrid visual layer.</summary>
    public bool HasNonEditablePageContent => SourceScope?.HasOmittedPageContent == true;

    /// <summary>Gets whether the selected projection omitted, truncated, simplified, or failed to render source content.</summary>
    public bool HasLoss =>
        _hasOmittedPageContent ||
        TableEntries.Any(static entry => entry.Truncated) ||
        EditablePages.Any(static page => page.HasOmittedContent) ||
        VisualPages.Any(static page =>
            !page.Succeeded ||
            page.CapabilityDiagnostics.Any(static diagnostic =>
                diagnostic.SupportLevel != OfficeIMO.Pdf.PdfRenderSupportLevel.Supported));

    private static IReadOnlyList<OfficeIMO.Pdf.PdfConversionWarning> CreateVisualPageWarnings(
        IReadOnlyList<PdfPowerPointVisualPageEntry> visualPages) {
        var warnings = new List<OfficeIMO.Pdf.PdfConversionWarning> {
            new(
                "OfficeIMO.PowerPoint.Pdf",
                "PdfVisualPageSlidesNotEditable",
                "Slide content",
                "Each PDF page is retained as one page-sized image. Text, shapes, charts, and tables are not editable PowerPoint objects in this mode.",
                details: new Dictionary<string, string> {
                    ["Disposition"] = "VisualOnly",
                    ["construct"] = "Visual page slides"
                })
        };
        AddRendererWarnings(warnings, visualPages);
        return warnings.AsReadOnly();
    }

    private static void AddRendererWarnings(
        ICollection<OfficeIMO.Pdf.PdfConversionWarning> warnings,
        IReadOnlyList<PdfPowerPointVisualPageEntry> visualPages) {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int pageIndex = 0; pageIndex < visualPages.Count; pageIndex++) {
            PdfPowerPointVisualPageEntry page = visualPages[pageIndex];
            for (int diagnosticIndex = 0; diagnosticIndex < page.CapabilityDiagnostics.Count; diagnosticIndex++) {
                OfficeIMO.Pdf.PdfRenderCapabilityDiagnostic diagnostic = page.CapabilityDiagnostics[diagnosticIndex];
                string key = page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\0" + diagnostic.Code + "\0" + diagnostic.Subject;
                if (!seen.Add(key)) continue;
                warnings.Add(new OfficeIMO.Pdf.PdfConversionWarning(
                    "OfficeIMO.PowerPoint.Pdf",
                    diagnostic.Code,
                    "PDF page " + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    diagnostic.Message,
                    diagnostic.SupportLevel == OfficeIMO.Pdf.PdfRenderSupportLevel.Unsupported
                        ? OfficeIMO.Pdf.PdfConversionWarningSeverity.Warning
                        : OfficeIMO.Pdf.PdfConversionWarningSeverity.Information,
                    details: new Dictionary<string, string> {
                        ["pageNumber"] = page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        ["construct"] = diagnostic.Capability.Feature,
                        ["supportLevel"] = diagnostic.SupportLevel.ToString()
                    }));
            }
        }
    }

    /// <summary>Throws when the conversion reported possible content loss.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException("PDF-to-PowerPoint conversion reported possible content loss.");
    }

    private static IReadOnlyList<OfficeIMO.Pdf.PdfConversionWarning> CreateProjectionWarnings(
        OfficeIMO.Pdf.PdfTableExtractionScopeReport scope,
        OfficeIMO.Pdf.PdfTableExtractionScopeReport? failedVisualScope,
        bool hasFailedVisualPages) {
        var warnings = new List<OfficeIMO.Pdf.PdfConversionWarning>();
        bool hasVisualLayer = failedVisualScope != null;
        AddProjectionWarning(warnings, "PdfTextNotEditable", "Text", scope.NonTableTextBlockCount,
            failedVisualScope?.NonTableTextBlockCount ?? scope.NonTableTextBlockCount, hasVisualLayer, hasFailedVisualPages, pageCorrelationAvailable: true,
            description: "non-table text blocks");
        AddProjectionWarning(warnings, "PdfImagesNotEditable", "Images", scope.ImageCount,
            failedVisualScope?.ImageCount ?? scope.ImageCount, hasVisualLayer, hasFailedVisualPages, pageCorrelationAvailable: true,
            description: "images");
        AddProjectionWarning(warnings, "PdfVectorsNotEditable", "Vectors", scope.VectorPrimitiveCount,
            failedVisualScope?.VectorPrimitiveCount ?? scope.VectorPrimitiveCount, hasVisualLayer, hasFailedVisualPages, pageCorrelationAvailable: true,
            description: "vector primitives");
        AddProjectionWarning(warnings, "PdfNavigationNotEditable", "Navigation", scope.LinkCount + scope.PageActionCount,
            (failedVisualScope?.LinkCount ?? scope.LinkCount) + (failedVisualScope?.PageActionCount ?? scope.PageActionCount),
            hasVisualLayer, hasFailedVisualPages, pageCorrelationAvailable: true,
            description: "links and page actions");
        AddProjectionWarning(warnings, "PdfFormsAndControlsNotEditable", "Forms", scope.FormWidgetCount,
            failedVisualScope?.FormWidgetCount ?? scope.FormWidgetCount, hasVisualLayer, hasFailedVisualPages, pageCorrelationAvailable: true,
            description: "forms and interactive controls");
        AddProjectionWarning(warnings, "PdfAnnotationsNotEditable", "Annotations", scope.AnnotationCount,
            failedVisualScope?.AnnotationCount ?? scope.AnnotationCount, hasVisualLayer, hasFailedVisualPages, pageCorrelationAvailable: true,
            description: "annotations");
        AddProjectionWarning(warnings, "PdfGroupsNotEditable", "Groups", scope.OptionalContentGroupCount,
            failedCount: 0, hasVisualLayer, hasFailedVisualPages, pageCorrelationAvailable: false,
            description: "optional-content groups");
        AddProjectionWarning(warnings, "PdfAnimationsNotEditable", "Animations", scope.InteractiveMediaAnnotationCount,
            failedVisualScope?.InteractiveMediaAnnotationCount ?? scope.InteractiveMediaAnnotationCount,
            hasVisualLayer, hasFailedVisualPages, pageCorrelationAvailable: true,
            description: "interactive media and animations");
        if (scope.AnalysisTruncated) {
            ProjectionDisposition disposition = failedVisualScope == null
                ? ProjectionDisposition.Omitted
                : !hasFailedVisualPages || !failedVisualScope.AnalysisTruncated
                    ? ProjectionDisposition.VisualOnly
                    : ProjectionDisposition.Omitted;
            warnings.Add(new OfficeIMO.Pdf.PdfConversionWarning(
                "OfficeIMO.PowerPoint.Pdf",
                "PdfProjectionAnalysisTruncated",
                "Document",
                disposition == ProjectionDisposition.VisualOnly
                    ? "The bounded source-content analysis stopped before every text block could be classified; the unclassified content remains in the visual page layer."
                    : "The bounded source-content analysis stopped on content that is not retained in a visual page layer.",
                disposition == ProjectionDisposition.VisualOnly
                    ? OfficeIMO.Pdf.PdfConversionWarningSeverity.Information
                    : OfficeIMO.Pdf.PdfConversionWarningSeverity.Warning,
                details: new Dictionary<string, string> {
                    ["Disposition"] = GetDispositionValue(disposition)
                }));
        }

        return warnings.Count == 0
            ? Array.Empty<OfficeIMO.Pdf.PdfConversionWarning>()
            : Array.AsReadOnly(warnings.ToArray());
    }

    private static void AddProjectionWarning(
        ICollection<OfficeIMO.Pdf.PdfConversionWarning> warnings,
        string code,
        string source,
        int count,
        int failedCount,
        bool hasVisualLayer,
        bool hasFailedVisualPages,
        bool pageCorrelationAvailable,
        string description) {
        if (count <= 0) return;
        ProjectionDisposition disposition = ResolveDisposition(
            count,
            failedCount,
            hasVisualLayer,
            hasFailedVisualPages,
            pageCorrelationAvailable);
        warnings.Add(new OfficeIMO.Pdf.PdfConversionWarning(
            "OfficeIMO.PowerPoint.Pdf",
            code,
            source,
            "PDF " + description + " are not reconstructed as editable PowerPoint objects; " + GetDispositionMessage(disposition) + ".",
            disposition == ProjectionDisposition.VisualOnly
                ? OfficeIMO.Pdf.PdfConversionWarningSeverity.Information
                : OfficeIMO.Pdf.PdfConversionWarningSeverity.Warning,
            details: new Dictionary<string, string> {
                ["Count"] = count.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["Disposition"] = GetDispositionValue(disposition)
            }));
    }

    private static ProjectionDisposition ResolveDisposition(
        int count,
        int failedCount,
        bool hasVisualLayer,
        bool hasFailedVisualPages,
        bool pageCorrelationAvailable) {
        if (!hasVisualLayer) return ProjectionDisposition.Omitted;
        if (!hasFailedVisualPages) return ProjectionDisposition.VisualOnly;
        if (!pageCorrelationAvailable) return ProjectionDisposition.Unknown;
        if (failedCount <= 0) return ProjectionDisposition.VisualOnly;
        return failedCount >= count
            ? ProjectionDisposition.Omitted
            : ProjectionDisposition.PartiallyOmitted;
    }

    private static string GetDispositionMessage(ProjectionDisposition disposition) => disposition switch {
        ProjectionDisposition.VisualOnly => "they are retained in the visual page layer",
        ProjectionDisposition.Omitted => "they are omitted because no visual page layer retains them",
        ProjectionDisposition.PartiallyOmitted => "some are retained visually and some are omitted because their source pages did not render",
        _ => "their visual retention cannot be correlated to individual source pages after a render failure"
    };

    private static string GetDispositionValue(ProjectionDisposition disposition) => disposition switch {
        ProjectionDisposition.VisualOnly => "VisualOnly",
        ProjectionDisposition.Omitted => "Omitted",
        ProjectionDisposition.PartiallyOmitted => "PartiallyOmitted",
        _ => "Unknown"
    };

    private enum ProjectionDisposition {
        VisualOnly,
        Omitted,
        PartiallyOmitted,
        Unknown
    }
}

/// <summary>Contains a PowerPoint presentation and the corresponding PDF conversion report.</summary>
public sealed class PdfPowerPointConversionResult {
    internal PdfPowerPointConversionResult(PptCore.PowerPointPresentation value, PdfPowerPointConversionReport report) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Gets the generated PowerPoint presentation. The caller owns and disposes it.</summary>
    public PptCore.PowerPointPresentation Value { get; }

    /// <summary>Gets the immutable conversion report.</summary>
    public PdfPowerPointConversionReport Report { get; }

    /// <summary>Gets whether the conversion reported possible content loss.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Gets whether the source contained page content outside the imported tables.</summary>
    public bool HasOmittedPageContent => Report.HasOmittedPageContent;

    /// <summary>Gets typed warnings for content that was retained only visually or omitted by the selected projection.</summary>
    public IReadOnlyList<OfficeIMO.Pdf.PdfConversionWarning> Warnings => Report.Warnings;

    /// <summary>Returns the generated PowerPoint presentation.</summary>
    public PptCore.PowerPointPresentation RequireValue() => Value;

    /// <summary>Returns the generated presentation only when the selected conversion mode reported no loss.</summary>
    public PptCore.PowerPointPresentation RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
