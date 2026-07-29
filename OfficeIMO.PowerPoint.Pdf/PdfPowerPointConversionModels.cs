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
        bool headerRowIncluded) {
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
}

/// <summary>Reports a PDF-to-PowerPoint conversion in visual-page or editable-table mode.</summary>
public sealed class PdfPowerPointConversionReport {
    internal PdfPowerPointConversionReport(
        IReadOnlyList<PdfPowerPointTableImportEntry> entries,
        OfficeIMO.Pdf.PdfTableExtractionScopeReport sourceScope) {
        Mode = PdfPowerPointImportMode.EditableTables;
        TableEntries = Array.AsReadOnly((entries ?? throw new ArgumentNullException(nameof(entries))).ToArray());
        SourceScope = sourceScope ?? throw new ArgumentNullException(nameof(sourceScope));
        VisualPages = Array.Empty<PdfPowerPointVisualPageEntry>();
    }

    internal PdfPowerPointConversionReport(IReadOnlyList<PdfPowerPointVisualPageEntry> visualPages) {
        Mode = PdfPowerPointImportMode.VisualPages;
        TableEntries = Array.Empty<PdfPowerPointTableImportEntry>();
        VisualPages = Array.AsReadOnly((visualPages ?? throw new ArgumentNullException(nameof(visualPages))).ToArray());
    }

    /// <summary>Gets the conversion strategy used for this operation.</summary>
    public PdfPowerPointImportMode Mode { get; }

    /// <summary>Gets a snapshot of imported table segment metadata.</summary>
    public IReadOnlyList<PdfPowerPointTableImportEntry> TableEntries { get; }

    /// <summary>Gets a snapshot of visual page-to-slide mappings.</summary>
    public IReadOnlyList<PdfPowerPointVisualPageEntry> VisualPages { get; }

    /// <summary>Gets source-page content that was outside this table-only import.</summary>
    public OfficeIMO.Pdf.PdfTableExtractionScopeReport? SourceScope { get; }

    /// <summary>Gets whether the source contained page content outside the imported tables.</summary>
    public bool HasOmittedPageContent => SourceScope?.HasOmittedPageContent == true;

    /// <summary>Gets whether table reconstruction truncated data or visual rendering reported a failure/simplification.</summary>
    public bool HasLoss =>
        TableEntries.Any(static entry => entry.Truncated) ||
        VisualPages.Any(static page =>
            !page.Succeeded ||
            page.CapabilityDiagnostics.Any(static diagnostic =>
                diagnostic.SupportLevel != OfficeIMO.Pdf.PdfRenderSupportLevel.Supported));

    /// <summary>Throws when the conversion reported possible content loss.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException("PDF-to-PowerPoint conversion reported possible content loss.");
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

    /// <summary>Returns the generated PowerPoint presentation.</summary>
    public PptCore.PowerPointPresentation RequireValue() => Value;

    /// <summary>Returns the generated presentation only when the selected conversion mode reported no loss.</summary>
    public PptCore.PowerPointPresentation RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
