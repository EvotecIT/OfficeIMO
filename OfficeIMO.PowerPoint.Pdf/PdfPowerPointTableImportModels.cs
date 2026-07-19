using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

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

/// <summary>Reports the detected tables imported from a logical PDF into a PowerPoint presentation.</summary>
public sealed class PdfPowerPointTableImportReport {
    internal PdfPowerPointTableImportReport(
        IReadOnlyList<PdfPowerPointTableImportEntry> entries,
        OfficeIMO.Pdf.PdfTableExtractionScopeReport sourceScope) {
        Entries = Array.AsReadOnly((entries ?? throw new ArgumentNullException(nameof(entries))).ToArray());
        SourceScope = sourceScope ?? throw new ArgumentNullException(nameof(sourceScope));
    }

    /// <summary>Gets a snapshot of imported table segment metadata.</summary>
    public IReadOnlyList<PdfPowerPointTableImportEntry> Entries { get; }

    /// <summary>Gets source-page content that was outside this table-only import.</summary>
    public OfficeIMO.Pdf.PdfTableExtractionScopeReport SourceScope { get; }

    /// <summary>Gets whether the source contained page content outside the imported tables.</summary>
    public bool HasOmittedPageContent => SourceScope.HasOmittedPageContent;

    /// <summary>Gets whether any detected source table was truncated by the configured row limit.</summary>
    public bool HasLoss => Entries.Any(static entry => entry.Truncated);

    /// <summary>Throws when at least one detected source table was truncated.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException("PDF table import to PowerPoint truncated one or more detected source tables.");
    }
}

/// <summary>Contains an editable PowerPoint presentation and the corresponding PDF table import report.</summary>
public sealed class PdfPowerPointTableImportResult {
    internal PdfPowerPointTableImportResult(PptCore.PowerPointPresentation value, PdfPowerPointTableImportReport report) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>Gets the generated editable PowerPoint presentation. The caller owns and disposes it.</summary>
    public PptCore.PowerPointPresentation Value { get; }

    /// <summary>Gets the immutable table import report.</summary>
    public PdfPowerPointTableImportReport Report { get; }

    /// <summary>Gets whether the import truncated content within a detected source table.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Gets whether the source contained page content outside the imported tables.</summary>
    public bool HasOmittedPageContent => Report.HasOmittedPageContent;

    /// <summary>Returns the generated editable PowerPoint presentation.</summary>
    public PptCore.PowerPointPresentation RequireValue() => Value;

    /// <summary>Returns the generated editable presentation only when no detected table was truncated.</summary>
    public PptCore.PowerPointPresentation RequireNoLoss() {
        Report.RequireNoLoss();
        return Value;
    }
}
