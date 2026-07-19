namespace OfficeIMO.Pdf;

/// <summary>
/// Describes page content that is and is not in scope when logical PDF tables are extracted.
/// </summary>
public sealed class PdfTableExtractionScopeReport {
    internal PdfTableExtractionScopeReport(
        int sourcePageCount,
        int pagesWithTables,
        int detectedTableCount,
        int nonTableTextBlockCount,
        int imageCount,
        int linkCount,
        int formWidgetCount,
        int annotationCount,
        int pageActionCount) {
        SourcePageCount = sourcePageCount;
        PagesWithTables = pagesWithTables;
        DetectedTableCount = detectedTableCount;
        NonTableTextBlockCount = nonTableTextBlockCount;
        ImageCount = imageCount;
        LinkCount = linkCount;
        FormWidgetCount = formWidgetCount;
        AnnotationCount = annotationCount;
        PageActionCount = pageActionCount;
    }

    /// <summary>Number of logical source pages inspected.</summary>
    public int SourcePageCount { get; }

    /// <summary>Number of source pages on which at least one logical table was detected.</summary>
    public int PagesWithTables { get; }

    /// <summary>Number of source pages on which no logical table was detected.</summary>
    public int PagesWithoutTables => SourcePageCount - PagesWithTables;

    /// <summary>Total number of logical tables detected on the inspected pages.</summary>
    public int DetectedTableCount { get; }

    /// <summary>Number of visible text blocks that were not represented by a detected table.</summary>
    public int NonTableTextBlockCount { get; }

    /// <summary>Number of source images, which table-only adapters do not import.</summary>
    public int ImageCount { get; }

    /// <summary>Number of source link annotations, which table-only adapters do not import.</summary>
    public int LinkCount { get; }

    /// <summary>Number of source form widgets, which table-only adapters do not import.</summary>
    public int FormWidgetCount { get; }

    /// <summary>
    /// Number of generic source annotation records, which table-only adapters do not import.
    /// Link and widget annotations may also appear in their dedicated counts, so these counts are not additive.
    /// </summary>
    public int AnnotationCount { get; }

    /// <summary>Number of source page actions, which table-only adapters do not import.</summary>
    public int PageActionCount { get; }

    /// <summary>
    /// Gets whether visible or interactive page content existed outside the detected tables.
    /// This is expected for table-only extraction and is separate from truncation within a table.
    /// </summary>
    public bool HasOmittedPageContent =>
        NonTableTextBlockCount > 0 ||
        ImageCount > 0 ||
        LinkCount > 0 ||
        FormWidgetCount > 0 ||
        AnnotationCount > 0 ||
        PageActionCount > 0;
}
