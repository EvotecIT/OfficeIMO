namespace OfficeIMO.Pdf;

/// <summary>Read-only page information used to compose bounded running header or footer content.</summary>
public sealed class PdfRunningContentContext {
    internal PdfRunningContentContext(int pageNumber, int totalPages, int documentPageNumber,
        int documentPages, int sectionPageNumber, int sectionPages, double contentWidth, double pageWidth, double pageHeight,
        PdfPageNumberStyle pageNumberStyle) {
        PageNumber = pageNumber;
        TotalPages = totalPages;
        DocumentPageNumber = documentPageNumber;
        DocumentPages = documentPages;
        SectionPageNumber = sectionPageNumber;
        SectionPages = sectionPages;
        ContentWidth = global::OfficeIMO.OfficeLength.Points(contentWidth);
        PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth);
        PageHeight = global::OfficeIMO.OfficeLength.Points(pageHeight);
        PageNumberStyle = pageNumberStyle;
    }

    /// <summary>Current visible page number, including an explicitly restarted numbering sequence.</summary>
    public int PageNumber { get; }
    /// <summary>Last visible page number in the current numbering sequence, matching the existing {pages} token.</summary>
    public int TotalPages { get; }
    /// <summary>Current one-based physical output page number.</summary>
    public int DocumentPageNumber { get; }
    /// <summary>Total physical pages in the document.</summary>
    public int DocumentPages { get; }
    /// <summary>Current one-based physical page within the tracked source section, or the PDF page group when no source sections are tracked.</summary>
    public int SectionPageNumber { get; }
    /// <summary>Physical page count in the tracked source section or PDF page group, independent of visible numbering and restarts.</summary>
    public int SectionPages { get; }
    /// <summary>Width of the page's content frame in points, including mirrored margins.</summary>
    public global::OfficeIMO.OfficeLength ContentWidth { get; }
    /// <summary>Physical page width in points.</summary>
    public global::OfficeIMO.OfficeLength PageWidth { get; }
    /// <summary>Physical page height in points.</summary>
    public global::OfficeIMO.OfficeLength PageHeight { get; }
    /// <summary>Number style selected for this page's visible numbering sequence.</summary>
    public PdfPageNumberStyle PageNumberStyle { get; }
    /// <summary>Formats a positive page count using the supplied style or this page's numbering style.</summary>
    public string FormatPageNumber(int number, PdfPageNumberStyle? style = null) =>
        PdfPageNumberFormatter.Format(number, style ?? PageNumberStyle);
}
