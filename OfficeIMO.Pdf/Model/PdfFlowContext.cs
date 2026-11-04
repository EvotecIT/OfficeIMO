namespace OfficeIMO.Pdf;

/// <summary>Read-only layout context supplied to conditional and replayable flow callbacks.</summary>
public sealed class PdfFlowContext {
    internal PdfFlowContext(int pageNumber, double availableHeight, double fullContentHeight, double contentWidth, double pageWidth, double pageHeight, bool isAtPageTop) {
        PageNumber = pageNumber;
        AvailableHeight = global::OfficeIMO.OfficeLength.Points(availableHeight);
        FullContentHeight = global::OfficeIMO.OfficeLength.Points(fullContentHeight);
        ContentWidth = global::OfficeIMO.OfficeLength.Points(contentWidth);
        PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth);
        PageHeight = global::OfficeIMO.OfficeLength.Points(pageHeight);
        IsAtPageTop = isAtPageTop;
    }

    /// <summary>Current one-based physical output page number.</summary>
    public int PageNumber { get; }
    /// <summary>Remaining vertical content space in points.</summary>
    public global::OfficeIMO.OfficeLength AvailableHeight { get; }
    /// <summary>Total vertical content space in the current flow frame on a new page, after active container top padding.</summary>
    public global::OfficeIMO.OfficeLength FullContentHeight { get; }
    /// <summary>Available content width in points.</summary>
    public global::OfficeIMO.OfficeLength ContentWidth { get; }
    /// <summary>Current page width in points.</summary>
    public global::OfficeIMO.OfficeLength PageWidth { get; }
    /// <summary>Current page height in points.</summary>
    public global::OfficeIMO.OfficeLength PageHeight { get; }
    /// <summary>True when layout is at the top of the current flow frame.</summary>
    public bool IsAtPageTop { get; }
}
