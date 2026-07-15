namespace OfficeIMO.Pdf;

/// <summary>Read-only layout context supplied to conditional and replayable flow callbacks.</summary>
public sealed class PdfFlowContext {
    internal PdfFlowContext(int pageNumber, double availableHeight, double fullContentHeight, double contentWidth, double pageWidth, double pageHeight) {
        PageNumber = pageNumber;
        AvailableHeight = availableHeight;
        FullContentHeight = fullContentHeight;
        ContentWidth = contentWidth;
        PageWidth = pageWidth;
        PageHeight = pageHeight;
    }

    /// <summary>Current one-based physical output page number.</summary>
    public int PageNumber { get; }
    /// <summary>Remaining vertical content space in points.</summary>
    public double AvailableHeight { get; }
    /// <summary>Total vertical content space on an empty current page.</summary>
    public double FullContentHeight { get; }
    /// <summary>Available content width in points.</summary>
    public double ContentWidth { get; }
    /// <summary>Current page width in points.</summary>
    public double PageWidth { get; }
    /// <summary>Current page height in points.</summary>
    public double PageHeight { get; }
    /// <summary>True when layout is at the top of the current page.</summary>
    public bool IsAtPageTop => AvailableHeight >= FullContentHeight - 0.001D;
}
