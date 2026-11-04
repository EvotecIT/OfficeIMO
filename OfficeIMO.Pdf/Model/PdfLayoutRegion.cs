namespace OfficeIMO.Pdf;

/// <summary>One captured rectangular flow region on an output page.</summary>
public sealed class PdfLayoutRegion {
    internal PdfLayoutRegion(int pageNumber, double x, double y, double width, double height) {
        PageNumber = pageNumber;
        X = global::OfficeIMO.OfficeLength.Points(x);
        Y = global::OfficeIMO.OfficeLength.Points(y);
        Width = global::OfficeIMO.OfficeLength.Points(width);
        Height = global::OfficeIMO.OfficeLength.Points(height);
    }

    /// <summary>One-based physical output page number.</summary>
    public int PageNumber { get; }
    /// <summary>Left coordinate in PDF points.</summary>
    public global::OfficeIMO.OfficeLength X { get; }
    /// <summary>Bottom coordinate in PDF points.</summary>
    public global::OfficeIMO.OfficeLength Y { get; }
    /// <summary>Region width in PDF points.</summary>
    public global::OfficeIMO.OfficeLength Width { get; }
    /// <summary>Region height in PDF points.</summary>
    public global::OfficeIMO.OfficeLength Height { get; }
}
