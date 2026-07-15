namespace OfficeIMO.Pdf;

/// <summary>One captured rectangular flow region on an output page.</summary>
public sealed class PdfLayoutRegion {
    internal PdfLayoutRegion(int pageNumber, double x, double y, double width, double height) {
        PageNumber = pageNumber;
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    /// <summary>One-based physical output page number.</summary>
    public int PageNumber { get; }
    /// <summary>Left coordinate in PDF points.</summary>
    public double X { get; }
    /// <summary>Bottom coordinate in PDF points.</summary>
    public double Y { get; }
    /// <summary>Region width in PDF points.</summary>
    public double Width { get; }
    /// <summary>Region height in PDF points.</summary>
    public double Height { get; }
}
