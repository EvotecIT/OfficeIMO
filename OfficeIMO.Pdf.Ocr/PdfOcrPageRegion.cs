namespace OfficeIMO.Pdf.Ocr;

/// <summary>One OCR rectangle in normalized top-left visual page coordinates, after the PDF crop and page rotation.</summary>
public sealed class PdfOcrPageRegion {
    /// <summary>Creates a nonempty rectangle inside one one-based page.</summary>
    public PdfOcrPageRegion(int pageNumber, double x, double y, double width, double height) {
        if (pageNumber < 1) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        if (!Finite(x) || !Finite(y) || !Finite(width) || !Finite(height) || x < 0 || y < 0 ||
            width <= 0 || height <= 0 || x + width > 1 || y + height > 1)
            throw new ArgumentException("OCR regions must be nonempty normalized rectangles inside the visual page.");
        PageNumber = pageNumber; X = x; Y = y; Width = width; Height = height;
    }
    /// <summary>One-based source page.</summary>
    public int PageNumber { get; }
    /// <summary>Left edge, normalized to visual page width.</summary>
    public double X { get; }
    /// <summary>Top edge, normalized to visual page height.</summary>
    public double Y { get; }
    /// <summary>Width, normalized to visual page width.</summary>
    public double Width { get; }
    /// <summary>Height, normalized to visual page height.</summary>
    public double Height { get; }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}