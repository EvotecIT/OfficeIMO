namespace OfficeIMO.Pdf;

/// <summary>A point in visual page coordinates whose origin is the rendered page's top-left corner.</summary>
/// <remarks>Visual coordinates are numeric physical points (1/72 inch), with a top-left origin after crop, rotation and /UserUnit scaling. Use the page coordinate mapping methods when calling native read/edit APIs.</remarks>
public readonly struct PdfSelectionPoint {
    /// <summary>Creates a point.</summary>
    public PdfSelectionPoint(double x, double y) {
        X = x;
        Y = y;
    }

    /// <summary>Horizontal coordinate in PDF points.</summary>
    public double X { get; }

    /// <summary>Vertical coordinate in PDF points, increasing downward.</summary>
    public double Y { get; }
}
