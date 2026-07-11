namespace OfficeIMO.Pdf;

/// <summary>A point in visual page coordinates whose origin is the rendered page's top-left corner.</summary>
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
