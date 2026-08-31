namespace OfficeIMO.Pdf;

/// <summary>Immutable point in PDF default user-space coordinates measured from the page bottom-left.</summary>
public readonly struct PdfPagePoint {
    /// <summary>Creates a PDF page point.</summary>
    public PdfPagePoint(double x, double y) {
        if (!IsFinite(x)) throw new ArgumentOutOfRangeException(nameof(x), "PDF page point coordinates must be finite.");
        if (!IsFinite(y)) throw new ArgumentOutOfRangeException(nameof(y), "PDF page point coordinates must be finite.");
        X = x;
        Y = y;
    }

    /// <summary>Horizontal coordinate in PDF default user-space units.</summary>
    public double X { get; }

    /// <summary>Vertical coordinate in PDF default user-space units.</summary>
    public double Y { get; }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
