namespace OfficeIMO.Pdf;

/// <summary>Immutable rectangle in PDF default user-space coordinates measured from the page bottom-left.</summary>
public sealed class PdfPageRectangle {
    /// <summary>Creates a normalized PDF page rectangle.</summary>
    public PdfPageRectangle(double left, double bottom, double right, double top) {
        if (!IsFinite(left) || !IsFinite(bottom) || !IsFinite(right) || !IsFinite(top)) {
            throw new ArgumentOutOfRangeException(nameof(left), "PDF page rectangle coordinates must be finite.");
        }
        if (right <= left) throw new ArgumentOutOfRangeException(nameof(right), "Rectangle right must be greater than left.");
        if (top <= bottom) throw new ArgumentOutOfRangeException(nameof(top), "Rectangle top must be greater than bottom.");
        Left = left;
        Bottom = bottom;
        Right = right;
        Top = top;
    }

    /// <summary>Left edge in PDF default user-space units.</summary>
    public double Left { get; }

    /// <summary>Bottom edge in PDF default user-space units.</summary>
    public double Bottom { get; }

    /// <summary>Right edge in PDF default user-space units.</summary>
    public double Right { get; }

    /// <summary>Top edge in PDF default user-space units.</summary>
    public double Top { get; }

    /// <summary>Rectangle width in PDF default user-space units.</summary>
    public double Width => Right - Left;

    /// <summary>Rectangle height in PDF default user-space units.</summary>
    public double Height => Top - Bottom;

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
