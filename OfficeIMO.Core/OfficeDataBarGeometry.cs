namespace OfficeIMO.Drawing;

/// <summary>
/// Resolved geometry for a proportional data bar inside a rectangular region.
/// </summary>
public readonly struct OfficeDataBarGeometry {
    internal OfficeDataBarGeometry(double x, double y, double width, double height) {
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    /// <summary>Left edge of the resolved data bar.</summary>
    public double X { get; }

    /// <summary>Top or bottom edge of the resolved data bar, depending on the target coordinate system.</summary>
    public double Y { get; }

    /// <summary>Resolved data-bar width.</summary>
    public double Width { get; }

    /// <summary>Resolved data-bar height.</summary>
    public double Height { get; }

    /// <summary>Indicates whether the resolved data bar has visible area.</summary>
    public bool IsVisible => Width > 0D && Height > 0D;
}
