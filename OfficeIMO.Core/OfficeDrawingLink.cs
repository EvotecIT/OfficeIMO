using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Non-painting interactive link rectangle inside an <see cref="OfficeDrawing"/> scene.
/// Renderers that support interaction preserve it while raster renderers safely ignore it.
/// </summary>
public sealed class OfficeDrawingLink : OfficeDrawingElement {
    /// <summary>Creates a drawing link over a local rectangle.</summary>
    public OfficeDrawingLink(string uri, double x, double y, double width, double height, string? alternativeText = null) {
        if (string.IsNullOrWhiteSpace(uri)) throw new ArgumentException("Drawing link URI cannot be empty.", nameof(uri));
        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
        ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
        Uri = uri!.Trim();
        X = x;
        Y = y;
        Width = width;
        Height = height;
        AlternativeText = string.IsNullOrWhiteSpace(alternativeText) ? null : alternativeText!.Trim();
    }

    /// <summary>URI or local fragment target.</summary>
    public string Uri { get; }

    /// <summary>Left coordinate in drawing units.</summary>
    public double X { get; }

    /// <summary>Top coordinate in drawing units.</summary>
    public double Y { get; }

    /// <summary>Interactive width in drawing units.</summary>
    public double Width { get; }

    /// <summary>Interactive height in drawing units.</summary>
    public double Height { get; }

    /// <summary>Optional accessible description.</summary>
    public string? AlternativeText { get; }

    internal override OfficeDrawingElement CloneElement() =>
        new OfficeDrawingLink(Uri, X, Y, Width, Height, AlternativeText);

    private static void ValidateFinite(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name);
    }

    private static void ValidatePositive(double value, string name) {
        ValidateFinite(value, name);
        if (value <= 0D) throw new ArgumentOutOfRangeException(name);
    }
}
