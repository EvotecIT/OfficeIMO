using System;

namespace OfficeIMO.Drawing;

/// <summary>Reversible projective mapping between original and corrected image pixel-edge coordinates.</summary>
public sealed class OfficeScanPerspectiveMap {
    private readonly double[] _forward, _inverse;
    internal OfficeScanPerspectiveMap(OfficeScanPerspectiveOptions options, int sourceWidth, int sourceHeight, int width, int height) {
        SourceWidth = sourceWidth; SourceHeight = sourceHeight; Width = width; Height = height;
        OfficePoint a = options.TopLeft, b = options.TopRight, c = options.BottomRight, d = options.BottomLeft;
        double dx1 = b.X - c.X, dx2 = d.X - c.X, dy1 = b.Y - c.Y, dy2 = d.Y - c.Y;
        double dx3 = a.X - b.X + c.X - d.X, dy3 = a.Y - b.Y + c.Y - d.Y;
        double determinant = dx1 * dy2 - dx2 * dy1;
        if (Math.Abs(determinant) < 1E-12) throw new ArgumentException("Perspective corners cannot be inverted.");
        double g = (dx3 * dy2 - dx2 * dy3) / determinant;
        double h = (dx1 * dy3 - dx3 * dy1) / determinant;
        _inverse = new[] { b.X - a.X + g * b.X, d.X - a.X + h * d.X, a.X,
            b.Y - a.Y + g * b.Y, d.Y - a.Y + h * d.Y, a.Y, g, h, 1D };
        // A projective pole must not cross any part of the output rectangle.
        if (Math.Min(Math.Min(1D, 1D + g), Math.Min(1D + h, 1D + g + h)) <= 1E-8)
            throw new ArgumentException("Perspective corners produce an unstable mapping.");
        double[] m = _inverse;
        _forward = new[] { m[4]*m[8]-m[5]*m[7], m[2]*m[7]-m[1]*m[8], m[1]*m[5]-m[2]*m[4],
            m[5]*m[6]-m[3]*m[8], m[0]*m[8]-m[2]*m[6], m[2]*m[3]-m[0]*m[5],
            m[3]*m[7]-m[4]*m[6], m[1]*m[6]-m[0]*m[7], m[0]*m[4]-m[1]*m[3] };
        double det = m[0] * _forward[0] + m[1] * _forward[3] + m[2] * _forward[6];
        if (Math.Abs(det) < 1E-12) throw new ArgumentException("Perspective corners produce a singular mapping.");
    }
    /// <summary>Original image width in pixels.</summary>
    public int SourceWidth { get; }
    /// <summary>Original image height in pixels.</summary>
    public int SourceHeight { get; }
    /// <summary>Corrected image width in pixels.</summary>
    public int Width { get; }
    /// <summary>Corrected image height in pixels.</summary>
    public int Height { get; }
    /// <summary>Maps corrected pixel-edge coordinates back to the original image.</summary>
    public OfficePoint MapProcessedToSource(OfficePoint point) => Map(_inverse, point.X / Width, point.Y / Height, SourceWidth, SourceHeight);
    /// <summary>Maps original pixel-edge coordinates into the corrected image. Points on a projective pole are rejected.</summary>
    public OfficePoint MapSourceToProcessed(OfficePoint point) => Map(_forward, point.X / SourceWidth, point.Y / SourceHeight, Width, Height);
    private static OfficePoint Map(double[] m, double x, double y, int width, int height) {
        double w = m[6] * x + m[7] * y + m[8];
        if (double.IsNaN(w) || double.IsInfinity(w) || Math.Abs(w) < 1E-12)
            throw new ArgumentOutOfRangeException(nameof(x), "Point lies outside the finite projective coordinate domain.");
        return new OfficePoint((m[0] * x + m[1] * y + m[2]) / w * width, (m[3] * x + m[4] * y + m[5]) / w * height);
    }
}