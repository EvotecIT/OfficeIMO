using System;

namespace OfficeIMO.Drawing;

/// <summary>Four explicitly reviewed page corners for perspective correction, normalized to the source image edges.</summary>
public sealed class OfficeScanPerspectiveOptions {
    /// <summary>Top-left corner; both coordinates are from zero through one.</summary>
    public OfficePoint TopLeft { get; set; } = new OfficePoint(0, 0);
    /// <summary>Top-right corner.</summary>
    public OfficePoint TopRight { get; set; } = new OfficePoint(1, 0);
    /// <summary>Bottom-right corner.</summary>
    public OfficePoint BottomRight { get; set; } = new OfficePoint(1, 1);
    /// <summary>Bottom-left corner.</summary>
    public OfficePoint BottomLeft { get; set; } = new OfficePoint(0, 1);
    /// <summary>Maximum source and output pixels.</summary>
    public long MaximumPixels { get; set; } = 20_000_000;
    /// <summary>Maximum accounted source and output pixel-buffer bytes.</summary>
    public long MaximumWorkingBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Creates an independent settings snapshot.</summary>
    public OfficeScanPerspectiveOptions Clone() => (OfficeScanPerspectiveOptions)MemberwiseClone();
    /// <summary>Rejects crossed, concave, reversed, degenerate, or out-of-image corners and invalid budgets.</summary>
    public void Validate() {
        var corners = new[] { TopLeft, TopRight, BottomRight, BottomLeft };
        foreach (var p in corners) {
            if (double.IsNaN(p.X) || double.IsNaN(p.Y) || p.X < 0 || p.X > 1 || p.Y < 0 || p.Y > 1)
                throw new ArgumentException("Perspective corners must be normalized points inside the source image.");
        }
        for (int i = 0; i < 4; i++) {
            OfficePoint a = corners[i], b = corners[(i + 1) % 4], c = corners[(i + 2) % 4];
            if ((b.X - a.X) * (c.Y - b.Y) - (b.Y - a.Y) * (c.X - b.X) < 0.000001D)
                throw new ArgumentException("Perspective corners must form a non-degenerate convex page in clockwise order.");
        }
        if (MaximumPixels <= 0 || MaximumPixels > OfficeRasterGuards.MaximumPixels) throw new ArgumentOutOfRangeException(nameof(MaximumPixels));
        if (MaximumWorkingBytes <= 0 || MaximumWorkingBytes > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentOutOfRangeException(nameof(MaximumWorkingBytes));
    }
}