using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>
/// Keeps drawing-unit geometry separate from physical text and stroke metrics during image export.
/// Native page values and inherited caches are never changed by this projection.
/// </summary>
internal readonly struct VisioRenderProjection {
    private readonly OfficeTransform _pageTransform;

    private VisioRenderProjection(VisioPage page, VisioPage surface, double pixelsPerInch, int supersampling) {
        DrawingToPhysical = page.GetEffectivePageScale().ToInches() / page.GetEffectiveDrawingScale().ToInches();
        PhysicalDensity = pixelsPerInch * supersampling;
        GeometryDensity = PhysicalDensity * DrawingToPhysical;
        WidthInches = Math.Max(page.Width * DrawingToPhysical, 0.01D);
        HeightInches = Math.Max(page.Height * DrawingToPhysical, 0.01D);
        double surfaceScale = surface.GetEffectivePageScale().ToInches() / surface.GetEffectiveDrawingScale().ToInches();
        ContentOffsetY = (surface.Height * surfaceScale - page.Height * DrawingToPhysical) * PhysicalDensity;
        _pageTransform = OfficeTransform.Scale(1D, -1D)
            .Then(OfficeTransform.Translate(0D, page.Height));
    }

    internal double DrawingToPhysical { get; }
    internal double GeometryDensity { get; }
    internal double PhysicalDensity { get; }
    internal double WidthInches { get; }
    internal double HeightInches { get; }
    internal double ContentOffsetY { get; }

    internal static VisioRenderProjection Create(VisioPage page, double pixelsPerInch = 1D, int supersampling = 1) =>
        new(page, page, pixelsPerInch, supersampling);

    // Pages share the physical lower-left origin. A background keeps its own drawing scale and size.
    internal static VisioRenderProjection CreateForContent(VisioPage page, VisioPage surface,
        double pixelsPerInch, int supersampling = 1) => new(page, surface, pixelsPerInch, supersampling);

    internal (double X, double Y) PagePoint(double x, double y) {
        OfficePoint point = _pageTransform.TransformPoint(new OfficePoint(x, y));
        return (point.X * GeometryDensity, point.Y * GeometryDensity + ContentOffsetY);
    }
}
