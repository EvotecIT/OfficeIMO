using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssRadialGradientDefinition {
    private const double MinimumRadius = 0.000001D;
    private static readonly double CornerScale = Math.Sqrt(2D);

    internal HtmlCssRadialGradientDefinition(
        HtmlCssRadialGradientShape shape,
        HtmlCssRadialGradientSize size,
        string centerX,
        string centerY,
        string? radiusX,
        string? radiusY,
        IReadOnlyList<OfficeGradientStop> stops) {
        Shape = shape;
        Size = size;
        CenterX = centerX;
        CenterY = centerY;
        RadiusX = radiusX;
        RadiusY = radiusY;
        Stops = stops;
    }

    private HtmlCssRadialGradientShape Shape { get; }
    private HtmlCssRadialGradientSize Size { get; }
    private string CenterX { get; }
    private string CenterY { get; }
    private string? RadiusX { get; }
    private string? RadiusY { get; }
    private IReadOnlyList<OfficeGradientStop> Stops { get; }

    internal bool TryResolve(
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        out OfficeRadialGradient? gradient) {
        gradient = null;
        if (width <= 0D || height <= 0D
            || !HtmlRenderCssValues.TryLength(CenterX, width, fontSize, rootFontSize, out double centerXPixels)
            || !HtmlRenderCssValues.TryLength(CenterY, height, fontSize, rootFontSize, out double centerYPixels)) {
            return false;
        }

        double radiusXPixels;
        double radiusYPixels;
        if (Size == HtmlCssRadialGradientSize.Explicit) {
            if (string.IsNullOrWhiteSpace(RadiusX)
                || !HtmlRenderCssValues.TryLength(RadiusX, width, fontSize, rootFontSize, out radiusXPixels)
                || radiusXPixels < 0D) {
                return false;
            }

            if (Shape == HtmlCssRadialGradientShape.Circle) {
                radiusYPixels = radiusXPixels;
            } else if (string.IsNullOrWhiteSpace(RadiusY)
                || !HtmlRenderCssValues.TryLength(RadiusY, height, fontSize, rootFontSize, out radiusYPixels)
                || radiusYPixels < 0D) {
                return false;
            }
        } else {
            ResolveExtent(width, height, centerXPixels, centerYPixels, out radiusXPixels, out radiusYPixels);
        }

        double centerX = centerXPixels / width;
        double centerY = centerYPixels / height;
        double radiusX = Math.Max(MinimumRadius, radiusXPixels) / width;
        double radiusY = Math.Max(MinimumRadius, radiusYPixels) / height;
        gradient = new OfficeRadialGradient(centerX, centerY, 0D, 0D, centerX, centerY, radiusX, radiusY, Stops);
        return true;
    }

    private void ResolveExtent(
        double width,
        double height,
        double centerX,
        double centerY,
        out double radiusX,
        out double radiusY) {
        double left = Math.Abs(centerX);
        double right = Math.Abs(width - centerX);
        double top = Math.Abs(centerY);
        double bottom = Math.Abs(height - centerY);
        bool closest = Size == HtmlCssRadialGradientSize.ClosestSide || Size == HtmlCssRadialGradientSize.ClosestCorner;
        bool corner = Size == HtmlCssRadialGradientSize.ClosestCorner || Size == HtmlCssRadialGradientSize.FarthestCorner;
        double horizontal = closest ? Math.Min(left, right) : Math.Max(left, right);
        double vertical = closest ? Math.Min(top, bottom) : Math.Max(top, bottom);
        if (Shape == HtmlCssRadialGradientShape.Circle) {
            double circleRadius = corner
                ? Math.Sqrt((horizontal * horizontal) + (vertical * vertical))
                : closest ? Math.Min(horizontal, vertical) : Math.Max(horizontal, vertical);
            radiusX = circleRadius;
            radiusY = circleRadius;
            return;
        }

        radiusX = corner ? horizontal * CornerScale : horizontal;
        radiusY = corner ? vertical * CornerScale : vertical;
    }
}

internal enum HtmlCssRadialGradientShape {
    Circle,
    Ellipse
}

internal enum HtmlCssRadialGradientSize {
    ClosestSide,
    ClosestCorner,
    FarthestSide,
    FarthestCorner,
    Explicit
}
