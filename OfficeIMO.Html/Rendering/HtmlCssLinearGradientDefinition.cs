using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssLinearGradientDefinition {
    private readonly double _angle;
    private readonly HtmlCssGradientStops _stops;

    internal HtmlCssLinearGradientDefinition(double angle, HtmlCssGradientStops stops) {
        _angle = angle;
        _stops = stops;
    }

    internal bool TryResolve(double width, double height, double fontSize, double rootFontSize, out OfficeLinearGradient? gradient) {
        gradient = null;
        if (width <= 0D || height <= 0D) return false;
        OfficeLinearGradient geometry = OfficeLinearGradient.FromAngle(OfficeColor.Black, OfficeColor.White, _angle);
        double dx = (geometry.EndX - geometry.StartX) * width;
        double dy = (geometry.EndY - geometry.StartY) * height;
        double lineLength = Math.Sqrt((dx * dx) + (dy * dy));
        if (!_stops.TryResolve(lineLength, fontSize, rootFontSize, out IReadOnlyList<OfficeGradientStop>? stops) || stops == null) return false;
        gradient = new OfficeLinearGradient(geometry.StartX, geometry.StartY, geometry.EndX, geometry.EndY, stops);
        return true;
    }
}
