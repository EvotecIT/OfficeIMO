using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssLinearGradientDefinition {
    private readonly double _angle;
    private readonly HtmlCssGradientStops _stops;
    private readonly bool _repeating;

    internal HtmlCssLinearGradientDefinition(double angle, HtmlCssGradientStops stops, bool repeating = false) {
        _angle = angle;
        _stops = stops;
        _repeating = repeating;
    }

    internal bool TryResolve(double width, double height, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, out OfficeLinearGradient? gradient) {
        return TryResolve(width, height, fontSize, rootFontSize, viewportWidth, viewportHeight, out gradient, out _);
    }

    internal bool TryResolve(double width, double height, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, out OfficeLinearGradient? gradient, out bool stopLimitExceeded) {
        gradient = null;
        stopLimitExceeded = false;
        if (width <= 0D || height <= 0D) return false;
        OfficeLinearGradient geometry = OfficeLinearGradient.FromAngle(OfficeColor.Black, OfficeColor.White, _angle);
        double dx = (geometry.EndX - geometry.StartX) * width;
        double dy = (geometry.EndY - geometry.StartY) * height;
        double lineLength = Math.Sqrt((dx * dx) + (dy * dy));
        if (!_stops.TryResolve(lineLength, fontSize, rootFontSize, viewportWidth, viewportHeight, _repeating, out IReadOnlyList<OfficeGradientStop>? stops, out stopLimitExceeded) || stops == null) return false;
        gradient = new OfficeLinearGradient(geometry.StartX, geometry.StartY, geometry.EndX, geometry.EndY, stops);
        return true;
    }
}
