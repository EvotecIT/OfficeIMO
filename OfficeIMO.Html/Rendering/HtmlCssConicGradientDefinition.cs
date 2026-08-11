using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssConicGradientDefinition {
    private readonly double _angle;
    private readonly string _centerX;
    private readonly string _centerY;
    private readonly HtmlCssGradientStops _stops;
    private readonly bool _repeating;

    internal HtmlCssConicGradientDefinition(
        double angle,
        string centerX,
        string centerY,
        HtmlCssGradientStops stops,
        bool repeating) {
        _angle = angle;
        _centerX = centerX;
        _centerY = centerY;
        _stops = stops;
        _repeating = repeating;
    }

    internal bool TryResolve(
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        out OfficeConicGradient? gradient,
        out bool stopLimitExceeded) {
        gradient = null;
        stopLimitExceeded = false;
        if (width <= 0D || height <= 0D
            || !HtmlRenderCssValues.TryLength(_centerX, width, fontSize, rootFontSize, viewportWidth, viewportHeight, out double centerX)
            || !HtmlRenderCssValues.TryLength(_centerY, height, fontSize, rootFontSize, viewportWidth, viewportHeight, out double centerY)
            || double.IsNaN(centerX) || double.IsInfinity(centerX)
            || double.IsNaN(centerY) || double.IsInfinity(centerY)
            || !_stops.TryResolveConic(_repeating, out IReadOnlyList<OfficeGradientStop>? stops, out stopLimitExceeded)
            || stops == null) return false;
        gradient = new OfficeConicGradient(centerX / width, centerY / height, _angle, stops);
        return true;
    }
}
