using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssLinearGradientDefinition {
    private readonly double _angle;
    private readonly HtmlCssGradientStops _stops;
    private readonly bool _repeating;
    private readonly bool _explicitAngle;

    internal HtmlCssLinearGradientDefinition(double angle, HtmlCssGradientStops stops, bool repeating = false, bool explicitAngle = false) {
        _angle = angle;
        _stops = stops;
        _repeating = repeating;
        _explicitAngle = explicitAngle;
    }

    internal bool TryResolve(double width, double height, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, out OfficeLinearGradient? gradient) {
        return TryResolve(width, height, fontSize, rootFontSize, viewportWidth, viewportHeight, out gradient, out _);
    }

    internal bool TryResolve(double width, double height, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, out OfficeLinearGradient? gradient, out bool stopLimitExceeded) {
        return TryResolve(width, height, fontSize, rootFontSize, viewportWidth, viewportHeight, double.NaN, double.NaN, out gradient, out stopLimitExceeded);
    }

    internal bool TryResolve(double width, double height, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, double containerWidth, double containerHeight, out OfficeLinearGradient? gradient, out bool stopLimitExceeded) {
        gradient = null;
        stopLimitExceeded = false;
        if (width <= 0D || height <= 0D) return false;
        OfficeLinearGradient geometry = ResolveGeometry(width, height);
        double dx = (geometry.EndX - geometry.StartX) * width;
        double dy = (geometry.EndY - geometry.StartY) * height;
        double lineLength = Math.Sqrt((dx * dx) + (dy * dy));
        if (!_stops.TryResolve(lineLength, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, _repeating, out IReadOnlyList<OfficeGradientStop>? stops, out stopLimitExceeded) || stops == null) return false;
        gradient = _explicitAngle
            ? OfficeLinearGradient.CreateImported(geometry.StartX, geometry.StartY, geometry.EndX, geometry.EndY, stops)
            : new OfficeLinearGradient(geometry.StartX, geometry.StartY, geometry.EndX, geometry.EndY, stops);
        return true;
    }

    private OfficeLinearGradient ResolveGeometry(double width, double height) {
        if (!_explicitAngle) {
            return OfficeLinearGradient.FromAngle(OfficeColor.Black, OfficeColor.White, _angle);
        }

        double radians = _angle * Math.PI / 180D;
        double directionX = Math.Cos(radians);
        double directionY = Math.Sin(radians);
        double halfLineLength = 0.5D * ((Math.Abs(directionX) * width) + (Math.Abs(directionY) * height));
        double centerX = width * 0.5D;
        double centerY = height * 0.5D;
        return OfficeLinearGradient.CreateImported(
            (centerX - (directionX * halfLineLength)) / width,
            (centerY - (directionY * halfLineLength)) / height,
            (centerX + (directionX * halfLineLength)) / width,
            (centerY + (directionY * halfLineLength)) / height,
            new[] {
                new OfficeGradientStop(0D, OfficeColor.Black),
                new OfficeGradientStop(1D, OfficeColor.White)
            });
    }
}
