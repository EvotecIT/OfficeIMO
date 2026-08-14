using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssClipPathParser {
    private const double CircleBezierKappa = 0.5522847498307936D;

    internal static bool IsSupportedSyntax(string? value) =>
        TryResolve(value, 100D, 100D, 16D, 16D, 100D, 100D, double.NaN, double.NaN, out _, out _);

    internal static bool TryResolve(
        string? value,
        double boxWidth,
        double boxHeight,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        out HtmlCssResolvedClipPath? resolved,
        out string detail) {
        resolved = null;
        detail = string.Empty;
        string normalized = string.IsNullOrWhiteSpace(value) ? "none" : value!.Trim().ToLowerInvariant();
        if (normalized == "none") return true;

        if (!TryGetFunction(normalized, out string function, out string arguments)) {
            detail = "clip-path=" + normalized;
            return false;
        }

        bool success = function switch {
            "inset" => TryInset(arguments, boxWidth, boxHeight, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out resolved),
            "circle" => TryCircle(arguments, boxWidth, boxHeight, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out resolved),
            "ellipse" => TryEllipse(arguments, boxWidth, boxHeight, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out resolved),
            "polygon" => TryPolygon(arguments, boxWidth, boxHeight, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out resolved),
            _ => false
        };
        if (!success) detail = "clip-path=" + normalized;
        return success;
    }

    private static bool TryInset(
        string arguments,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        out HtmlCssResolvedClipPath? resolved) {
        resolved = null;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(arguments);
        int roundIndex = tokens.ToList().FindIndex(token => string.Equals(token, "round", StringComparison.OrdinalIgnoreCase));
        IReadOnlyList<string> insetTokens = roundIndex < 0 ? tokens : tokens.Take(roundIndex).ToList();
        if (insetTokens.Count < 1 || insetTokens.Count > 4 || roundIndex == tokens.Count - 1) return false;

        var values = new double[insetTokens.Count];
        for (int index = 0; index < insetTokens.Count; index++) {
            double reference = index % 2 == 0 ? height : width;
            if (!TryLength(insetTokens[index], reference, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out values[index])) return false;
        }
        ExpandFour(values, out double top, out double right, out double bottom, out double left);
        double clipWidth = width - left - right;
        double clipHeight = height - top - bottom;
        if (clipWidth <= 0.0001D || clipHeight <= 0.0001D) return false;

        OfficeClipPath path;
        if (roundIndex < 0) {
            path = OfficeClipPath.Rectangle(clipWidth, clipHeight);
        } else {
            var style = new HtmlRenderBoxStyle {
                BorderRadius = string.Join(" ", tokens.Skip(roundIndex + 1)),
                Font = new OfficeFontInfo("Arial", fontSize)
            };
            if (!HtmlCssBorderRadiusParser.TryResolve(
                    style, clipWidth, clipHeight, rootFontSize, viewportWidth, viewportHeight,
                    containerWidth, containerHeight, out HtmlResolvedBorderRadii radii, out _)) return false;
            path = radii.IsZero
                ? OfficeClipPath.Rectangle(clipWidth, clipHeight)
                : radii.IsUniformCircular
                    ? OfficeClipPath.RoundedRectangle(clipWidth, clipHeight, radii.UniformRadius)
                    : OfficeClipPath.Path(radii.CreatePathCommands(clipWidth, clipHeight));
        }

        resolved = new HtmlCssResolvedClipPath(left, top, path);
        return true;
    }

    private static bool TryCircle(
        string arguments,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        out HtmlCssResolvedClipPath? resolved) {
        resolved = null;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(arguments);
        SplitAtPosition(tokens, out IReadOnlyList<string> shapeTokens, out IReadOnlyList<string> positionTokens);
        if (shapeTokens.Count > 1 || !TryPosition(positionTokens, width, height, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out double centerX, out double centerY)) return false;

        double radius;
        string radiusToken = shapeTokens.Count == 0 ? "closest-side" : shapeTokens[0];
        if (radiusToken == "closest-side") radius = Math.Min(Math.Min(centerX, width - centerX), Math.Min(centerY, height - centerY));
        else if (radiusToken == "farthest-side") radius = Math.Max(Math.Max(centerX, width - centerX), Math.Max(centerY, height - centerY));
        else {
            double reference = Math.Sqrt(width * width + height * height) / Math.Sqrt(2D);
            if (!TryLength(radiusToken, reference, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out radius)) return false;
        }
        if (radius <= 0.0001D) return false;
        resolved = new HtmlCssResolvedClipPath(centerX - radius, centerY - radius, CreateEllipse(radius, radius));
        return true;
    }

    private static bool TryEllipse(
        string arguments,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        out HtmlCssResolvedClipPath? resolved) {
        resolved = null;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(arguments);
        SplitAtPosition(tokens, out IReadOnlyList<string> shapeTokens, out IReadOnlyList<string> positionTokens);
        if (shapeTokens.Count != 0 && shapeTokens.Count != 2
            || !TryPosition(positionTokens, width, height, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out double centerX, out double centerY)) return false;

        string horizontal = shapeTokens.Count == 0 ? "closest-side" : shapeTokens[0];
        string vertical = shapeTokens.Count == 0 ? "closest-side" : shapeTokens[1];
        if (!TryShapeRadius(horizontal, centerX, width - centerX, width, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out double radiusX)
            || !TryShapeRadius(vertical, centerY, height - centerY, height, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out double radiusY)
            || radiusX <= 0.0001D || radiusY <= 0.0001D) return false;
        resolved = new HtmlCssResolvedClipPath(centerX - radiusX, centerY - radiusY, CreateEllipse(radiusX, radiusY));
        return true;
    }

    private static bool TryPolygon(
        string arguments,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        out HtmlCssResolvedClipPath? resolved) {
        resolved = null;
        IReadOnlyList<string> entries = HtmlRenderCssValues.SplitTopLevelCommas(arguments);
        if (entries.Count < 3) return false;
        OfficeFillRule fillRule = OfficeFillRule.NonZero;
        int start = 0;
        if (entries[0].Equals("evenodd", StringComparison.OrdinalIgnoreCase)) { fillRule = OfficeFillRule.EvenOdd; start = 1; }
        else if (entries[0].Equals("nonzero", StringComparison.OrdinalIgnoreCase)) start = 1;
        if (entries.Count - start < 3) return false;

        var points = new List<OfficePoint>(entries.Count - start);
        double minX = double.MaxValue;
        double minY = double.MaxValue;
        for (int index = start; index < entries.Count; index++) {
            IReadOnlyList<string> coordinates = HtmlRenderCssValues.SplitWhitespace(entries[index]);
            if (coordinates.Count != 2
                || !TryLength(coordinates[0], width, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out double x)
                || !TryLength(coordinates[1], height, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out double y)) return false;
            points.Add(new OfficePoint(x, y));
            minX = Math.Min(minX, x);
            minY = Math.Min(minY, y);
        }

        var commands = new List<OfficePathCommand>(points.Count + 1) { OfficePathCommand.MoveTo(points[0]) };
        for (int index = 1; index < points.Count; index++) commands.Add(OfficePathCommand.LineTo(points[index]));
        commands.Add(OfficePathCommand.Close());
        try {
            resolved = new HtmlCssResolvedClipPath(minX, minY, OfficeClipPath.Path(commands, fillRule));
            return true;
        } catch (ArgumentException) {
            return false;
        }
    }

    private static OfficeClipPath CreateEllipse(double radiusX, double radiusY) {
        double width = radiusX * 2D;
        double height = radiusY * 2D;
        double kx = radiusX * CircleBezierKappa;
        double ky = radiusY * CircleBezierKappa;
        return OfficeClipPath.Path(new[] {
            OfficePathCommand.MoveTo(radiusX, 0D),
            OfficePathCommand.CubicBezierTo(radiusX + kx, 0D, width, radiusY - ky, width, radiusY),
            OfficePathCommand.CubicBezierTo(width, radiusY + ky, radiusX + kx, height, radiusX, height),
            OfficePathCommand.CubicBezierTo(radiusX - kx, height, 0D, radiusY + ky, 0D, radiusY),
            OfficePathCommand.CubicBezierTo(0D, radiusY - ky, radiusX - kx, 0D, radiusX, 0D),
            OfficePathCommand.Close()
        }, OfficeFillRule.NonZero);
    }

    private static bool TryPosition(
        IReadOnlyList<string> tokens,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        out double x,
        out double y) {
        x = y = 0D;
        if (!HtmlCssGradientPositionParser.TryParse(tokens, out string xValue, out string yValue)) return false;
        return TryLength(xValue, width, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out x)
            && TryLength(yValue, height, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out y);
    }

    private static void SplitAtPosition(IReadOnlyList<string> tokens, out IReadOnlyList<string> shape, out IReadOnlyList<string> position) {
        int at = tokens.ToList().FindIndex(token => string.Equals(token, "at", StringComparison.OrdinalIgnoreCase));
        shape = at < 0 ? tokens : tokens.Take(at).ToList();
        position = at < 0 ? Array.Empty<string>() : tokens.Skip(at + 1).ToList();
    }

    private static bool TryShapeRadius(string token, double near, double far, double reference, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, double containerWidth, double containerHeight, out double radius) {
        if (token == "closest-side") { radius = Math.Min(near, far); return true; }
        if (token == "farthest-side") { radius = Math.Max(near, far); return true; }
        return TryLength(token, reference, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out radius);
    }

    private static bool TryLength(string value, double reference, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, double containerWidth, double containerHeight, out double length) =>
        HtmlRenderCssValues.TryLength(value, reference, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out length);

    private static bool TryGetFunction(string value, out string name, out string arguments) {
        name = string.Empty;
        arguments = string.Empty;
        int open = value.IndexOf('(');
        if (open <= 0 || HtmlRenderCssValues.FindMatchingParenthesis(value, open) != value.Length - 1) return false;
        name = value.Substring(0, open).Trim();
        arguments = value.Substring(open + 1, value.Length - open - 2).Trim();
        return arguments.Length > 0;
    }

    private static void ExpandFour(double[] values, out double top, out double right, out double bottom, out double left) {
        top = values[0];
        right = values.Length > 1 ? values[1] : top;
        bottom = values.Length > 2 ? values[2] : top;
        left = values.Length > 3 ? values[3] : right;
    }
}

internal sealed class HtmlCssResolvedClipPath {
    internal HtmlCssResolvedClipPath(double x, double y, OfficeClipPath clipPath) {
        X = x;
        Y = y;
        ClipPath = clipPath;
    }

    internal double X { get; }
    internal double Y { get; }
    internal OfficeClipPath ClipPath { get; }
}
