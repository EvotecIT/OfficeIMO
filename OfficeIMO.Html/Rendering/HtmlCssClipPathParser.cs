using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssClipPathParser {
    private const double CircleBezierKappa = 0.5522847498307936D;

    internal static bool IsSupportedSyntax(string? value) =>
        TryResolve(value, 100D, 100D, 16D, 16D, 100D, 100D, double.NaN, double.NaN, null, out _, out _);

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
        HtmlRenderBoxStyle? style,
        out HtmlCssResolvedClipPath? resolved,
        out string detail) {
        resolved = null;
        detail = string.Empty;
        string normalized = string.IsNullOrWhiteSpace(value) ? "none" : value!.Trim().ToLowerInvariant();
        if (normalized == "none") return true;

        if (!TryGetClipPathParts(normalized, out string function, out string arguments, out string geometryBox)) {
            detail = "clip-path=" + normalized;
            return false;
        }

        ResolveReferenceBox(geometryBox, boxWidth, boxHeight, style, out double referenceX, out double referenceY, out double referenceWidth, out double referenceHeight);
        if (referenceWidth <= 0.0001D || referenceHeight <= 0.0001D) {
            resolved = HtmlCssResolvedClipPath.Empty;
            return true;
        }
        if (function.Length == 0) {
            resolved = new HtmlCssResolvedClipPath(referenceX, referenceY, OfficeClipPath.Rectangle(referenceWidth, referenceHeight));
            return true;
        }

        bool success = function switch {
            "inset" => TryInset(arguments, referenceWidth, referenceHeight, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out resolved),
            "circle" => TryCircle(arguments, referenceWidth, referenceHeight, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out resolved),
            "ellipse" => TryEllipse(arguments, referenceWidth, referenceHeight, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out resolved),
            "polygon" => TryPolygon(arguments, referenceWidth, referenceHeight, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out resolved),
            _ => false
        };
        if (success && resolved != null && resolved.ClipPath.Kind != OfficeClipPathKind.Empty && (referenceX != 0D || referenceY != 0D)) {
            resolved = new HtmlCssResolvedClipPath(referenceX + resolved.X, referenceY + resolved.Y, resolved.ClipPath);
        }
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
        if (clipWidth <= 0.0001D || clipHeight <= 0.0001D) {
            resolved = HtmlCssResolvedClipPath.Empty;
            return true;
        }

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
        bool hasPosition = SplitAtPosition(tokens, out IReadOnlyList<string> shapeTokens, out IReadOnlyList<string> positionTokens);
        if (shapeTokens.Count > 1
            || hasPosition && positionTokens.Count == 0
            || !TryPosition(positionTokens, width, height, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out double centerX, out double centerY)) return false;

        double radius;
        string radiusToken = shapeTokens.Count == 0 ? "closest-side" : shapeTokens[0];
        if (radiusToken == "closest-side") radius = Math.Min(Math.Min(Math.Abs(centerX), Math.Abs(width - centerX)), Math.Min(Math.Abs(centerY), Math.Abs(height - centerY)));
        else if (radiusToken == "farthest-side") radius = Math.Max(Math.Max(Math.Abs(centerX), Math.Abs(width - centerX)), Math.Max(Math.Abs(centerY), Math.Abs(height - centerY)));
        else if (radiusToken == "closest-corner") radius = ResolveCircleCornerRadius(centerX, centerY, width, height, closest: true);
        else if (radiusToken == "farthest-corner") radius = ResolveCircleCornerRadius(centerX, centerY, width, height, closest: false);
        else {
            double reference = Math.Sqrt(width * width + height * height) / Math.Sqrt(2D);
            if (!TryLength(radiusToken, reference, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out radius)) return false;
            if (radius < 0D) return false;
        }
        if (radius <= 0.0001D) {
            resolved = HtmlCssResolvedClipPath.Empty;
            return true;
        }
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
        bool hasPosition = SplitAtPosition(tokens, out IReadOnlyList<string> shapeTokens, out IReadOnlyList<string> positionTokens);
        bool extentKeyword = shapeTokens.Count == 1 && IsRadialExtent(shapeTokens[0]);
        if (shapeTokens.Count != 0 && !extentKeyword && shapeTokens.Count != 2
            || shapeTokens.Count == 2 && (IsRadialExtent(shapeTokens[0]) || IsRadialExtent(shapeTokens[1]))
            || hasPosition && positionTokens.Count == 0
            || !TryPosition(positionTokens, width, height, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out double centerX, out double centerY)) return false;

        double radiusX;
        double radiusY;
        if (shapeTokens.Count == 0 || extentKeyword) {
            string extent = shapeTokens.Count == 0 ? "closest-side" : shapeTokens[0];
            ResolveEllipseExtent(extent, centerX, centerY, width, height, out radiusX, out radiusY);
        } else if (!TryLength(shapeTokens[0], width, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out radiusX)
            || !TryLength(shapeTokens[1], height, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out radiusY)
            || radiusX < 0D || radiusY < 0D) {
            return false;
        }
        if (radiusX <= 0.0001D || radiusY <= 0.0001D) {
            resolved = HtmlCssResolvedClipPath.Empty;
            return true;
        }
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

        if (AreCollinear(points)) {
            resolved = HtmlCssResolvedClipPath.Empty;
            return true;
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

    private static bool AreCollinear(IReadOnlyList<OfficePoint> points) {
        const double tolerance = 0.0001D;
        OfficePoint origin = points[0];
        int distinctIndex = 1;
        while (distinctIndex < points.Count
               && Math.Abs(points[distinctIndex].X - origin.X) <= tolerance
               && Math.Abs(points[distinctIndex].Y - origin.Y) <= tolerance) {
            distinctIndex++;
        }
        if (distinctIndex == points.Count) return true;

        double directionX = points[distinctIndex].X - origin.X;
        double directionY = points[distinctIndex].Y - origin.Y;
        double scale = Math.Max(1D, Math.Sqrt(directionX * directionX + directionY * directionY));
        for (int index = distinctIndex + 1; index < points.Count; index++) {
            double offsetX = points[index].X - origin.X;
            double offsetY = points[index].Y - origin.Y;
            if (Math.Abs(directionX * offsetY - directionY * offsetX) > tolerance * scale) return false;
        }
        return true;
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

    private static bool SplitAtPosition(IReadOnlyList<string> tokens, out IReadOnlyList<string> shape, out IReadOnlyList<string> position) {
        int at = tokens.ToList().FindIndex(token => string.Equals(token, "at", StringComparison.OrdinalIgnoreCase));
        shape = at < 0 ? tokens : tokens.Take(at).ToList();
        position = at < 0 ? Array.Empty<string>() : tokens.Skip(at + 1).ToList();
        return at >= 0;
    }

    private static double ResolveCircleCornerRadius(double centerX, double centerY, double width, double height, bool closest) {
        double[] distances = {
            Math.Sqrt(centerX * centerX + centerY * centerY),
            Math.Sqrt((width - centerX) * (width - centerX) + centerY * centerY),
            Math.Sqrt(centerX * centerX + (height - centerY) * (height - centerY)),
            Math.Sqrt((width - centerX) * (width - centerX) + (height - centerY) * (height - centerY))
        };
        return closest ? distances.Min() : distances.Max();
    }

    private static bool IsRadialExtent(string token) =>
        token == "closest-side"
        || token == "farthest-side"
        || token == "closest-corner"
        || token == "farthest-corner";

    private static void ResolveEllipseExtent(
        string extent,
        double centerX,
        double centerY,
        double width,
        double height,
        out double radiusX,
        out double radiusY) {
        bool closest = extent == "closest-side" || extent == "closest-corner";
        radiusX = closest ? Math.Min(Math.Abs(centerX), Math.Abs(width - centerX)) : Math.Max(Math.Abs(centerX), Math.Abs(width - centerX));
        radiusY = closest ? Math.Min(Math.Abs(centerY), Math.Abs(height - centerY)) : Math.Max(Math.Abs(centerY), Math.Abs(height - centerY));
        if (extent == "closest-side" || extent == "farthest-side" || radiusX <= 0D || radiusY <= 0D) return;

        double[] scales = {
            Math.Sqrt(centerX * centerX / (radiusX * radiusX) + centerY * centerY / (radiusY * radiusY)),
            Math.Sqrt((width - centerX) * (width - centerX) / (radiusX * radiusX) + centerY * centerY / (radiusY * radiusY)),
            Math.Sqrt(centerX * centerX / (radiusX * radiusX) + (height - centerY) * (height - centerY) / (radiusY * radiusY)),
            Math.Sqrt((width - centerX) * (width - centerX) / (radiusX * radiusX) + (height - centerY) * (height - centerY) / (radiusY * radiusY))
        };
        double scale = closest ? scales.Min() : scales.Max();
        radiusX *= scale;
        radiusY *= scale;
    }

    private static bool TryLength(string value, double reference, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, double containerWidth, double containerHeight, out double length) =>
        HtmlRenderCssValues.TryLength(value, reference, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out length);

    private static bool TryGetClipPathParts(string value, out string name, out string arguments, out string geometryBox) {
        name = string.Empty;
        arguments = string.Empty;
        geometryBox = "border-box";
        int open = value.IndexOf('(');
        if (open < 0) {
            if (!IsGeometryBox(value)) return false;
            geometryBox = value;
            return true;
        }
        if (open == 0) return false;
        int close = HtmlRenderCssValues.FindMatchingParenthesis(value, open);
        if (close < 0) return false;

        IReadOnlyList<string> prefix = HtmlRenderCssValues.SplitWhitespace(value.Substring(0, open));
        if (prefix.Count < 1 || prefix.Count > 2) return false;
        name = prefix[prefix.Count - 1];
        bool hasPrefixBox = prefix.Count == 2;
        if (hasPrefixBox && !IsGeometryBox(prefix[0])) return false;

        string suffixValue = value.Substring(close + 1).Trim();
        IReadOnlyList<string> suffix = suffixValue.Length == 0
            ? Array.Empty<string>()
            : HtmlRenderCssValues.SplitWhitespace(suffixValue);
        if (suffix.Count > 1 || hasPrefixBox && suffix.Count > 0 || suffix.Count == 1 && !IsGeometryBox(suffix[0])) return false;
        if (hasPrefixBox) geometryBox = prefix[0];
        else if (suffix.Count == 1) geometryBox = suffix[0];

        arguments = value.Substring(open + 1, close - open - 1).Trim();
        return arguments.Length > 0
            || name == "circle"
            || name == "ellipse";
    }

    private static bool IsGeometryBox(string value) =>
        value == "margin-box"
        || value == "border-box"
        || value == "padding-box"
        || value == "content-box";

    private static void ResolveReferenceBox(
        string geometryBox,
        double boxWidth,
        double boxHeight,
        HtmlRenderBoxStyle? style,
        out double x,
        out double y,
        out double width,
        out double height) {
        x = y = 0D;
        width = boxWidth;
        height = boxHeight;
        if (style == null || geometryBox == "border-box") return;
        if (geometryBox == "margin-box") {
            x = -style.MarginLeft;
            y = -style.MarginTop;
            width += style.MarginLeft + style.MarginRight;
            height += style.MarginTop + style.MarginBottom;
            return;
        }

        x = style.BorderLeftWidth;
        y = style.BorderTopWidth;
        width -= style.BorderLeftWidth + style.BorderRightWidth;
        height -= style.BorderTopWidth + style.BorderBottomWidth;
        if (geometryBox == "padding-box") return;

        x += style.PaddingLeft;
        y += style.PaddingTop;
        width -= style.PaddingLeft + style.PaddingRight;
        height -= style.PaddingTop + style.PaddingBottom;
    }

    private static void ExpandFour(double[] values, out double top, out double right, out double bottom, out double left) {
        top = values[0];
        right = values.Length > 1 ? values[1] : top;
        bottom = values.Length > 2 ? values[2] : top;
        left = values.Length > 3 ? values[3] : right;
    }
}

internal sealed class HtmlCssResolvedClipPath {
    internal static HtmlCssResolvedClipPath Empty { get; } = new HtmlCssResolvedClipPath(0D, 0D, OfficeClipPath.Empty());

    internal HtmlCssResolvedClipPath(double x, double y, OfficeClipPath clipPath) {
        X = x;
        Y = y;
        ClipPath = clipPath;
    }

    internal double X { get; }
    internal double Y { get; }
    internal OfficeClipPath ClipPath { get; }
}
