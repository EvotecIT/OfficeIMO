namespace OfficeIMO.Html;

internal static class HtmlCssBorderRadiusParser {
    internal static bool TryResolve(
        HtmlRenderBoxStyle style,
        double width,
        double height,
        double rootFontSize,
        out HtmlResolvedBorderRadii radii,
        out string detail) {
        radii = default;
        detail = string.Empty;
        if (!TryParseShorthand(style.BorderRadius, width, height, style.Font.Size, rootFontSize, out double[] horizontal, out double[] vertical)) {
            detail = "border-radius=" + style.BorderRadius;
            return false;
        }

        string[] overrides = {
            style.BorderTopLeftRadius,
            style.BorderTopRightRadius,
            style.BorderBottomRightRadius,
            style.BorderBottomLeftRadius
        };
        for (int index = 0; index < overrides.Length; index++) {
            if (string.IsNullOrWhiteSpace(overrides[index])) continue;
            if (!TryParseCorner(overrides[index], width, height, style.Font.Size, rootFontSize, out horizontal[index], out vertical[index])) {
                detail = CornerPropertyName(index) + "=" + overrides[index];
                return false;
            }
        }

        radii = new HtmlResolvedBorderRadii(
            horizontal[0], vertical[0],
            horizontal[1], vertical[1],
            horizontal[2], vertical[2],
            horizontal[3], vertical[3]).Normalize(width, height);
        return true;
    }

    internal static bool IsSupportedShorthandSyntax(string value) {
        var style = new HtmlRenderBoxStyle {
            BorderRadius = value,
            Font = new OfficeIMO.Drawing.OfficeFontInfo("Arial", 16D)
        };
        return TryResolve(style, 100D, 100D, 16D, out _, out _);
    }

    internal static bool IsSupportedCornerSyntax(string value) =>
        TryParseCorner(value, 100D, 100D, 16D, 16D, out _, out _);

    private static bool TryParseShorthand(
        string value,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        out double[] horizontal,
        out double[] vertical) {
        horizontal = new double[4];
        vertical = new double[4];
        string normalized = string.IsNullOrWhiteSpace(value) ? "0" : value.Trim().ToLowerInvariant();
        string[] axes = normalized.Split('/');
        if (axes.Length > 2
            || !TryParseAxis(axes[0], width, fontSize, rootFontSize, out horizontal)) return false;
        if (axes.Length == 1) {
            Array.Copy(horizontal, vertical, horizontal.Length);
            return true;
        }
        return TryParseAxis(axes[1], height, fontSize, rootFontSize, out vertical);
    }

    private static bool TryParseCorner(
        string value,
        double width,
        double height,
        double fontSize,
        double rootFontSize,
        out double horizontal,
        out double vertical) {
        horizontal = 0D;
        vertical = 0D;
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(value.Trim().ToLowerInvariant());
        if (values.Count < 1 || values.Count > 2
            || !TryLength(values[0], width, fontSize, rootFontSize, out horizontal)) return false;
        if (values.Count == 1) {
            vertical = horizontal;
            return true;
        }
        return TryLength(values[1], height, fontSize, rootFontSize, out vertical);
    }

    private static bool TryParseAxis(string value, double reference, double fontSize, double rootFontSize, out double[] expanded) {
        expanded = new double[4];
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(value.Trim());
        if (values.Count < 1 || values.Count > 4) return false;
        var resolved = new double[values.Count];
        for (int index = 0; index < values.Count; index++) {
            if (!TryLength(values[index], reference, fontSize, rootFontSize, out resolved[index])) return false;
        }
        expanded[0] = resolved[0];
        expanded[1] = resolved.Length > 1 ? resolved[1] : resolved[0];
        expanded[2] = resolved.Length > 2 ? resolved[2] : resolved[0];
        expanded[3] = resolved.Length > 3 ? resolved[3] : expanded[1];
        return true;
    }

    private static bool TryLength(string value, double reference, double fontSize, double rootFontSize, out double length) =>
        HtmlRenderCssValues.TryLength(value, reference, fontSize, rootFontSize, out length)
        && length >= 0D
        && !double.IsNaN(length)
        && !double.IsInfinity(length);

    private static string CornerPropertyName(int index) => index switch {
        0 => "border-top-left-radius",
        1 => "border-top-right-radius",
        2 => "border-bottom-right-radius",
        _ => "border-bottom-left-radius"
    };
}
