using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssBoxStrokeParser {
    private static readonly string[] SideNames = { "top", "right", "bottom", "left" };
    private static readonly string[] SideProperties = SideNames.SelectMany(side => new[] {
        "border-" + side,
        "border-" + side + "-width",
        "border-" + side + "-style",
        "border-" + side + "-color"
    }).ToArray();

    internal static bool HasBorderDeclaration(HtmlComputedStyle computed) =>
        computed.GetValue("border").Length > 0
        || computed.GetValue("border-width").Length > 0
        || computed.GetValue("border-style").Length > 0
        || computed.GetValue("border-color").Length > 0
        || SideProperties.Any(property => computed.GetValue(property).Length > 0);

    internal static bool TryParseBorder(
        HtmlComputedStyle computed,
        double reference,
        double fontSize,
        double rootFontSize,
        OfficeColor currentColor,
        out HtmlRenderBorderEdges borders,
        out string detail) {
        string shorthand = computed.GetValue("border").Trim();
        string widthValue = computed.GetValue("border-width").Trim();
        string styleValue = computed.GetValue("border-style").Trim();
        string colorValue = computed.GetValue("border-color").Trim();
        var sides = Enumerable.Range(0, 4)
            .Select(_ => new HtmlRenderBorderSide(3D, "none", currentColor))
            .ToArray();
        borders = HtmlRenderBorderEdges.Uniform(0D, "none", currentColor);
        detail = string.Empty;
        bool declared = shorthand.Length > 0 || widthValue.Length > 0 || styleValue.Length > 0 || colorValue.Length > 0
            || SideProperties.Any(property => computed.GetValue(property).Length > 0);
        if (!declared) {
            return true;
        }
        if (shorthand.Length > 0) {
            double width = 3D;
            string style = "none";
            OfficeColor color = currentColor;
            if (!TryParseStrokeShorthand(shorthand, reference, fontSize, rootFontSize, currentColor, ref width, ref style, ref color)) {
                detail = "border=" + shorthand;
                return false;
            }
            for (int index = 0; index < sides.Length; index++) sides[index] = new HtmlRenderBorderSide(width, style, color);
        }
        if (widthValue.Length > 0) {
            if (!TryParseWidths(widthValue, reference, fontSize, rootFontSize, out double[] widths)) {
                detail = "border-width=" + widthValue;
                return false;
            }
            for (int index = 0; index < sides.Length; index++) sides[index] = sides[index].WithWidth(widths[index]);
        }
        if (styleValue.Length > 0) {
            if (!TryParseStyles(styleValue, out string[] styles)) {
                detail = "border-style=" + styleValue;
                return false;
            }
            for (int index = 0; index < sides.Length; index++) sides[index] = sides[index].WithStyle(styles[index]);
        }
        if (colorValue.Length > 0) {
            if (!TryParseColors(colorValue, currentColor, out OfficeColor[] colors)) {
                detail = "border-color=" + colorValue;
                return false;
            }
            for (int index = 0; index < sides.Length; index++) sides[index] = sides[index].WithColor(colors[index]);
        }

        for (int index = 0; index < SideNames.Length; index++) {
            string prefix = "border-" + SideNames[index];
            string sideShorthand = computed.GetValue(prefix).Trim();
            if (sideShorthand.Length > 0) {
                double width = 3D;
                string style = "none";
                OfficeColor color = currentColor;
                if (!TryParseStrokeShorthand(sideShorthand, reference, fontSize, rootFontSize, currentColor, ref width, ref style, ref color)) {
                    detail = prefix + "=" + sideShorthand;
                    return false;
                }
                sides[index] = new HtmlRenderBorderSide(width, style, color);
            }

            string sideWidth = computed.GetValue(prefix + "-width").Trim();
            if (sideWidth.Length > 0) {
                if (!TryStrokeWidth(sideWidth, reference, fontSize, rootFontSize, out double width)) {
                    detail = prefix + "-width=" + sideWidth;
                    return false;
                }
                sides[index] = sides[index].WithWidth(width);
            }

            string sideStyle = computed.GetValue(prefix + "-style").Trim();
            if (sideStyle.Length > 0) {
                if (!TryStrokeStyle(sideStyle, out string parsedStyle)) {
                    detail = prefix + "-style=" + sideStyle;
                    return false;
                }
                sides[index] = sides[index].WithStyle(parsedStyle);
            }

            string sideColor = computed.GetValue(prefix + "-color").Trim();
            if (sideColor.Length > 0) {
                if (!TryStrokeColor(sideColor, currentColor, out OfficeColor parsedColor)) {
                    detail = prefix + "-color=" + sideColor;
                    return false;
                }
                sides[index] = sides[index].WithColor(parsedColor);
            }
        }

        borders = new HtmlRenderBorderEdges(sides[0], sides[1], sides[2], sides[3]);
        return true;
    }

    internal static bool TryParseOutline(
        HtmlComputedStyle computed,
        double reference,
        double fontSize,
        double rootFontSize,
        OfficeColor currentColor,
        out double width,
        out string style,
        out OfficeColor color,
        out double offset,
        out string detail) {
        string shorthand = computed.GetValue("outline").Trim();
        string widthValue = computed.GetValue("outline-width").Trim();
        string styleValue = computed.GetValue("outline-style").Trim();
        string colorValue = computed.GetValue("outline-color").Trim();
        string offsetValue = computed.GetValue("outline-offset").Trim();
        width = 3D;
        style = "none";
        color = currentColor;
        offset = 0D;
        detail = string.Empty;
        bool declared = shorthand.Length > 0 || widthValue.Length > 0 || styleValue.Length > 0 || colorValue.Length > 0 || offsetValue.Length > 0;
        if (!declared) {
            width = 0D;
            return true;
        }
        if (shorthand.Length > 0 && !TryParseStrokeShorthand(shorthand, reference, fontSize, rootFontSize, currentColor, ref width, ref style, ref color)) {
            width = 0D;
            detail = "outline=" + shorthand;
            return false;
        }
        if (widthValue.Length > 0 && !TryStrokeWidth(widthValue, reference, fontSize, rootFontSize, out width)) {
            width = 0D;
            detail = "outline-width=" + widthValue;
            return false;
        }
        if (styleValue.Length > 0 && !TryStrokeStyle(styleValue, out style)) {
            width = 0D;
            detail = "outline-style=" + styleValue;
            return false;
        }
        if (colorValue.Length > 0 && !TryStrokeColor(colorValue, currentColor, out color)) {
            width = 0D;
            detail = "outline-color=" + colorValue;
            return false;
        }
        if (offsetValue.Length > 0
            && (offsetValue.EndsWith("%", StringComparison.Ordinal)
                || !HtmlRenderCssValues.TryLength(offsetValue, reference, fontSize, rootFontSize, out offset))) {
            width = 0D;
            detail = "outline-offset=" + offsetValue;
            return false;
        }
        if (style == "none" || style == "hidden") width = 0D;
        return true;
    }

    internal static bool IsSupportedBorderSyntax(string value) {
        double width = 3D;
        string style = "none";
        OfficeColor color = OfficeColor.Black;
        return TryParseStrokeShorthand(value, 100D, 16D, 16D, OfficeColor.Black, ref width, ref style, ref color);
    }

    internal static bool IsSupportedOutlineSyntax(string value) => IsSupportedBorderSyntax(value);
    internal static bool IsSupportedWidthSyntax(string value) => TryParseWidths(value, 100D, 16D, 16D, out _);
    internal static bool IsSupportedStyleSyntax(string value) => TryParseStyles(value, out _);
    internal static bool IsSupportedColorSyntax(string value) => TryParseColors(value, OfficeColor.Black, out _);
    internal static bool IsSupportedSideWidthSyntax(string value) => TryStrokeWidth(value, 100D, 16D, 16D, out _);
    internal static bool IsSupportedSideStyleSyntax(string value) => TryStrokeStyle(value, out _);
    internal static bool IsSupportedSideColorSyntax(string value) => TryStrokeColor(value, OfficeColor.Black, out _);

    private static bool TryParseStrokeShorthand(
        string value,
        double reference,
        double fontSize,
        double rootFontSize,
        OfficeColor currentColor,
        ref double width,
        ref string style,
        ref OfficeColor color) {
        bool widthSet = false;
        bool styleSet = false;
        bool colorSet = false;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value.Trim().ToLowerInvariant());
        if (tokens.Count < 1 || tokens.Count > 3) return false;
        foreach (string token in tokens) {
            if (!widthSet && TryStrokeWidth(token, reference, fontSize, rootFontSize, out double parsedWidth)) {
                width = parsedWidth;
                widthSet = true;
            } else if (!styleSet && TryStrokeStyle(token, out string parsedStyle)) {
                style = parsedStyle;
                styleSet = true;
            } else if (!colorSet && TryStrokeColor(token, currentColor, out OfficeColor parsedColor)) {
                color = parsedColor;
                colorSet = true;
            } else {
                return false;
            }
        }
        return true;
    }

    private static bool TryParseWidths(string value, double reference, double fontSize, double rootFontSize, out double[] widths) {
        widths = new double[4];
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 4) return false;
        var parsed = new double[tokens.Count];
        for (int index = 0; index < tokens.Count; index++)
            if (!TryStrokeWidth(tokens[index], reference, fontSize, rootFontSize, out parsed[index])) return false;
        ExpandFour(parsed, widths);
        return true;
    }

    private static bool TryParseStyles(string value, out string[] styles) {
        styles = new string[4];
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 4) return false;
        var parsed = new string[tokens.Count];
        for (int index = 0; index < tokens.Count; index++)
            if (!TryStrokeStyle(tokens[index], out parsed[index])) return false;
        ExpandFour(parsed, styles);
        return true;
    }

    private static bool TryParseColors(string value, OfficeColor currentColor, out OfficeColor[] colors) {
        colors = new OfficeColor[4];
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 4) return false;
        var parsed = new OfficeColor[tokens.Count];
        for (int index = 0; index < tokens.Count; index++)
            if (!TryStrokeColor(tokens[index], currentColor, out parsed[index])) return false;
        ExpandFour(parsed, colors);
        return true;
    }

    private static void ExpandFour<T>(IReadOnlyList<T> source, T[] target) {
        target[0] = source[0];
        target[1] = source.Count > 1 ? source[1] : source[0];
        target[2] = source.Count > 2 ? source[2] : source[0];
        target[3] = source.Count > 3 ? source[3] : target[1];
    }

    private static bool TryStrokeWidth(string value, double reference, double fontSize, double rootFontSize, out double width) {
        width = 0D;
        switch (value.Trim().ToLowerInvariant()) {
            case "thin": width = 1D; return true;
            case "medium": width = 3D; return true;
            case "thick": width = 5D; return true;
        }
        return !value.EndsWith("%", StringComparison.Ordinal)
            && HtmlRenderCssValues.TryLength(value, reference, fontSize, rootFontSize, out width)
            && width >= 0D;
    }

    private static bool TryStrokeStyle(string value, out string style) {
        style = value.Trim().ToLowerInvariant();
        return style == "none" || style == "hidden" || style == "solid" || style == "dashed" || style == "dotted" || style == "double";
    }

    private static bool TryStrokeColor(string value, OfficeColor currentColor, out OfficeColor color) {
        if (string.Equals(value.Trim(), "currentcolor", StringComparison.OrdinalIgnoreCase)) {
            color = currentColor;
            return true;
        }
        return HtmlRenderCssValues.TryColor(value, out color);
    }
}
