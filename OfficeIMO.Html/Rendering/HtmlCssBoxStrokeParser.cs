using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssBoxStrokeParser {
    private static readonly string[] SideProperties = {
        "border-top", "border-right", "border-bottom", "border-left",
        "border-top-width", "border-right-width", "border-bottom-width", "border-left-width",
        "border-top-style", "border-right-style", "border-bottom-style", "border-left-style",
        "border-top-color", "border-right-color", "border-bottom-color", "border-left-color"
    };

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
        out double width,
        out string style,
        out OfficeColor color,
        out string detail) {
        string shorthand = computed.GetValue("border").Trim();
        string widthValue = computed.GetValue("border-width").Trim();
        string styleValue = computed.GetValue("border-style").Trim();
        string colorValue = computed.GetValue("border-color").Trim();
        width = 3D;
        style = "none";
        color = currentColor;
        detail = string.Empty;
        bool declared = shorthand.Length > 0 || widthValue.Length > 0 || styleValue.Length > 0 || colorValue.Length > 0;
        string side = SideProperties.FirstOrDefault(property => computed.GetValue(property).Length > 0) ?? string.Empty;
        if (side.Length > 0) {
            width = 0D;
            detail = side + "=" + computed.GetValue(side).Trim() + ";asymmetric-side";
            return false;
        }
        if (!declared) {
            width = 0D;
            return true;
        }
        if (shorthand.Length > 0 && !TryParseStrokeShorthand(shorthand, reference, fontSize, rootFontSize, currentColor, ref width, ref style, ref color)) {
            width = 0D;
            detail = "border=" + shorthand;
            return false;
        }
        if (widthValue.Length > 0 && !TryParseUniformWidths(widthValue, reference, fontSize, rootFontSize, out width)) {
            width = 0D;
            detail = "border-width=" + widthValue;
            return false;
        }
        if (styleValue.Length > 0 && !TryParseUniformStyles(styleValue, out style)) {
            width = 0D;
            detail = "border-style=" + styleValue;
            return false;
        }
        if (colorValue.Length > 0 && !TryParseUniformColors(colorValue, currentColor, out color)) {
            width = 0D;
            detail = "border-color=" + colorValue;
            return false;
        }
        if (style == "none" || style == "hidden") width = 0D;
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
    internal static bool IsSupportedWidthSyntax(string value) => TryParseUniformWidths(value, 100D, 16D, 16D, out _);
    internal static bool IsSupportedStyleSyntax(string value) => TryParseUniformStyles(value, out _);
    internal static bool IsSupportedColorSyntax(string value) => TryParseUniformColors(value, OfficeColor.Black, out _);

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

    private static bool TryParseUniformWidths(string value, double reference, double fontSize, double rootFontSize, out double width) {
        width = 0D;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 4 || !TryStrokeWidth(tokens[0], reference, fontSize, rootFontSize, out width)) return false;
        for (int index = 1; index < tokens.Count; index++) {
            if (!TryStrokeWidth(tokens[index], reference, fontSize, rootFontSize, out double other) || Math.Abs(other - width) > 0.0001D) return false;
        }
        return true;
    }

    private static bool TryParseUniformStyles(string value, out string style) {
        style = "none";
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 4 || !TryStrokeStyle(tokens[0], out style)) return false;
        string expected = style;
        for (int index = 1; index < tokens.Count; index++) {
            if (!TryStrokeStyle(tokens[index], out string other) || other != expected) return false;
        }
        return true;
    }

    private static bool TryParseUniformColors(string value, OfficeColor currentColor, out OfficeColor color) {
        color = OfficeColor.Black;
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count < 1 || tokens.Count > 4 || !TryStrokeColor(tokens[0], currentColor, out color)) return false;
        OfficeColor expected = color;
        for (int index = 1; index < tokens.Count; index++) {
            if (!TryStrokeColor(tokens[index], currentColor, out OfficeColor other) || other != expected) return false;
        }
        return true;
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
