namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private static string ResolveWritingMode(string value, string? inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return inherited ?? "horizontal-tb";
        if (normalized == "horizontal-tb" || normalized == "vertical-rl" || normalized == "vertical-lr"
            || normalized == "sideways-rl" || normalized == "sideways-lr") return normalized;
        return "horizontal-tb";
    }

    private static string ResolveTextOrientation(string value, string? inherited) {
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized.Length == 0 || normalized == "inherit" || normalized == "unset") return inherited ?? "mixed";
        return normalized == "mixed" || normalized == "upright" || normalized == "sideways"
            ? normalized
            : "mixed";
    }

    private static HtmlComputedStyle PhysicalizeLogicalProperties(
        HtmlComputedStyle computed,
        string writingMode,
        string direction) {
        var properties = new Dictionary<string, string>(HtmlCssPropertyNameComparer.Instance);
        foreach (KeyValuePair<string, string> property in computed.Properties) properties[property.Key] = property.Value;
        ResolveLogicalSides(writingMode, direction, out string inlineStart, out string inlineEnd, out string blockStart, out string blockEnd);

        MapPair(computed, properties, "margin-inline", "margin-" + inlineStart, "margin-" + inlineEnd);
        MapSingle(computed, properties, "margin-inline-start", "margin-" + inlineStart);
        MapSingle(computed, properties, "margin-inline-end", "margin-" + inlineEnd);
        MapPair(computed, properties, "margin-block", "margin-" + blockStart, "margin-" + blockEnd);
        MapSingle(computed, properties, "margin-block-start", "margin-" + blockStart);
        MapSingle(computed, properties, "margin-block-end", "margin-" + blockEnd);

        MapPair(computed, properties, "padding-inline", "padding-" + inlineStart, "padding-" + inlineEnd);
        MapSingle(computed, properties, "padding-inline-start", "padding-" + inlineStart);
        MapSingle(computed, properties, "padding-inline-end", "padding-" + inlineEnd);
        MapPair(computed, properties, "padding-block", "padding-" + blockStart, "padding-" + blockEnd);
        MapSingle(computed, properties, "padding-block-start", "padding-" + blockStart);
        MapSingle(computed, properties, "padding-block-end", "padding-" + blockEnd);

        MapPair(computed, properties, "inset-inline", inlineStart, inlineEnd);
        MapSingle(computed, properties, "inset-inline-start", inlineStart);
        MapSingle(computed, properties, "inset-inline-end", inlineEnd);
        MapPair(computed, properties, "inset-block", blockStart, blockEnd);
        MapSingle(computed, properties, "inset-block-start", blockStart);
        MapSingle(computed, properties, "inset-block-end", blockEnd);

        bool vertical = writingMode != "horizontal-tb";
        MapSingle(computed, properties, "inline-size", vertical ? "height" : "width");
        MapSingle(computed, properties, "block-size", vertical ? "width" : "height");
        MapSingle(computed, properties, "min-inline-size", vertical ? "min-height" : "min-width");
        MapSingle(computed, properties, "min-block-size", vertical ? "min-width" : "min-height");
        MapSingle(computed, properties, "max-inline-size", vertical ? "max-height" : "max-width");
        MapSingle(computed, properties, "max-block-size", vertical ? "max-width" : "max-height");

        MapBorderAxis(computed, properties, "inline", inlineStart, inlineEnd);
        MapBorderAxis(computed, properties, "block", blockStart, blockEnd);
        MapLogicalCornerRadius(computed, properties, "border-start-start-radius", blockStart, inlineStart);
        MapLogicalCornerRadius(computed, properties, "border-start-end-radius", blockStart, inlineEnd);
        MapLogicalCornerRadius(computed, properties, "border-end-start-radius", blockEnd, inlineStart);
        MapLogicalCornerRadius(computed, properties, "border-end-end-radius", blockEnd, inlineEnd);
        return new HtmlComputedStyle(properties);
    }

    private static void ResolveLogicalSides(
        string writingMode,
        string direction,
        out string inlineStart,
        out string inlineEnd,
        out string blockStart,
        out string blockEnd) {
        bool rtl = direction == "rtl";
        if (writingMode == "horizontal-tb") {
            inlineStart = rtl ? "right" : "left";
            inlineEnd = rtl ? "left" : "right";
            blockStart = "top";
            blockEnd = "bottom";
            return;
        }

        inlineStart = rtl ? "bottom" : "top";
        inlineEnd = rtl ? "top" : "bottom";
        bool rightToLeftBlocks = writingMode == "vertical-rl" || writingMode == "sideways-rl";
        blockStart = rightToLeftBlocks ? "right" : "left";
        blockEnd = rightToLeftBlocks ? "left" : "right";
    }

    private static void MapBorderAxis(
        HtmlComputedStyle computed,
        IDictionary<string, string> properties,
        string axis,
        string start,
        string end) {
        string prefix = "border-" + axis;
        MapPair(computed, properties, prefix, "border-" + start, "border-" + end);
        MapPair(computed, properties, prefix + "-width", "border-" + start + "-width", "border-" + end + "-width");
        MapPair(computed, properties, prefix + "-style", "border-" + start + "-style", "border-" + end + "-style");
        MapPair(computed, properties, prefix + "-color", "border-" + start + "-color", "border-" + end + "-color");
        MapSingle(computed, properties, prefix + "-start", "border-" + start);
        MapSingle(computed, properties, prefix + "-start-width", "border-" + start + "-width");
        MapSingle(computed, properties, prefix + "-start-style", "border-" + start + "-style");
        MapSingle(computed, properties, prefix + "-start-color", "border-" + start + "-color");
        MapSingle(computed, properties, prefix + "-end", "border-" + end);
        MapSingle(computed, properties, prefix + "-end-width", "border-" + end + "-width");
        MapSingle(computed, properties, prefix + "-end-style", "border-" + end + "-style");
        MapSingle(computed, properties, prefix + "-end-color", "border-" + end + "-color");
    }

    private static void MapLogicalCornerRadius(
        HtmlComputedStyle computed,
        IDictionary<string, string> properties,
        string logicalName,
        string firstSide,
        string secondSide) {
        string verticalSide = firstSide == "top" || firstSide == "bottom" ? firstSide : secondSide;
        string horizontalSide = firstSide == "left" || firstSide == "right" ? firstSide : secondSide;
        MapSingle(computed, properties, logicalName, "border-" + verticalSide + "-" + horizontalSide + "-radius");
    }

    private static void MapSingle(
        HtmlComputedStyle computed,
        IDictionary<string, string> properties,
        string logicalName,
        string physicalName) {
        string value = computed.GetValue(logicalName).Trim();
        if (value.Length > 0) properties[physicalName] = value;
    }

    private static void MapPair(
        HtmlComputedStyle computed,
        IDictionary<string, string> properties,
        string logicalName,
        string startName,
        string endName) {
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(computed.GetValue(logicalName));
        if (values.Count == 0 || values.Count > 2) return;
        properties[startName] = values[0];
        properties[endName] = values.Count == 1 ? values[0] : values[1];
    }
}
