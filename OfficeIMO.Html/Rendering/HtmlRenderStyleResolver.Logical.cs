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
        Dictionary<string, HtmlCssCascadePriority> priorities = computed.CopyCascadePriorities();
        ResolveLogicalSides(writingMode, direction, out string inlineStart, out string inlineEnd, out string blockStart, out string blockEnd);

        MapPair(computed, properties, priorities, "margin-inline", "margin-" + inlineStart, "margin-" + inlineEnd);
        MapSingle(computed, properties, priorities, "margin-inline-start", "margin-" + inlineStart);
        MapSingle(computed, properties, priorities, "margin-inline-end", "margin-" + inlineEnd);
        MapPair(computed, properties, priorities, "margin-block", "margin-" + blockStart, "margin-" + blockEnd);
        MapSingle(computed, properties, priorities, "margin-block-start", "margin-" + blockStart);
        MapSingle(computed, properties, priorities, "margin-block-end", "margin-" + blockEnd);

        MapPair(computed, properties, priorities, "padding-inline", "padding-" + inlineStart, "padding-" + inlineEnd);
        MapSingle(computed, properties, priorities, "padding-inline-start", "padding-" + inlineStart);
        MapSingle(computed, properties, priorities, "padding-inline-end", "padding-" + inlineEnd);
        MapPair(computed, properties, priorities, "padding-block", "padding-" + blockStart, "padding-" + blockEnd);
        MapSingle(computed, properties, priorities, "padding-block-start", "padding-" + blockStart);
        MapSingle(computed, properties, priorities, "padding-block-end", "padding-" + blockEnd);

        MapPair(computed, properties, priorities, "inset-inline", inlineStart, inlineEnd);
        MapSingle(computed, properties, priorities, "inset-inline-start", inlineStart);
        MapSingle(computed, properties, priorities, "inset-inline-end", inlineEnd);
        MapPair(computed, properties, priorities, "inset-block", blockStart, blockEnd);
        MapSingle(computed, properties, priorities, "inset-block-start", blockStart);
        MapSingle(computed, properties, priorities, "inset-block-end", blockEnd);

        bool vertical = writingMode != "horizontal-tb";
        MapSingle(computed, properties, priorities, "inline-size", vertical ? "height" : "width");
        MapSingle(computed, properties, priorities, "block-size", vertical ? "width" : "height");
        MapSingle(computed, properties, priorities, "min-inline-size", vertical ? "min-height" : "min-width");
        MapSingle(computed, properties, priorities, "min-block-size", vertical ? "min-width" : "min-height");
        MapSingle(computed, properties, priorities, "max-inline-size", vertical ? "max-height" : "max-width");
        MapSingle(computed, properties, priorities, "max-block-size", vertical ? "max-width" : "max-height");

        MapBorderAxis(computed, properties, priorities, "inline", inlineStart, inlineEnd);
        MapBorderAxis(computed, properties, priorities, "block", blockStart, blockEnd);
        MapLogicalCornerRadius(computed, properties, priorities, "border-start-start-radius", blockStart, inlineStart);
        MapLogicalCornerRadius(computed, properties, priorities, "border-start-end-radius", blockStart, inlineEnd);
        MapLogicalCornerRadius(computed, properties, priorities, "border-end-start-radius", blockEnd, inlineStart);
        MapLogicalCornerRadius(computed, properties, priorities, "border-end-end-radius", blockEnd, inlineEnd);
        return computed.WithMappedProperties(properties, priorities);
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

        bool inlineStartsAtBottom = rtl ^ (writingMode == "sideways-lr");
        inlineStart = inlineStartsAtBottom ? "bottom" : "top";
        inlineEnd = inlineStartsAtBottom ? "top" : "bottom";
        bool rightToLeftBlocks = writingMode == "vertical-rl" || writingMode == "sideways-rl";
        blockStart = rightToLeftBlocks ? "right" : "left";
        blockEnd = rightToLeftBlocks ? "left" : "right";
    }

    private static void MapBorderAxis(
        HtmlComputedStyle computed,
        IDictionary<string, string> properties,
        Dictionary<string, HtmlCssCascadePriority> priorities,
        string axis,
        string start,
        string end) {
        string prefix = "border-" + axis;
        MapPair(computed, properties, priorities, prefix, "border-" + start, "border-" + end);
        MapPair(computed, properties, priorities, prefix + "-width", "border-" + start + "-width", "border-" + end + "-width");
        MapPair(computed, properties, priorities, prefix + "-style", "border-" + start + "-style", "border-" + end + "-style");
        MapPair(computed, properties, priorities, prefix + "-color", "border-" + start + "-color", "border-" + end + "-color");
        MapSingle(computed, properties, priorities, prefix + "-start", "border-" + start);
        MapSingle(computed, properties, priorities, prefix + "-start-width", "border-" + start + "-width");
        MapSingle(computed, properties, priorities, prefix + "-start-style", "border-" + start + "-style");
        MapSingle(computed, properties, priorities, prefix + "-start-color", "border-" + start + "-color");
        MapSingle(computed, properties, priorities, prefix + "-end", "border-" + end);
        MapSingle(computed, properties, priorities, prefix + "-end-width", "border-" + end + "-width");
        MapSingle(computed, properties, priorities, prefix + "-end-style", "border-" + end + "-style");
        MapSingle(computed, properties, priorities, prefix + "-end-color", "border-" + end + "-color");
    }

    private static void MapLogicalCornerRadius(
        HtmlComputedStyle computed,
        IDictionary<string, string> properties,
        Dictionary<string, HtmlCssCascadePriority> priorities,
        string logicalName,
        string firstSide,
        string secondSide) {
        string verticalSide = firstSide == "top" || firstSide == "bottom" ? firstSide : secondSide;
        string horizontalSide = firstSide == "left" || firstSide == "right" ? firstSide : secondSide;
        MapSingle(computed, properties, priorities, logicalName, "border-" + verticalSide + "-" + horizontalSide + "-radius");
    }

    private static void MapSingle(
        HtmlComputedStyle computed,
        IDictionary<string, string> properties,
        Dictionary<string, HtmlCssCascadePriority> priorities,
        string logicalName,
        string physicalName) {
        string value = computed.GetValue(logicalName).Trim();
        if (value.Length == 0 || !computed.ShouldOverride(logicalName, physicalName)) return;
        properties[physicalName] = value;
        if (computed.TryGetCascadePriority(logicalName, out HtmlCssCascadePriority priority)) priorities[physicalName] = priority;
    }

    private static void MapPair(
        HtmlComputedStyle computed,
        IDictionary<string, string> properties,
        Dictionary<string, HtmlCssCascadePriority> priorities,
        string logicalName,
        string startName,
        string endName) {
        IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(computed.GetValue(logicalName));
        if (values.Count == 0 || values.Count > 2) return;
        if (computed.ShouldOverride(logicalName, startName)) {
            properties[startName] = values[0];
            if (computed.TryGetCascadePriority(logicalName, out HtmlCssCascadePriority startPriority)) priorities[startName] = startPriority;
        }
        if (computed.ShouldOverride(logicalName, endName)) {
            properties[endName] = values.Count == 1 ? values[0] : values[1];
            if (computed.TryGetCascadePriority(logicalName, out HtmlCssCascadePriority endPriority)) priorities[endName] = endPriority;
        }
    }
}
