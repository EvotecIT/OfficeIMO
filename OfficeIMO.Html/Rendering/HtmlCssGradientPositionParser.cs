namespace OfficeIMO.Html;

/// <summary>
/// Parses CSS gradient-center positions, including three- and four-component edge offsets.
/// </summary>
internal static class HtmlCssGradientPositionParser {
    internal static bool TryParse(IReadOnlyList<string> parts, out string x, out string y) {
        x = "50%";
        y = "50%";
        if (parts.Count == 0) return true;
        if (parts.Count > 4) return false;

        if (parts.Count == 1) {
            if (TryHorizontal(parts[0], out x)) return true;
            return TryVertical(parts[0], out y);
        }

        if (parts.Count == 2) {
            return TryHorizontal(parts[0], out x) && TryVertical(parts[1], out y)
                || TryHorizontal(parts[1], out x) && TryVertical(parts[0], out y);
        }

        return TryExtendedEdgeOffsets(parts, out x, out y);
    }

    private static bool TryExtendedEdgeOffsets(IReadOnlyList<string> parts, out string x, out string y) {
        x = string.Empty;
        y = string.Empty;
        bool hasHorizontal = false;
        bool hasVertical = false;
        for (int index = 0; index < parts.Count;) {
            string keyword = parts[index++].ToLowerInvariant();
            bool horizontal = keyword == "left" || keyword == "right";
            bool vertical = keyword == "top" || keyword == "bottom";
            if (!horizontal && !vertical || horizontal && hasHorizontal || vertical && hasVertical) return false;

            string? offset = null;
            if (index < parts.Count && IsLengthPercentage(parts[index])) offset = parts[index++];
            string resolved = ResolveEdge(keyword, offset);
            if (horizontal) {
                x = resolved;
                hasHorizontal = true;
            } else {
                y = resolved;
                hasVertical = true;
            }
        }

        return hasHorizontal && hasVertical;
    }

    private static string ResolveEdge(string keyword, string? offset) {
        bool farEdge = keyword == "right" || keyword == "bottom";
        if (offset == null) return farEdge ? "100%" : "0%";
        return farEdge ? "calc(100% - (" + offset + "))" : offset;
    }

    private static bool TryHorizontal(string value, out string result) {
        switch (value) {
            case "left": result = "0%"; return true;
            case "center": result = "50%"; return true;
            case "right": result = "100%"; return true;
            case "top":
            case "bottom": result = string.Empty; return false;
            default:
                result = value;
                return IsLengthPercentage(value);
        }
    }

    private static bool TryVertical(string value, out string result) {
        switch (value) {
            case "top": result = "0%"; return true;
            case "center": result = "50%"; return true;
            case "bottom": result = "100%"; return true;
            case "left":
            case "right": result = string.Empty; return false;
            default:
                result = value;
                return IsLengthPercentage(value);
        }
    }

    private static bool IsLengthPercentage(string value) =>
        HtmlRenderCssValues.HasExplicitLengthSyntax(value, allowPercentage: true, allowUnitlessZero: true)
        && HtmlRenderCssValues.TryLength(value, 100D, 16D, 16D, 100D, 100D, out double resolved)
        && !double.IsNaN(resolved)
        && !double.IsInfinity(resolved);
}
