using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlCssRadialGradientParser {
    internal static bool TryParse(
        string? value,
        int maximumStops,
        out HtmlCssRadialGradientDefinition? definition,
        out bool stopLimitExceeded) {
        definition = null;
        stopLimitExceeded = false;
        if (string.IsNullOrWhiteSpace(value) || maximumStops < 2) return false;

        string text = value!.Trim();
        string functionName;
        bool repeating;
        if (text.StartsWith("repeating-radial-gradient", StringComparison.OrdinalIgnoreCase)) {
            functionName = "repeating-radial-gradient";
            repeating = true;
        } else if (text.StartsWith("radial-gradient", StringComparison.OrdinalIgnoreCase)) {
            functionName = "radial-gradient";
            repeating = false;
        } else return false;
        int open = functionName.Length;
        if (open >= text.Length || text[open] != '(' || text[text.Length - 1] != ')') return false;

        if (!HtmlRenderCssValues.TrySplitTopLevelCommas(
                text.Substring(open + 1, text.Length - open - 2),
                maximumStops == int.MaxValue ? int.MaxValue : maximumStops + 1,
                out IReadOnlyList<string> arguments)) {
            stopLimitExceeded = true;
            return false;
        }
        if (arguments.Count < 2) return false;
        int stopStart = HtmlCssGradientStops.IsColorStop(arguments[0]) ? 0 : 1;
        if (!HtmlCssGradientStops.TryParse(arguments, stopStart, maximumStops, out HtmlCssGradientStops? stops, out stopLimitExceeded)
            || stops == null
            || !TryParseDescriptor(stopStart == 0 ? string.Empty : arguments[0], stops, repeating, out definition)) {
            return false;
        }

        return true;
    }

    private static bool TryParseDescriptor(
        string descriptor,
        HtmlCssGradientStops stops,
        bool repeating,
        out HtmlCssRadialGradientDefinition? definition) {
        definition = null;
        List<string> parts = HtmlRenderCssValues.SplitWhitespace(descriptor)
            .Select(part => part.ToLowerInvariant())
            .ToList();
        int at = parts.IndexOf("at");
        if (at >= 0 && at != parts.LastIndexOf("at")) return false;
        IReadOnlyList<string> positionParts = at >= 0 ? parts.Skip(at + 1).ToList() : Array.Empty<string>();
        if (at >= 0 && positionParts.Count == 0) return false;
        if (!TryParsePosition(positionParts, out string centerX, out string centerY)) return false;
        if (at >= 0) parts.RemoveRange(at, parts.Count - at);

        bool circle = RemoveSingle(parts, "circle", out bool duplicateCircle);
        bool ellipse = RemoveSingle(parts, "ellipse", out bool duplicateEllipse);
        if (duplicateCircle || duplicateEllipse || circle && ellipse) return false;

        HtmlCssRadialGradientShape shape;
        HtmlCssRadialGradientSize size;
        string? radiusX = null;
        string? radiusY = null;
        if (parts.Count == 0) {
            shape = circle ? HtmlCssRadialGradientShape.Circle : HtmlCssRadialGradientShape.Ellipse;
            size = HtmlCssRadialGradientSize.FarthestCorner;
        } else if (parts.Count == 1 && TryParseExtent(parts[0], out size)) {
            shape = circle ? HtmlCssRadialGradientShape.Circle : HtmlCssRadialGradientShape.Ellipse;
        } else if (parts.Count == 1
            && !ellipse
            && parts[0].IndexOf('%') < 0
            && IsNonNegativeLength(parts[0])) {
            shape = HtmlCssRadialGradientShape.Circle;
            size = HtmlCssRadialGradientSize.Explicit;
            radiusX = parts[0];
        } else if (parts.Count == 2
            && !circle
            && IsNonNegativeLength(parts[0])
            && IsNonNegativeLength(parts[1])) {
            shape = HtmlCssRadialGradientShape.Ellipse;
            size = HtmlCssRadialGradientSize.Explicit;
            radiusX = parts[0];
            radiusY = parts[1];
        } else {
            return false;
        }

        definition = new HtmlCssRadialGradientDefinition(shape, size, centerX, centerY, radiusX, radiusY, stops, repeating);
        return true;
    }

    private static bool TryParsePosition(IReadOnlyList<string> parts, out string x, out string y) {
        return HtmlCssGradientPositionParser.TryParse(parts, out x, out y);
    }

    private static bool TryParseExtent(string value, out HtmlCssRadialGradientSize size) {
        switch (value) {
            case "closest-side":
                size = HtmlCssRadialGradientSize.ClosestSide;
                return true;
            case "closest-corner":
                size = HtmlCssRadialGradientSize.ClosestCorner;
                return true;
            case "farthest-side":
                size = HtmlCssRadialGradientSize.FarthestSide;
                return true;
            case "farthest-corner":
                size = HtmlCssRadialGradientSize.FarthestCorner;
                return true;
            default:
                size = default;
                return false;
        }
    }

    private static bool RemoveSingle(ICollection<string> parts, string value, out bool duplicate) {
        duplicate = false;
        bool found = parts.Remove(value);
        if (found && parts.Contains(value)) duplicate = true;
        return found;
    }

    private static bool IsLength(string value) =>
        HtmlRenderCssValues.TryLength(value, 100D, 16D, 16D, 100D, 100D, out _);

    private static bool IsNonNegativeLength(string value) =>
        HtmlRenderCssValues.TryLength(value, 100D, 16D, 16D, 100D, 100D, out double length) && length >= 0D;
}
