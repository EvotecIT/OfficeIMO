using System.Globalization;

namespace OfficeIMO.Html;

internal static class HtmlCssConicGradientParser {
    internal static bool TryParse(
        string? value,
        int maximumStops,
        out HtmlCssConicGradientDefinition? definition,
        out bool stopLimitExceeded) {
        definition = null;
        stopLimitExceeded = false;
        if (string.IsNullOrWhiteSpace(value) || maximumStops < 2) return false;
        string text = value!.Trim();
        string functionName;
        bool repeating;
        if (text.StartsWith("repeating-conic-gradient", StringComparison.OrdinalIgnoreCase)) {
            functionName = "repeating-conic-gradient";
            repeating = true;
        } else if (text.StartsWith("conic-gradient", StringComparison.OrdinalIgnoreCase)) {
            functionName = "conic-gradient";
            repeating = false;
        } else return false;
        int open = functionName.Length;
        if (open >= text.Length || text[open] != '(' || text[text.Length - 1] != ')') return false;
        int maximumArguments = maximumStops == int.MaxValue ? int.MaxValue : maximumStops + 2;
        if (!HtmlRenderCssValues.TrySplitTopLevelCommas(
                text.Substring(open + 1, text.Length - open - 2),
                maximumArguments,
                out IReadOnlyList<string> arguments)) {
            stopLimitExceeded = true;
            return false;
        }
        arguments = NormalizeSerializedDescriptor(arguments);
        if (arguments.Count < 2) return false;
        int stopStart = HtmlCssGradientStops.IsConicColorStop(arguments[0]) ? 0 : 1;
        if (!HtmlCssGradientStops.TryParseConic(arguments, stopStart, maximumStops, out HtmlCssGradientStops? stops, out stopLimitExceeded)
            || stops == null
            || !TryParseDescriptor(stopStart == 0 ? string.Empty : arguments[0], out double angle, out string centerX, out string centerY)) return false;
        definition = new HtmlCssConicGradientDefinition(angle, centerX, centerY, stops, repeating);
        return true;
    }

    private static IReadOnlyList<string> NormalizeSerializedDescriptor(IReadOnlyList<string> arguments) {
        if (arguments.Count < 3
            || !arguments[0].TrimStart().StartsWith("from ", StringComparison.OrdinalIgnoreCase)
            || !arguments[1].TrimStart().StartsWith("at ", StringComparison.OrdinalIgnoreCase)) return arguments;
        var normalized = new List<string>(arguments.Count - 1) {
            arguments[0].Trim() + " " + arguments[1].Trim()
        };
        for (int index = 2; index < arguments.Count; index++) normalized.Add(arguments[index]);
        return normalized;
    }

    private static bool TryParseDescriptor(string descriptor, out double angle, out string centerX, out string centerY) {
        angle = 0D;
        centerX = "50%";
        centerY = "50%";
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(descriptor.Trim().ToLowerInvariant());
        int index = 0;
        if (index < parts.Count && parts[index] == "from") {
            if (++index >= parts.Count || !TryAngle(parts[index++], out angle)) return false;
        }
        if (index < parts.Count && parts[index] == "at") {
            index++;
            int remaining = parts.Count - index;
            if (remaining < 1 || remaining > 2) return false;
            if (!TryPosition(parts.Skip(index).ToList(), out centerX, out centerY)) return false;
            index = parts.Count;
        }
        return index == parts.Count;
    }

    private static bool TryAngle(string value, out double degrees) {
        degrees = 0D;
        string normalized = value.Trim().ToLowerInvariant();
        double multiplier;
        int suffix;
        if (normalized.EndsWith("deg", StringComparison.Ordinal)) { multiplier = 1D; suffix = 3; }
        else if (normalized.EndsWith("grad", StringComparison.Ordinal)) { multiplier = 0.9D; suffix = 4; }
        else if (normalized.EndsWith("rad", StringComparison.Ordinal)) { multiplier = 180D / Math.PI; suffix = 3; }
        else if (normalized.EndsWith("turn", StringComparison.Ordinal)) { multiplier = 360D; suffix = 4; }
        else return false;
        if (!double.TryParse(normalized.Substring(0, normalized.Length - suffix), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            || double.IsNaN(number)
            || double.IsInfinity(number)) return false;
        degrees = number * multiplier;
        return !double.IsNaN(degrees) && !double.IsInfinity(degrees);
    }

    private static bool TryPosition(IReadOnlyList<string> parts, out string x, out string y) {
        x = "50%";
        y = "50%";
        if (parts.Count == 1) {
            if (TryHorizontal(parts[0], out x)) return true;
            return TryVertical(parts[0], out y);
        }
        return TryHorizontal(parts[0], out x) && TryVertical(parts[1], out y)
            || TryHorizontal(parts[1], out x) && TryVertical(parts[0], out y);
    }

    private static bool TryHorizontal(string value, out string result) {
        if (value == "left") { result = "0%"; return true; }
        if (value == "center") { result = "50%"; return true; }
        if (value == "right") { result = "100%"; return true; }
        if (value == "top" || value == "bottom") { result = string.Empty; return false; }
        result = value;
        return HtmlRenderCssValues.TryLength(value, 100D, 16D, 16D, 100D, 100D, out _);
    }

    private static bool TryVertical(string value, out string result) {
        if (value == "top") { result = "0%"; return true; }
        if (value == "center") { result = "50%"; return true; }
        if (value == "bottom") { result = "100%"; return true; }
        if (value == "left" || value == "right") { result = string.Empty; return false; }
        result = value;
        return HtmlRenderCssValues.TryLength(value, 100D, 16D, 16D, 100D, 100D, out _);
    }
}
