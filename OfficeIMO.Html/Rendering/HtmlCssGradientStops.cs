using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

internal static class HtmlCssGradientStops {
    internal static bool TryParse(
        IReadOnlyList<string> arguments,
        int startIndex,
        int maximumStops,
        out IReadOnlyList<OfficeGradientStop>? stops,
        out bool stopLimitExceeded) {
        stops = null;
        stopLimitExceeded = false;
        int stopCount = arguments.Count - startIndex;
        if (stopCount < 2) return false;
        if (stopCount > maximumStops) {
            stopLimitExceeded = true;
            return false;
        }

        var colors = new OfficeColor[stopCount];
        var offsets = new double?[stopCount];
        for (int index = 0; index < stopCount; index++) {
            if (!TryParseColorStop(arguments[index + startIndex], out colors[index], out offsets[index])) return false;
        }

        return TryResolveStops(colors, offsets, maximumStops, out stops, out stopLimitExceeded);
    }

    internal static bool IsColorStop(string value) => TryParseColorStop(value, out _, out _);

    private static bool TryParseColorStop(string value, out OfficeColor color, out double? offset) {
        color = default;
        offset = null;
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
        if (parts.Count == 0 || parts.Count > 2 || !HtmlRenderCssValues.TryColor(parts[0], out color) || color.A != 255) return false;
        if (parts.Count == 1) return true;

        string position = parts[1].Trim();
        if (position == "0") {
            offset = 0D;
            return true;
        }

        if (!position.EndsWith("%", StringComparison.Ordinal)
            || !double.TryParse(position.Substring(0, position.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)
            || double.IsNaN(percent)
            || double.IsInfinity(percent)
            || percent < 0D
            || percent > 100D) {
            return false;
        }

        offset = percent / 100D;
        return true;
    }

    private static bool TryResolveStops(
        IReadOnlyList<OfficeColor> colors,
        double?[] offsets,
        int maximumStops,
        out IReadOnlyList<OfficeGradientStop>? stops,
        out bool stopLimitExceeded) {
        stops = null;
        stopLimitExceeded = false;
        offsets[0] ??= 0D;
        offsets[offsets.Length - 1] ??= 1D;

        int previousSpecified = 0;
        for (int index = 1; index < offsets.Length; index++) {
            if (!offsets[index].HasValue) continue;
            double previous = offsets[previousSpecified]!.Value;
            double current = offsets[index]!.Value;
            if (current <= previous) return false;
            int gap = index - previousSpecified;
            for (int fill = 1; fill < gap; fill++) {
                offsets[previousSpecified + fill] = previous + ((current - previous) * fill / gap);
            }

            previousSpecified = index;
        }

        bool addLeading = offsets[0]!.Value > 0D;
        bool addTrailing = offsets[offsets.Length - 1]!.Value < 1D;
        int resolvedCount = offsets.Length + (addLeading ? 1 : 0) + (addTrailing ? 1 : 0);
        if (resolvedCount > maximumStops) {
            stopLimitExceeded = true;
            return false;
        }

        var resolved = new List<OfficeGradientStop>(resolvedCount);
        if (addLeading) resolved.Add(new OfficeGradientStop(0D, colors[0]));
        for (int index = 0; index < offsets.Length; index++) {
            resolved.Add(new OfficeGradientStop(offsets[index]!.Value, colors[index]));
        }

        if (addTrailing) resolved.Add(new OfficeGradientStop(1D, colors[colors.Count - 1]));
        stops = resolved.AsReadOnly();
        return true;
    }
}
