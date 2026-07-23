using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlCssGradientStops {
    private readonly IReadOnlyList<HtmlCssGradientStop> _stops;
    private readonly int _maximumStops;

    private HtmlCssGradientStops(IReadOnlyList<HtmlCssGradientStop> stops, int maximumStops) {
        _stops = stops;
        _maximumStops = maximumStops;
    }

    internal static bool TryParse(
        IReadOnlyList<string> arguments,
        int startIndex,
        int maximumStops,
        out HtmlCssGradientStops? definition,
        out bool stopLimitExceeded) {
        definition = null;
        stopLimitExceeded = false;
        int stopCount = arguments.Count - startIndex;
        if (stopCount < 2) return false;
        if (stopCount > maximumStops) {
            stopLimitExceeded = true;
            return false;
        }

        var stops = new List<HtmlCssGradientStop>(stopCount);
        for (int index = 0; index < stopCount; index++) {
            if (!TryParseColorStop(arguments[index + startIndex], out OfficeColor color, out string? firstPosition, out string? secondPosition)) return false;
            stops.Add(new HtmlCssGradientStop(color, firstPosition));
            if (secondPosition != null) stops.Add(new HtmlCssGradientStop(color, secondPosition));
            if (stops.Count > maximumStops) {
                stopLimitExceeded = true;
                return false;
            }
        }
        definition = new HtmlCssGradientStops(stops.AsReadOnly(), maximumStops);
        return true;
    }

    internal static bool IsColorStop(string value) => TryParseColorStop(value, out _, out _, out _);

    internal bool TryResolve(double referenceLength, double fontSize, double rootFontSize, out IReadOnlyList<OfficeGradientStop>? stops) {
        stops = null;
        if (referenceLength <= 0D || double.IsNaN(referenceLength) || double.IsInfinity(referenceLength)) return false;
        var colors = new OfficeColor[_stops.Count];
        var offsets = new double?[_stops.Count];
        for (int index = 0; index < _stops.Count; index++) {
            HtmlCssGradientStop stop = _stops[index];
            colors[index] = stop.Color;
            if (stop.Position == null) continue;
            if (!HtmlRenderCssValues.TryLength(stop.Position, referenceLength, fontSize, rootFontSize, out double pixels)) return false;
            double offset = pixels / referenceLength;
            if (double.IsNaN(offset) || double.IsInfinity(offset) || offset < 0D || offset > 1D) return false;
            offsets[index] = offset;
        }

        offsets[0] ??= 0D;
        offsets[offsets.Length - 1] ??= 1D;
        int previousSpecified = 0;
        for (int index = 1; index < offsets.Length; index++) {
            if (!offsets[index].HasValue) continue;
            double previous = offsets[previousSpecified]!.Value;
            double current = Math.Max(previous, offsets[index]!.Value);
            offsets[index] = current;
            int gap = index - previousSpecified;
            for (int fill = 1; fill < gap; fill++) offsets[previousSpecified + fill] = previous + ((current - previous) * fill / gap);
            previousSpecified = index;
        }

        bool addLeading = offsets[0]!.Value > 0D;
        bool addTrailing = offsets[offsets.Length - 1]!.Value < 1D;
        int resolvedCount = offsets.Length + (addLeading ? 1 : 0) + (addTrailing ? 1 : 0);
        if (resolvedCount > _maximumStops) return false;
        var resolved = new List<OfficeGradientStop>(resolvedCount);
        if (addLeading) resolved.Add(new OfficeGradientStop(0D, colors[0]));
        for (int index = 0; index < offsets.Length; index++) resolved.Add(new OfficeGradientStop(offsets[index]!.Value, colors[index]));
        if (addTrailing) resolved.Add(new OfficeGradientStop(1D, colors[colors.Length - 1]));
        stops = resolved.AsReadOnly();
        return true;
    }

    private static bool TryParseColorStop(string value, out OfficeColor color, out string? firstPosition, out string? secondPosition) {
        color = default;
        firstPosition = null;
        secondPosition = null;
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
        if (parts.Count == 0 || parts.Count > 3 || !HtmlRenderCssValues.TryColor(parts[0], out color) || color.A != 255) return false;
        if (parts.Count >= 2) {
            firstPosition = parts[1].Trim();
            if (!IsStopPosition(firstPosition)) return false;
        }
        if (parts.Count == 3) {
            secondPosition = parts[2].Trim();
            if (!IsStopPosition(secondPosition)) return false;
        }
        return true;
    }

    private static bool IsStopPosition(string value) {
        if (value == "0") return true;
        return HtmlRenderCssValues.TryLength(value, 100D, 16D, 16D, out double result) && result >= 0D;
    }

    private sealed class HtmlCssGradientStop {
        internal HtmlCssGradientStop(OfficeColor color, string? position) {
            Color = color;
            Position = position;
        }
        internal OfficeColor Color { get; }
        internal string? Position { get; }
    }
}
