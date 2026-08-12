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

    internal static bool TryParseConic(
        IReadOnlyList<string> arguments,
        int startIndex,
        int maximumStops,
        out HtmlCssGradientStops? definition,
        out bool stopLimitExceeded) {
        definition = null;
        stopLimitExceeded = false;
        int stopCount = arguments.Count - startIndex;
        if (stopCount < 2) return false;
        var stops = new List<HtmlCssGradientStop>(stopCount);
        for (int index = 0; index < stopCount; index++) {
            if (!TryParseConicColorStop(arguments[index + startIndex], out OfficeColor color, out string? firstPosition, out string? secondPosition)) return false;
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

    internal static bool IsConicColorStop(string value) => TryParseConicColorStop(value, out _, out _, out _);

    internal bool TryResolve(double referenceLength, double fontSize, double rootFontSize, double viewportWidth, double viewportHeight, out IReadOnlyList<OfficeGradientStop>? stops) {
        return TryResolve(referenceLength, fontSize, rootFontSize, viewportWidth, viewportHeight, repeating: false, out stops, out _);
    }

    internal bool TryResolve(
        double referenceLength,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        bool repeating,
        out IReadOnlyList<OfficeGradientStop>? stops,
        out bool stopLimitExceeded) {
        stops = null;
        stopLimitExceeded = false;
        if (referenceLength <= 0D || double.IsNaN(referenceLength) || double.IsInfinity(referenceLength)) return false;
        var colors = new OfficeColor[_stops.Count];
        var offsets = new double?[_stops.Count];
        for (int index = 0; index < _stops.Count; index++) {
            HtmlCssGradientStop stop = _stops[index];
            colors[index] = stop.Color;
            if (stop.Position == null) continue;
            if (!HtmlRenderCssValues.TryLength(stop.Position, referenceLength, fontSize, rootFontSize, viewportWidth, viewportHeight, out double pixels)) return false;
            double offset = pixels / referenceLength;
            if (double.IsNaN(offset) || double.IsInfinity(offset)) return false;
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

        var raw = new List<ResolvedStop>(offsets.Length);
        for (int index = 0; index < offsets.Length; index++) raw.Add(new ResolvedStop(offsets[index]!.Value, colors[index]));
        List<ResolvedStop> clipped = repeating
            ? ExpandRepeating(raw, _maximumStops, out stopLimitExceeded)
            : ClipToUnit(raw);
        if (clipped.Count < 2 || clipped.Count > _maximumStops) {
            stopLimitExceeded = stopLimitExceeded || clipped.Count > _maximumStops;
            return false;
        }
        var resolved = new List<OfficeGradientStop>(clipped.Count);
        foreach (ResolvedStop stop in clipped) resolved.Add(new OfficeGradientStop(stop.Offset, stop.Color));
        stops = resolved.AsReadOnly();
        return true;
    }

    internal bool TryResolveConic(
        bool repeating,
        out IReadOnlyList<OfficeGradientStop>? stops,
        out bool stopLimitExceeded) {
        stops = null;
        stopLimitExceeded = false;
        var colors = new OfficeColor[_stops.Count];
        var offsets = new double?[_stops.Count];
        for (int index = 0; index < _stops.Count; index++) {
            HtmlCssGradientStop stop = _stops[index];
            colors[index] = stop.Color;
            if (stop.Position == null) continue;
            if (!TryConicPosition(stop.Position, out double offset)) return false;
            offsets[index] = offset;
        }
        ResolveImplicitOffsets(offsets);
        var raw = new List<ResolvedStop>(offsets.Length);
        for (int index = 0; index < offsets.Length; index++) raw.Add(new ResolvedStop(offsets[index]!.Value, colors[index]));
        List<ResolvedStop> clipped = repeating
            ? ExpandRepeating(raw, _maximumStops, out stopLimitExceeded)
            : ClipToUnit(raw);
        if (clipped.Count < 2 || clipped.Count > _maximumStops) {
            stopLimitExceeded = stopLimitExceeded || clipped.Count > _maximumStops;
            return false;
        }
        stops = clipped.Select(stop => new OfficeGradientStop(stop.Offset, stop.Color)).ToList().AsReadOnly();
        return true;
    }

    private static bool TryParseColorStop(string value, out OfficeColor color, out string? firstPosition, out string? secondPosition) {
        return TryParseColorStopCore(value, conic: false, out color, out firstPosition, out secondPosition);
    }

    private static bool TryParseConicColorStop(string value, out OfficeColor color, out string? firstPosition, out string? secondPosition) {
        return TryParseColorStopCore(value, conic: true, out color, out firstPosition, out secondPosition);
    }

    private static bool TryParseColorStopCore(string value, bool conic, out OfficeColor color, out string? firstPosition, out string? secondPosition) {
        color = default;
        firstPosition = null;
        secondPosition = null;
        if (!TrySplitColorAndPositions(value, out string colorText, out IReadOnlyList<string> parts)
            || parts.Count > 2
            || !HtmlRenderCssValues.TryColor(colorText, out color)) return false;
        if (parts.Count >= 1) {
            firstPosition = parts[0].Trim();
            if (!(conic ? IsConicStopPosition(firstPosition) : IsStopPosition(firstPosition))) return false;
        }
        if (parts.Count == 2) {
            secondPosition = parts[1].Trim();
            if (!(conic ? IsConicStopPosition(secondPosition) : IsStopPosition(secondPosition))) return false;
        }
        return true;
    }

    private static bool TrySplitColorAndPositions(string value, out string color, out IReadOnlyList<string> positions) {
        color = string.Empty;
        positions = Array.Empty<string>();
        string text = value.Trim();
        if (text.Length == 0) return false;
        int open = text.IndexOf('(');
        int whitespace = IndexOfWhitespace(text);
        int colorEnd;
        if (open >= 0 && (whitespace < 0 || open < whitespace)) {
            int depth = 0;
            colorEnd = -1;
            for (int index = open; index < text.Length; index++) {
                if (text[index] == '(') depth++;
                else if (text[index] == ')' && --depth == 0) {
                    colorEnd = index + 1;
                    break;
                }
                if (depth < 0 || depth > 8) return false;
            }
            if (colorEnd < 0) return false;
        } else {
            colorEnd = whitespace < 0 ? text.Length : whitespace;
        }
        color = text.Substring(0, colorEnd).Trim();
        positions = HtmlRenderCssValues.SplitWhitespace(text.Substring(colorEnd));
        return color.Length > 0;
    }

    private static int IndexOfWhitespace(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (char.IsWhiteSpace(value[index])) return index;
        }
        return -1;
    }

    private static bool IsStopPosition(string value) {
        if (value == "0") return true;
        return HtmlRenderCssValues.TryLength(value, 100D, 16D, 16D, 100D, 100D, out double result)
            && !double.IsNaN(result)
            && !double.IsInfinity(result);
    }

    private static bool IsConicStopPosition(string value) => TryConicPosition(value, out _);

    private static bool TryConicPosition(string value, out double turns) {
        turns = 0D;
        string normalized = value.Trim().ToLowerInvariant();
        if (normalized == "0") return true;
        double divisor;
        int suffixLength;
        if (normalized.EndsWith("%", StringComparison.Ordinal)) { divisor = 100D; suffixLength = 1; }
        else if (normalized.EndsWith("deg", StringComparison.Ordinal)) { divisor = 360D; suffixLength = 3; }
        else if (normalized.EndsWith("grad", StringComparison.Ordinal)) { divisor = 400D; suffixLength = 4; }
        else if (normalized.EndsWith("rad", StringComparison.Ordinal)) { divisor = 2D * Math.PI; suffixLength = 3; }
        else if (normalized.EndsWith("turn", StringComparison.Ordinal)) { divisor = 1D; suffixLength = 4; }
        else return false;
        if (!double.TryParse(normalized.Substring(0, normalized.Length - suffixLength).Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double number)
            || double.IsNaN(number)
            || double.IsInfinity(number)) return false;
        turns = number / divisor;
        return !double.IsNaN(turns) && !double.IsInfinity(turns);
    }

    private static void ResolveImplicitOffsets(double?[] offsets) {
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
    }

    private static List<ResolvedStop> ClipToUnit(IReadOnlyList<ResolvedStop> source) {
        var result = new List<ResolvedStop>(source.Count + 2);
        if (source[0].Offset > 0D) result.Add(new ResolvedStop(0D, source[0].Color));
        else if (source[0].Offset < 0D) result.Add(new ResolvedStop(0D, Sample(source, 0D)));
        foreach (ResolvedStop stop in source) {
            if (stop.Offset >= 0D && stop.Offset <= 1D) result.Add(stop);
        }
        if (result.Count == 0 || result[0].Offset > 0D) result.Insert(0, new ResolvedStop(0D, Sample(source, 0D)));
        if (source[source.Count - 1].Offset < 1D) result.Add(new ResolvedStop(1D, source[source.Count - 1].Color));
        else if (source[source.Count - 1].Offset > 1D) result.Add(new ResolvedStop(1D, Sample(source, 1D)));
        if (result[result.Count - 1].Offset < 1D) result.Add(new ResolvedStop(1D, Sample(source, 1D)));
        return result;
    }

    private static List<ResolvedStop> ExpandRepeating(
        IReadOnlyList<ResolvedStop> source,
        int maximumStops,
        out bool stopLimitExceeded) {
        stopLimitExceeded = false;
        double first = source[0].Offset;
        double last = source[source.Count - 1].Offset;
        double period = last - first;
        if (period <= 0D || double.IsNaN(period) || double.IsInfinity(period)) {
            OfficeColor solid = source[source.Count - 1].Color;
            return new List<ResolvedStop> { new ResolvedStop(0D, solid), new ResolvedStop(1D, solid) };
        }

        long firstCycle = (long)Math.Floor((0D - first) / period) - 1L;
        long lastCycle = (long)Math.Ceiling((1D - first) / period) + 1L;
        if (lastCycle - firstCycle > maximumStops + 2L) {
            stopLimitExceeded = true;
            return new List<ResolvedStop>();
        }
        var expanded = new List<ResolvedStop>();
        for (long cycle = firstCycle; cycle <= lastCycle; cycle++) {
            double shift = cycle * period;
            foreach (ResolvedStop stop in source) {
                double offset = stop.Offset + shift;
                if (offset >= 0D && offset <= 1D) expanded.Add(new ResolvedStop(offset, stop.Color));
                if (expanded.Count > maximumStops) {
                    stopLimitExceeded = true;
                    return new List<ResolvedStop>();
                }
            }
        }
        expanded = expanded.OrderBy(stop => stop.Offset).ToList();
        OfficeColor start = SampleRepeating(source, period, 0D);
        OfficeColor end = SampleRepeating(source, period, 1D);
        if (expanded.Count == 0 || expanded[0].Offset > 0D) expanded.Insert(0, new ResolvedStop(0D, start));
        if (expanded[expanded.Count - 1].Offset < 1D) expanded.Add(new ResolvedStop(1D, end));
        if (expanded.Count > maximumStops) stopLimitExceeded = true;
        return expanded;
    }

    private static OfficeColor SampleRepeating(IReadOnlyList<ResolvedStop> source, double period, double offset) {
        double local = source[0].Offset + PositiveModulo(offset - source[0].Offset, period);
        return Sample(source, local);
    }

    private static double PositiveModulo(double value, double divisor) {
        double result = value % divisor;
        return result < 0D ? result + divisor : result;
    }

    private static OfficeColor Sample(IReadOnlyList<ResolvedStop> source, double offset) {
        if (offset <= source[0].Offset) return source[0].Color;
        for (int index = 1; index < source.Count; index++) {
            ResolvedStop current = source[index];
            if (offset > current.Offset) continue;
            ResolvedStop previous = source[index - 1];
            if (current.Offset <= previous.Offset) return current.Color;
            double ratio = (offset - previous.Offset) / (current.Offset - previous.Offset);
            return InterpolatePremultiplied(previous.Color, current.Color, ratio);
        }
        return source[source.Count - 1].Color;
    }

    private static OfficeColor InterpolatePremultiplied(OfficeColor first, OfficeColor second, double ratio) {
        double firstAlpha = first.A / 255D;
        double secondAlpha = second.A / 255D;
        double alpha = firstAlpha + ((secondAlpha - firstAlpha) * ratio);
        if (alpha <= 0.000001D) return OfficeColor.FromRgba(0, 0, 0, 0);
        return OfficeColor.FromRgba(
            ToByte(((first.R * firstAlpha) + (((second.R * secondAlpha) - (first.R * firstAlpha)) * ratio)) / alpha),
            ToByte(((first.G * firstAlpha) + (((second.G * secondAlpha) - (first.G * firstAlpha)) * ratio)) / alpha),
            ToByte(((first.B * firstAlpha) + (((second.B * secondAlpha) - (first.B * firstAlpha)) * ratio)) / alpha),
            Interpolate(first.A, second.A, ratio));
    }

    private static byte ToByte(double value) =>
        (byte)Math.Round(Math.Max(0D, Math.Min(255D, value)), MidpointRounding.AwayFromZero);

    private static byte Interpolate(byte first, byte second, double ratio) =>
        (byte)Math.Round(first + ((second - first) * ratio), MidpointRounding.AwayFromZero);

    private readonly struct ResolvedStop {
        internal ResolvedStop(double offset, OfficeColor color) {
            Offset = offset;
            Color = color;
        }
        internal double Offset { get; }
        internal OfficeColor Color { get; }
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
