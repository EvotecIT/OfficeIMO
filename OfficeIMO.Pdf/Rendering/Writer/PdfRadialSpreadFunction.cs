using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Emits a two-input calculator function for a periodic two-circle field.</summary>
internal static class PdfRadialSpreadFunction {
    internal static string Build(double x0, double y0, double r0, double x1, double y1, double r1,
        IReadOnlyList<OfficeGradientStop> stops, OfficeGradientSpreadMode spread,
        OfficeGradientColorInterpolation interpolation, OfficeColor? outside, bool nativeSeam, bool alphaOnly, int colorChannel = -1) {
        string N(double value) => PdfNumberFormatter.Precise(value);
        bool reverse = r1 == 0D;
        double radius = Math.Max(r0, r1);
        if (!(radius > 0D)) throw new NotSupportedException("Periodic radial shading requires a positive circle radius.");
        double ox = reverse ? x1 : x0, oy = reverse ? y1 : y0;
        double dx = (reverse ? x0 - x1 : x1 - x0) / radius;
        double dy = (reverse ? y0 - y1 : y1 - y0) / radius;
        double rs = reverse ? 0D : r0 / radius;
        double dr = (reverse ? r0 : r1 - r0) / radius;
        double a = dx * dx + dy * dy - dr * dr;
        var code = new StringBuilder("{ ");
        // Input x,y becomes the quadratic coefficients b,c after normalization.
        code.Append(N(oy)).Append(" sub ").Append(N(radius)).Append(" div exch ")
            .Append(N(ox)).Append(" sub ").Append(N(radius)).Append(" div exch ")
            .Append("2 copy ").Append(N(dy)).Append(" mul exch ").Append(N(dx)).Append(" mul add ")
            .Append(N(rs * dr)).Append(" add -2 mul 3 1 roll dup mul exch dup mul add ")
            .Append(N(rs * rs)).Append(" sub ");
        if (a == 0D) {
            // Result is ratio,valid. The common tangent point has a limiting endpoint.
            code.Append("1 index 0 eq { exch pop dup 0 eq { pop ")
                .Append(N(reverse ? 0D : dr < 0D ? -rs / dr : 1D))
                .Append(" true } { pop 0 false } ifelse } { neg exch div dup ")
                .Append(N(dr)).Append(" mul ").Append(N(rs)).Append(" add 0 ge } ifelse ");
        } else {
            code.Append("1 index dup mul 1 index ").Append(N(4D * a)).Append(" mul sub dup 0 lt ")
                .Append("{ pop pop pop 0 false } { sqrt 2 index 0 ge ")
                .Append("{ 2 index add -0.5 mul } { 2 index exch sub -0.5 mul } ifelse ")
                .Append("3 -1 roll pop dup 0 eq { pop pop 0 0 } { dup ").Append(N(a))
                .Append(" div 3 1 roll div } ifelse ");
            string valid = N(dr) + " mul " + N(rs) + " add 0 ge ";
            code.Append("1 index ").Append(valid).Append("{ dup ").Append(valid)
                .Append("{ 2 copy ").Append(reverse ? "lt" : "gt")
                .Append(" { pop } { exch pop } ifelse } { pop } ifelse true } { dup ")
                .Append(valid).Append("{ exch pop true } { pop pop 0 false } ifelse } ifelse } ifelse ");
        }
        code.Append("{ ");
        if (reverse) code.Append("1 exch sub ");
        if (spread == OfficeGradientSpreadMode.Repeat) {
            code.Append("dup floor sub ");
            if (nativeSeam) code.Append("dup 0 eq { pop 1 } if ");
        } else code.Append("dup 2 div floor 2 mul sub dup 1 gt { 2 exch sub } if ");
        int channels = alphaOnly || colorChannel >= 0 ? 1 : 3;
        for (int channel = 0; channel < channels; channel++) {
            code.Append(channel).Append(" index ");
            AppendStops(code, stops, 0, stops.Count - 1, colorChannel >= 0 ? colorChannel : channel, alphaOnly, interpolation);
        }
        code.Append(channels + 1).Append(" -1 roll pop } { pop ");
        for (int channel = 0; channel < channels; channel++)
            code.Append(N(Component(outside ?? OfficeColor.Transparent, colorChannel >= 0 ? colorChannel : channel, alphaOnly, interpolation))).Append(' ');
        string result = code.Append("} ifelse }").ToString();
        if (a == 0D && rs == 0D && !reverse) {
            // The common tangent point has t=+infinity: Pad's endpoint rule
            // applies before periodic mapping, or SVG's boundary average.
            OfficeColor common = outside ?? stops[stops.Count - 1].Color;
            string components = string.Join(" ", Enumerable.Range(0, channels)
                .Select(channel => N(Component(common, colorChannel >= 0 ? colorChannel : channel, alphaOnly, interpolation))));
            result = "{ 2 copy " + N(y0) + " eq exch " + N(x0) + " eq and { pop pop " + components +
                " } " + result + " ifelse }";
        }
        return result;
    }

    private static void AppendStops(StringBuilder code, IReadOnlyList<OfficeGradientStop> stops,
        int first, int last, int channel, bool alpha, OfficeGradientColorInterpolation interpolation, int depth = 0) {
        string N(double value) => PdfNumberFormatter.Precise(value);
        if (last - first > 1 && depth < 5) {
            int middle = (first + last) / 2;
            code.Append("dup ").Append(N(stops[middle].Offset)).Append(" lt { ");
            AppendStops(code, stops, first, middle, channel, alpha, interpolation, depth + 1);
            code.Append("} { ");
            AppendStops(code, stops, middle, last, channel, alpha, interpolation, depth + 1);
            code.Append("} ifelse ");
            return;
        }
        if (last - first > 1) {
            // Keep calculator procedure nesting within ten levels, including
            // field validity and endpoint handling. Each bounded leaf bucket
            // selects among at most 32 segments without nested conditionals.
            code.Append("dup ");
            AppendSegment(code, stops, last - 1, last, channel, alpha, interpolation);
            for (int index = last - 2; index >= first; index--) {
                code.Append("1 index ").Append(N(stops[index + 1].Offset)).Append(" lt { pop dup ");
                AppendSegment(code, stops, index, index + 1, channel, alpha, interpolation);
                code.Append("} if ");
            }
            code.Append("exch pop ");
            return;
        }
        AppendSegment(code, stops, first, last, channel, alpha, interpolation);
    }

    private static void AppendSegment(StringBuilder code, IReadOnlyList<OfficeGradientStop> stops,
        int first, int last, int channel, bool alpha, OfficeGradientColorInterpolation interpolation) {
        string N(double value) => PdfNumberFormatter.Precise(value);
        double start = Component(stops[first].Color, channel, alpha, interpolation);
        double end = Component(stops[last].Color, channel, alpha, interpolation);
        double width = stops[last].Offset - stops[first].Offset;
        if (width == 0D) { code.Append("pop ").Append(N(end)).Append(' '); return; }
        // Periodic mapping and segment selection already bound the final ratio.
        code.Append(N(stops[first].Offset)).Append(" sub ").Append(N(width)).Append(" div ")
            .Append(N(end - start)).Append(" mul ").Append(N(start)).Append(" add ");
    }

    private static double Component(OfficeColor color, int channel, bool alpha, OfficeGradientColorInterpolation interpolation) {
        if (alpha) return color.A / 255D;
        double value = (channel == 0 ? color.R : channel == 1 ? color.G : color.B) / 255D;
        return interpolation == OfficeGradientColorInterpolation.LinearRgb ? OfficeColorSpaceConverter.FromSrgb(value) : value;
    }
}
