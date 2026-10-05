using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Emits a two-input calculator function for a periodic two-circle field.</summary>
internal static class PdfRadialSpreadFunction {
    internal static string Build(double x0, double y0, double r0, double x1, double y1, double r1,
        IReadOnlyList<OfficeGradientStop> stops, OfficeGradientSpreadMode spread,
        OfficeGradientColorInterpolation interpolation, OfficeColor? outside, bool nativeSeam, bool alphaOnly, int colorChannel = -1, PdfRadialSpreadColors? colors = null) {
        colors ??= PdfRadialSpreadColors.Create(stops, interpolation, outside, alphaOnly);
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
        bool hasCommonPoint = a == 0D && rs == 0D && !reverse;
        var code = new StringBuilder("{ ");
        if (hasCommonPoint) code.Append("2 copy ").Append(N(y0)).Append(" eq exch ").Append(N(x0))
            .Append(" eq and 3 1 roll ");
        // Input x,y becomes the quadratic coefficients b,c after normalization.
        code.Append(N(oy)).Append(" sub ").Append(N(radius)).Append(" div exch ")
            .Append(N(ox)).Append(" sub ").Append(N(radius)).Append(" div exch ")
            .Append("2 copy ").Append(N(dy)).Append(" mul exch ").Append(N(dx)).Append(" mul add ")
            .Append(N(rs * dr)).Append(" add -2 mul 3 1 roll dup mul exch dup mul add ")
            .Append(N(rs * rs)).Append(" sub ");
        // Comparisons produce boolean constants without literal tokens rejected by some PDF consumers.
        if (a == 0D) {
            // Result is ratio,valid. The common tangent point has a limiting endpoint.
            code.Append("1 index 0 eq { exch pop dup 0 eq { pop ")
                .Append(N(reverse ? 0D : dr < 0D ? -rs / dr : 1D))
                .Append(" 0 0 eq } { pop 0 0 1 eq } ifelse } { neg exch div dup ")
                .Append(N(dr)).Append(" mul ").Append(N(rs)).Append(" add 0 ge } ifelse ");
        } else {
            code.Append("1 index dup mul 1 index ").Append(N(4D * a)).Append(" mul sub dup 0 lt ")
                .Append("{ pop pop pop 0 0 1 eq } { sqrt 2 index 0 ge ")
                .Append("{ 2 index add -0.5 mul } { 2 index exch sub -0.5 mul } ifelse ")
                .Append("3 -1 roll pop dup 0 eq { pop pop 0 0 } { dup ").Append(N(a))
                .Append(" div 3 1 roll div } ifelse ");
            string valid = N(dr) + " mul " + N(rs) + " add 0 ge ";
            code.Append("1 index ").Append(valid).Append("{ dup ").Append(valid)
                .Append("{ 2 copy ").Append(reverse ? "lt" : "gt")
                .Append(" { pop } { exch pop } ifelse } { pop } ifelse 0 0 eq } { dup ")
                .Append(valid).Append("{ exch pop 0 0 eq } { pop pop 0 0 1 eq } ifelse } ifelse } ifelse ");
        }
        code.Append("exch ");
        if (reverse) code.Append("1 exch sub ");
        if (spread == OfficeGradientSpreadMode.Repeat) {
            code.Append("dup floor sub ");
            if (nativeSeam) code.Append("dup 0 eq { pop 1 } if ");
        } else code.Append("dup 2 div floor 2 mul sub dup 1 gt { 2 exch sub } if ");
        int channels = colorChannel >= 0 ? 1 : colors.ComponentCount;
        for (int channel = 0; channel < channels; channel++) {
            code.Append(channel).Append(" index ");
            AppendColor(code, colors, colorChannel >= 0 ? colorChannel : channel);
        }
        // Explicit boolean comparison avoids calculator consumers treating not as integer complement.
        code.Append(channels + 1).Append(" -1 roll pop ").Append(channels + 1).Append(" -1 roll 0 1 eq eq { ");
        for (int channel = 0; channel < channels; channel++) code.Append("pop ");
        for (int channel = 0; channel < channels; channel++)
            code.Append(N(colors.Outside[colorChannel >= 0 ? colorChannel : channel])).Append(' ');
        code.Append("} if ");
        if (hasCommonPoint) {
            code.Append(channels + 1).Append(" -1 roll { ");
            for (int channel = 0; channel < channels; channel++) code.Append("pop ");
            for (int channel = 0; channel < channels; channel++)
                code.Append(N(colors.Common[colorChannel >= 0 ? colorChannel : channel])).Append(' ');
            code.Append("} if ");
        }
        return code.Append('}').ToString();
    }

    private static void AppendColor(StringBuilder code, PdfRadialSpreadColors colors, int channel) {
        var stops = colors.Samples;
        if (stops.Count <= 1024) {
            AppendStops(code, colors, 0, stops.Count - 1, channel);
            return;
        }
        // Each independent range keeps branch bytecode below portable jump
        // limits. At most four ranges execute one bounded lookup each.
        code.Append(PdfNumberFormatter.Precise(stops[stops.Count - 1].Components[channel])).Append(' ');
        for (int first = 0; first < stops.Count - 1; first += 1024) {
            int last = Math.Min(first + 1024, stops.Count - 1);
            code.Append("1 index ").Append(PdfNumberFormatter.Precise(stops[first].Offset))
                .Append(" ge 2 index ").Append(PdfNumberFormatter.Precise(stops[last].Offset))
                .Append(" lt and { pop dup ");
            AppendStops(code, colors, first, last, channel);
            code.Append("} if ");
        }
        code.Append("exch pop ");
    }

    private static void AppendStops(StringBuilder code, PdfRadialSpreadColors colors,
        int first, int last, int channel, int depth = 0) {
        var stops = colors.Samples;
        string N(double value) => PdfNumberFormatter.Precise(value);
        if (last - first > 1 && depth < 7) {
            int middle = (first + last) / 2;
            code.Append("dup ").Append(N(stops[middle].Offset)).Append(" lt { ");
            AppendStops(code, colors, first, middle, channel, depth + 1);
            code.Append("} { ");
            AppendStops(code, colors, middle, last, channel, depth + 1);
            code.Append("} ifelse ");
            return;
        }
        if (last - first > 1) {
            // Keep calculator procedure nesting within ten levels, including
            // the independent range selector. Each bounded leaf bucket
            // selects among at most eight segments without nested conditionals.
            code.Append("dup ");
            AppendSegment(code, colors, last - 1, last, channel);
            for (int index = last - 2; index >= first; index--) {
                code.Append("1 index ").Append(N(stops[index + 1].Offset)).Append(" lt { pop dup ");
                AppendSegment(code, colors, index, index + 1, channel);
                code.Append("} if ");
            }
            code.Append("exch pop ");
            return;
        }
        AppendSegment(code, colors, first, last, channel);
    }

    private static void AppendSegment(StringBuilder code, PdfRadialSpreadColors colors,
        int first, int last, int channel) {
        string N(double value) => PdfNumberFormatter.Precise(value);
        var stops = colors.Samples;
        double start = stops[first].Components[channel];
        double end = stops[last].Components[channel];
        double width = stops[last].Offset - stops[first].Offset;
        if (width == 0D) { code.Append("pop ").Append(N(end)).Append(' '); return; }
        // Periodic mapping and segment selection already bound the final ratio.
        code.Append(N(stops[first].Offset)).Append(" sub ").Append(N(width)).Append(" div ")
            .Append(N(end - start)).Append(" mul ").Append(N(start)).Append(" add ");
    }

}
