using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfVisualResourceDictionaryBuilder {
    private const int MaximumGradientStops = 1024;
    private const int MaximumTransformedGradientSamples = 4096;
    private const int MinimumGradientSubdivisionDepth = 2;
    private const int MaximumGradientSubdivisionDepth = 8;
    private const double GradientTransformTolerance = 1D / 1024D;

    internal static string BuildExtGStateObject(
        double fillOpacity,
        double strokeOpacity,
        OfficeBlendMode blendMode = OfficeBlendMode.Normal) {
        ValidateOpacity(fillOpacity, nameof(fillOpacity));
        ValidateOpacity(strokeOpacity, nameof(strokeOpacity));
        if (blendMode < OfficeBlendMode.Normal || blendMode > OfficeBlendMode.Luminosity) {
            throw new ArgumentOutOfRangeException(nameof(blendMode), blendMode, "Unsupported PDF blend mode.");
        }

        return "<< /Type /ExtGState /ca " +
            FormatNumber(fillOpacity) +
            " /CA " +
            FormatNumber(strokeOpacity) +
            (blendMode == OfficeBlendMode.Normal ? string.Empty : " /BM /" + blendMode) +
            " >>\n";
    }

    internal static string BuildAxialShadingObject(
        double x0,
        double y0,
        double x1,
        double y1,
        OfficeColor startColor,
        OfficeColor endColor) => BuildAxialShadingObject(
            x0,
            y0,
            x1,
            y1,
            new[] { new OfficeGradientStop(0D, startColor), new OfficeGradientStop(1D, endColor) });

    internal static string BuildAxialShadingObject(
        double x0,
        double y0,
        double x1,
        double y1,
        IReadOnlyList<OfficeGradientStop> stops,
        PdfPrintColorTransform? printColorTransform = null) {
        ValidateFinite(x0, nameof(x0));
        ValidateFinite(y0, nameof(y0));
        ValidateFinite(x1, nameof(x1));
        ValidateFinite(y1, nameof(y1));
        ValidateStops(stops);

        return
            "<< /ShadingType 2 /ColorSpace " + (printColorTransform == null ? "/DeviceRGB" : "/DeviceCMYK") + " /Coords [" +
            FormatNumber(x0) + " " + FormatNumber(y0) + " " + FormatNumber(x1) + " " + FormatNumber(y1) +
            "] /Function " + BuildGradientFunction(stops, printColorTransform) + " /Extend [true true] >>\n";
    }

    internal static string BuildRadialShadingObject(
        double x0,
        double y0,
        double r0,
        double x1,
        double y1,
        double r1,
        IReadOnlyList<OfficeGradientStop> stops,
        PdfPrintColorTransform? printColorTransform = null) {
        ValidateFinite(x0, nameof(x0));
        ValidateFinite(y0, nameof(y0));
        ValidateRadius(r0, nameof(r0));
        ValidateFinite(x1, nameof(x1));
        ValidateFinite(y1, nameof(y1));
        ValidateRadius(r1, nameof(r1));
        ValidateStops(stops);
        if (x0.Equals(x1) && y0.Equals(y1) && r0.Equals(r1)) {
            throw new ArgumentException("Radial PDF shading circles must be different.", nameof(r1));
        }

        return
            "<< /ShadingType 3 /ColorSpace " + (printColorTransform == null ? "/DeviceRGB" : "/DeviceCMYK") + " /Coords [" +
            FormatNumber(x0) + " " + FormatNumber(y0) + " " + FormatNumber(r0) + " " +
            FormatNumber(x1) + " " + FormatNumber(y1) + " " + FormatNumber(r1) +
            "] /Function " + BuildGradientFunction(stops, printColorTransform) + " /Extend [true true] >>\n";
    }

    private static string BuildGradientFunction(IReadOnlyList<OfficeGradientStop> stops, PdfPrintColorTransform? printColorTransform) {
        IReadOnlyList<OfficeGradientStop> normalized = HasDuplicateOffsets(stops)
            ? NormalizeGradientStops(stops)
            : stops;
        if (printColorTransform != null) {
            return BuildTransformedGradientFunction(normalized, printColorTransform);
        }
        if (normalized.Count == 2) return BuildInterpolationFunction(normalized[0].Color, normalized[1].Color, printColorTransform);

        var builder = new System.Text.StringBuilder("<< /FunctionType 3 /Domain [0 1] /Functions [");
        for (int index = 1; index < normalized.Count; index++) {
            if (index > 1) builder.Append(' ');
            builder.Append(BuildInterpolationFunction(normalized[index - 1].Color, normalized[index].Color, printColorTransform));
        }

        builder.Append("] /Bounds [");
        for (int index = 1; index < normalized.Count - 1; index++) {
            if (index > 1) builder.Append(' ');
            builder.Append(FormatGradientOffset(normalized[index].Offset));
        }

        builder.Append("] /Encode [");
        for (int index = 1; index < normalized.Count; index++) {
            if (index > 1) builder.Append(' ');
            builder.Append("0 1");
        }

        return builder.Append("] >>").ToString();
    }

    private static string BuildTransformedGradientFunction(
        IReadOnlyList<OfficeGradientStop> stops,
        PdfPrintColorTransform printColorTransform) {
        var samples = new List<TransformedGradientSample>();
        for (int index = 1; index < stops.Count; index++) {
            OfficeGradientStop start = stops[index - 1];
            OfficeGradientStop end = stops[index];
            var startComponents = new double[4];
            var endComponents = new double[4];
            printColorTransform.Convert(start.Color, startComponents);
            printColorTransform.Convert(end.Color, endComponents);
            if (samples.Count == 0) samples.Add(new TransformedGradientSample(start.Offset, startComponents));
            AppendAdaptiveGradientSamples(
                start,
                end,
                0D,
                1D,
                startComponents,
                endComponents,
                depth: 0,
                printColorTransform,
                samples);
        }

        if (samples.Count == 2) {
            return BuildCmykInterpolationFunction(samples[0].Components, samples[1].Components);
        }

        var builder = new System.Text.StringBuilder("<< /FunctionType 3 /Domain [0 1] /Functions [");
        for (int index = 1; index < samples.Count; index++) {
            if (index > 1) builder.Append(' ');
            builder.Append(BuildCmykInterpolationFunction(samples[index - 1].Components, samples[index].Components));
        }

        builder.Append("] /Bounds [");
        for (int index = 1; index < samples.Count - 1; index++) {
            if (index > 1) builder.Append(' ');
            builder.Append(FormatGradientOffset(samples[index].Offset));
        }

        builder.Append("] /Encode [");
        for (int index = 1; index < samples.Count; index++) {
            if (index > 1) builder.Append(' ');
            builder.Append("0 1");
        }

        return builder.Append("] >>").ToString();
    }

    private static void AppendAdaptiveGradientSamples(
        OfficeGradientStop intervalStart,
        OfficeGradientStop intervalEnd,
        double startPosition,
        double endPosition,
        double[] startComponents,
        double[] endComponents,
        int depth,
        PdfPrintColorTransform printColorTransform,
        List<TransformedGradientSample> samples) {
        double middlePosition = (startPosition + endPosition) / 2D;
        OfficeColor middleColor = InterpolateColor(intervalStart.Color, intervalEnd.Color, middlePosition);
        var middleComponents = new double[4];
        printColorTransform.Convert(middleColor, middleComponents);

        bool needsSubdivision = depth < MinimumGradientSubdivisionDepth ||
            (depth < MaximumGradientSubdivisionDepth &&
             MaximumMidpointError(startComponents, middleComponents, endComponents) > GradientTransformTolerance);
        if (!needsSubdivision) {
            AddTransformedGradientSample(
                samples,
                InterpolateOffset(intervalStart.Offset, intervalEnd.Offset, endPosition),
                endComponents);
            return;
        }

        AppendAdaptiveGradientSamples(
            intervalStart,
            intervalEnd,
            startPosition,
            middlePosition,
            startComponents,
            middleComponents,
            depth + 1,
            printColorTransform,
            samples);
        AppendAdaptiveGradientSamples(
            intervalStart,
            intervalEnd,
            middlePosition,
            endPosition,
            middleComponents,
            endComponents,
            depth + 1,
            printColorTransform,
            samples);
    }

    private static void AddTransformedGradientSample(
        List<TransformedGradientSample> samples,
        double offset,
        double[] components) {
        if (samples.Count >= MaximumTransformedGradientSamples) {
            throw new InvalidOperationException("PDF gradient color conversion exceeded the bounded sample count.");
        }

        samples.Add(new TransformedGradientSample(offset, components));
    }

    private static double MaximumMidpointError(double[] start, double[] middle, double[] end) {
        double maximum = 0D;
        for (int index = 0; index < middle.Length; index++) {
            maximum = Math.Max(maximum, Math.Abs(middle[index] - ((start[index] + end[index]) / 2D)));
        }
        return maximum;
    }

    private static OfficeColor InterpolateColor(OfficeColor start, OfficeColor end, double position) =>
        OfficeColor.FromRgb(
            InterpolateByte(start.R, end.R, position),
            InterpolateByte(start.G, end.G, position),
            InterpolateByte(start.B, end.B, position));

    private static byte InterpolateByte(byte start, byte end, double position) =>
        (byte)Math.Round(start + ((end - start) * position), MidpointRounding.AwayFromZero);

    private static double InterpolateOffset(double start, double end, double position) =>
        start + ((end - start) * position);

    private static bool HasDuplicateOffsets(IReadOnlyList<OfficeGradientStop> stops) {
        for (int index = 1; index < stops.Count; index++) {
            if (stops[index].Offset.Equals(stops[index - 1].Offset)) return true;
        }
        return false;
    }

    private static List<OfficeGradientStop> NormalizeGradientStops(IReadOnlyList<OfficeGradientStop> stops) {
        var normalized = new List<OfficeGradientStop>(stops.Count);
        int index = 0;
        while (index < stops.Count) {
            int end = index + 1;
            while (end < stops.Count && stops[end].Offset.Equals(stops[index].Offset)) end++;
            if (end - index == 1) {
                normalized.Add(stops[index]);
            } else if (stops[index].Offset <= 0D) {
                normalized.Add(new OfficeGradientStop(0D, stops[end - 1].Color));
            } else if (stops[index].Offset >= 1D) {
                normalized.Add(new OfficeGradientStop(1D, stops[index].Color));
            } else {
                double previousOffset = stops[index - 1].Offset;
                double nextOffset = stops[end].Offset;
                double epsilon = Math.Min(0.0000001D, Math.Min(stops[index].Offset - previousOffset, nextOffset - stops[index].Offset) / 4D);
                normalized.Add(new OfficeGradientStop(stops[index].Offset - epsilon, stops[index].Color));
                normalized.Add(new OfficeGradientStop(stops[index].Offset + epsilon, stops[end - 1].Color));
            }
            index = end;
        }
        return normalized;
    }

    private static string BuildInterpolationFunction(OfficeColor startColor, OfficeColor endColor, PdfPrintColorTransform? printColorTransform) {
        if (printColorTransform == null) {
            return "<< /FunctionType 2 /Domain [0 1] /C0 [" +
                FormatColorComponent(startColor.R) + " " + FormatColorComponent(startColor.G) + " " + FormatColorComponent(startColor.B) +
                "] /C1 [" +
                FormatColorComponent(endColor.R) + " " + FormatColorComponent(endColor.G) + " " + FormatColorComponent(endColor.B) +
                "] /N 1 >>";
        }

        var start = new double[4];
        var end = new double[4];
        printColorTransform.Convert(startColor, start);
        printColorTransform.Convert(endColor, end);
        return BuildCmykInterpolationFunction(start, end);
    }

    private static string BuildCmykInterpolationFunction(double[] start, double[] end) =>
        "<< /FunctionType 2 /Domain [0 1] /C0 [" + FormatComponents(start) +
        "] /C1 [" + FormatComponents(end) + "] /N 1 >>";

    private static string FormatComponents(double[] components) =>
        string.Join(" ", components.Select(static component => FormatNumber(component)));

    private static void ValidateStops(IReadOnlyList<OfficeGradientStop>? stops) {
        if (stops == null || stops.Count < 2) throw new ArgumentException("A PDF shading needs at least two stops.", nameof(stops));
        if (stops.Count > MaximumGradientStops) {
            throw new ArgumentException("A PDF shading exceeds the bounded stop count.", nameof(stops));
        }
        if (!stops[0].Offset.Equals(0D) || !stops[stops.Count - 1].Offset.Equals(1D)) {
            throw new ArgumentException("PDF shading stops must start at zero and end at one.", nameof(stops));
        }

        double previous = -1D;
        for (int index = 0; index < stops.Count; index++) {
            double offset = stops[index].Offset;
            if (double.IsNaN(offset) || double.IsInfinity(offset) || offset < previous) {
                throw new ArgumentException("PDF shading stops must use non-decreasing finite offsets.", nameof(stops));
            }

            previous = offset;
        }
    }

    private static string FormatColorComponent(byte value) =>
        FormatNumber(value / 255D);

    private static string FormatGradientOffset(double value) =>
        value.ToString("0.########", CultureInfo.InvariantCulture);

    private static string FormatNumber(double value) =>
        value.ToString("0.###", CultureInfo.InvariantCulture);

    private static void ValidateOpacity(double value, string paramName) {
        ValidateFinite(value, paramName);
        if (value < 0 || value > 1) {
            throw new ArgumentOutOfRangeException(paramName, value, "PDF graphics-state opacity must be between 0 and 1.");
        }
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, value, "PDF visual resource numbers must be finite.");
        }
    }

    private static void ValidateRadius(double value, string paramName) {
        ValidateFinite(value, paramName);
        if (value < 0D) throw new ArgumentOutOfRangeException(paramName, value, "PDF radial shading radii must be non-negative.");
    }

    private readonly struct TransformedGradientSample {
        internal TransformedGradientSample(double offset, double[] components) {
            Offset = offset;
            Components = components;
        }

        internal double Offset { get; }
        internal double[] Components { get; }
    }
}
