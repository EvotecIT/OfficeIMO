using OfficeIMO.Drawing;
using Sample = OfficeIMO.Pdf.PdfVisualResourceDictionaryBuilder.TransformedGradientSample;

namespace OfficeIMO.Pdf;

/// <summary>One-dimensional color samples, separate from periodic field geometry and alpha.</summary>
internal sealed class PdfRadialSpreadColors {
    private PdfRadialSpreadColors(IReadOnlyList<Sample> samples, double[] outside, double[] common) {
        Samples = samples;
        Outside = outside;
        Common = common;
    }

    internal IReadOnlyList<Sample> Samples { get; }
    internal double[] Outside { get; }
    internal double[] Common { get; }
    internal int ComponentCount => Outside.Length;

    internal static PdfRadialSpreadColors Create(IReadOnlyList<OfficeGradientStop> stops,
        OfficeGradientColorInterpolation interpolation, OfficeColor? outside, bool alphaOnly,
        PdfPrintColorTransform? printTransform = null) {
        double[] Components(OfficeColor color) {
            if (alphaOnly) return new[] { color.A / 255D };
            if (printTransform != null) {
                var converted = new double[4];
                printTransform.Convert(color, converted);
                return converted;
            }
            double[] values = { color.R / 255D, color.G / 255D, color.B / 255D };
            if (interpolation == OfficeGradientColorInterpolation.LinearRgb)
                for (int index = 0; index < values.Length; index++) values[index] = OfficeColorSpaceConverter.FromSrgb(values[index]);
            return values;
        }

        IReadOnlyList<Sample> samples = printTransform != null && !alphaOnly
            ? PdfVisualResourceDictionaryBuilder.CreateTransformedGradientSamples(stops, printTransform, interpolation)
            : stops.Select(stop => new Sample(stop.Offset, Components(stop.Color))).ToArray();
        return new PdfRadialSpreadColors(samples, Components(outside ?? OfficeColor.Transparent),
            Components(outside ?? stops[stops.Count - 1].Color));
    }
}
