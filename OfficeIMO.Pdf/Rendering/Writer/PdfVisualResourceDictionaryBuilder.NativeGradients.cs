using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfVisualResourceDictionaryBuilder {
    /// <summary>
    /// Reports whether native shading can carry these stops, including the bounded intermediate
    /// samples that approximate premultiplied-alpha interpolation.
    /// </summary>
    /// <remarks>
    /// The checks match the ones the writer applies when it builds the shading function, so a
    /// producer can choose another representation before the document is serialized.
    /// </remarks>
    internal static bool CanWriteGradient(IReadOnlyList<OfficeGradientStop>? stops,
        OfficeGradientColorInterpolation colorInterpolation, bool interpolateAlphaSeparately) {
        if (stops == null) return true;
        try {
            ValidateStops(stops);
            IReadOnlyList<OfficeGradientStop> normalized = HasDuplicateOffsets(stops) ? NormalizeGradientStops(stops) : stops;
            if (!interpolateAlphaSeparately && HasVaryingGradientAlpha(normalized)) {
                CreateTransformedGradientSamples(normalized, printColorTransform: null, colorInterpolation, interpolateAlphaSeparately);
            }
            return true;
        } catch (ArgumentException) {
            return false;
        } catch (InvalidOperationException) {
            return false;
        }
    }
}
