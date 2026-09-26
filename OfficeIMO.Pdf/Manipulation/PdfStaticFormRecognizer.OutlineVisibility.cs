using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfStaticFormRecognizer {
    private static bool HasVisibleOutline(PdfPageVisualPrimitive primitive, VisualRect visual,
        PdfStaticFormEvidenceKind evidence, IReadOnlyList<PaintArea> filledAreas,
        CancellationToken cancellationToken) {
        if (primitive.StrokeTilingPattern is not null) return false;
        VisualRect paintBounds = evidence == PdfStaticFormEvidenceKind.OutlinedField
            ? visual : OutlinePaintBounds(primitive, visual);
        IReadOnlyList<OfficeGradientStop>? stops = primitive.StrokeGradient?.Stops ?? primitive.StrokeRadialGradient?.Stops;
        if (stops is not null) {
            if (stops.Count == 0) return false;
            for (int index = 0; index < stops.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeColor start = stops[index].Color;
                OfficeColor end = stops[Math.Min(index + 1, stops.Count - 1)].Color;
                if (IsOpaqueWhiteFill(primitive) && primitive.FillColor is OfficeColor ownFill) {
                    if (!ColorRangeContrasts(start, end, ownFill)) return false;
                } else if (!HasContrastingBackdrop(filledAreas, paintBounds, start,
                    primitive.PaintOrder, primitive.ContentOrderKey, outlinedStroke: true, gradientEnd: end)) return false;
            }
            return true;
        }
        if (primitive.StrokeColor is not OfficeColor ink) return false;
        return IsOpaqueWhiteFill(primitive) && primitive.FillColor is OfficeColor fill
            ? ColorsContrast(ink, fill)
            : HasContrastingBackdrop(filledAreas, paintBounds, ink,
                  primitive.PaintOrder, primitive.ContentOrderKey, outlinedStroke: true);
    }

    private static bool OutlineContrastsColor(PdfPageVisualPrimitive outline, OfficeColor backdrop) {
        if (outline.StrokeTilingPattern is not null) return false;
        IReadOnlyList<OfficeGradientStop>? stops = outline.StrokeGradient?.Stops ?? outline.StrokeRadialGradient?.Stops;
        if (stops is null) return outline.StrokeColor is OfficeColor ink && ColorsContrast(ink, backdrop);
        if (stops.Count == 0) return false;
        for (int index = 0; index < stops.Count; index++) {
            if (!ColorRangeContrasts(stops[index].Color,
                stops[Math.Min(index + 1, stops.Count - 1)].Color, backdrop)) return false;
        }
        return true;
    }

    // A separating RGB channel proves contrast over the entire interpolated segment.
    // Endpoint checks alone can miss a gradient passing through the backdrop color.
    private static bool ColorRangeContrasts(OfficeColor start, OfficeColor end, OfficeColor backdrop) =>
        ChannelRangeContrasts(start.R, end.R, backdrop.R) ||
        ChannelRangeContrasts(start.G, end.G, backdrop.G) ||
        ChannelRangeContrasts(start.B, end.B, backdrop.B);

    private static bool ChannelRangeContrasts(byte start, byte end, byte backdrop) =>
        Math.Min(start, end) - backdrop >= 45 || backdrop - Math.Max(start, end) >= 45;
}
