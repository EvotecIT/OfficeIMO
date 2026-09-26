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
                OfficeColor start = ApplyOutlineOpacity(stops[index].Color, primitive);
                OfficeColor end = ApplyOutlineOpacity(stops[Math.Min(index + 1, stops.Count - 1)].Color, primitive);
                if (IsOpaqueWhiteFill(primitive) && primitive.FillColor is OfficeColor ownFill) {
                    if (!ColorRangeContrasts(start, end, ownFill)) return false;
                } else if (!HasContrastingBackdrop(filledAreas, paintBounds, start,
                    primitive.PaintOrder, primitive.ContentOrderKey, outlinedStroke: true, gradientEnd: end)) return false;
            }
            return true;
        }
        if (primitive.StrokeColor is not OfficeColor ink) return false;
        ink = ApplyOutlineOpacity(ink, primitive);
        return IsOpaqueWhiteFill(primitive) && primitive.FillColor is OfficeColor fill
            ? ColorsContrast(ink, fill)
            : HasContrastingBackdrop(filledAreas, paintBounds, ink,
                  primitive.PaintOrder, primitive.ContentOrderKey, outlinedStroke: true);
    }

    private static bool OutlineContrastsColor(PdfPageVisualPrimitive outline, OfficeColor backdrop) {
        if (outline.StrokeTilingPattern is not null) return false;
        IReadOnlyList<OfficeGradientStop>? stops = outline.StrokeGradient?.Stops ?? outline.StrokeRadialGradient?.Stops;
        if (stops is null) return outline.StrokeColor is OfficeColor ink &&
            ColorsContrast(ApplyOutlineOpacity(ink, outline), backdrop);
        if (stops.Count == 0) return false;
        for (int index = 0; index < stops.Count; index++) {
            if (!ColorRangeContrasts(ApplyOutlineOpacity(stops[index].Color, outline),
                ApplyOutlineOpacity(stops[Math.Min(index + 1, stops.Count - 1)].Color, outline), backdrop)) return false;
        }
        return true;
    }

    // A separating RGB channel proves contrast over the entire interpolated segment.
    // Endpoint checks alone can miss a gradient passing through the backdrop color.
    private static OfficeColor ApplyOutlineOpacity(OfficeColor color, PdfPageVisualPrimitive outline) =>
        new OfficeColor(color.R, color.G, color.B,
            (byte)Math.Floor(color.A * Math.Min(1D, Math.Max(0D, outline.StrokeOpacity ?? 1D))));

    private static bool ColorRangeContrasts(OfficeColor start, OfficeColor end, OfficeColor backdrop) =>
        Math.Max(ChannelRangeContrast(start.R, end.R, backdrop.R),
            Math.Max(ChannelRangeContrast(start.G, end.G, backdrop.G),
                ChannelRangeContrast(start.B, end.B, backdrop.B))) * Math.Min(start.A, end.A) / 255D >= 45D;

    private static int ChannelRangeContrast(byte start, byte end, byte backdrop) =>
        Math.Max(0, Math.Max(Math.Min(start, end) - backdrop, backdrop - Math.Max(start, end)));
}
