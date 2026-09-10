using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private Dictionary<TextMeasurementKey, OfficeTextPaintBounds>? _textPaintBoundsCache;

    // Reuse the renderer's bounded, shaped outlines and fallback selection. Width alone
    // cannot establish whether an accent or descender fits a drawing frame.
    internal OfficeTextPaintBounds MeasureTextPaintBounds(string? text, double fontSize, string? family, OfficeFontStyle style) {
        if (string.IsNullOrWhiteSpace(text)) return default;
        _cancellationToken.ThrowIfCancellationRequested();
        double size = Math.Max(.1D, fontSize);
        var key = new TextMeasurementKey(text!, size, family, style);
        var cache = _textPaintBoundsCache ??= new Dictionary<TextMeasurementKey, OfficeTextPaintBounds>();
        if (cache.TryGetValue(key, out OfficeTextPaintBounds found)) return found;
        double top = -size * .84D, bottom = size * .16D;
        if (_fonts != null) {
            IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, family, style);
            if (ShouldUseFallbackRuns(runs, family)) {
                foreach (OfficeFontFallbackRun run in runs) {
                    OfficeTextPaintBounds bounds = MeasureTextPaintBounds(run.Text, size, run.FamilyName, style);
                    top = Math.Min(top, bounds.Top); bottom = Math.Max(bottom, bounds.Bottom);
                }
                return Store();
            }
        }
        IOfficeFontProgram? font = ResolveTextFont(text, family, style);
        if (font != null) {
            double origin = -ResolveRasterBaseline(font, size);
            // Fitted raster text paints base outlines, while positioned/SVG paths
            // may use color layers. Bound both representations of the same font.
            Include(GetResolvedTextContours(text!, font, 0D, origin, size));
            if (TryGetResolvedColorTextContours(text!, font, 0D, origin, size, null, null, OfficeColor.Black, out List<OfficeColorGlyphContours> layers)) {
                foreach (OfficeColorGlyphContours layer in layers) Include(layer.Contours);
            }
        }
        return Store();

        void Include(List<List<OfficePoint>> contours) {
            foreach (List<OfficePoint> contour in contours)
                foreach (OfficePoint point in contour) {
                    top = Math.Min(top, point.Y); bottom = Math.Max(bottom, point.Y);
                }
        }

        OfficeTextPaintBounds Store() {
            var bounds = new OfficeTextPaintBounds(top, bottom);
            if (cache.Count >= MaxTextMeasurementCacheEntries) cache.Clear();
            cache[key] = bounds;
            return bounds;
        }
    }
}
