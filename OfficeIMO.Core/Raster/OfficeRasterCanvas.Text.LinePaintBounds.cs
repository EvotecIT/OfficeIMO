using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private Dictionary<TextMeasurementKey, (double Left, double Top, double Right, double Bottom, bool HasInk)>? _textLineInkBoundsCache;

    // Actual DrawTextLine ink relative to its left/top, without a nominal em box
    // or whitespace advance. Selection, shaping and synthetic styles match paint.
    internal (double Left, double Top, double Right, double Bottom, bool HasInk) MeasureTextLineInkBounds(
        string text, double size, string? family, OfficeFontStyle style) {
        _cancellationToken.ThrowIfCancellationRequested();
        size = Math.Max(1D, size);
        var key = new TextMeasurementKey(text, size, family, style, PreservePaintedGlyphOrder, RequestedTextFace(style));
        var cache = _textLineInkBoundsCache ??= new Dictionary<TextMeasurementKey, (double, double, double, double, bool)>();
        if (cache.TryGetValue(key, out var found)) return found;
        double left = double.PositiveInfinity, top = double.PositiveInfinity;
        double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
        bool hasInk = false;
        if (!string.IsNullOrWhiteSpace(text) && _fonts != null) {
            IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, family, RequestedTextFace(style));
            if (ShouldUseFallbackRuns(runs, family)) {
                double cursor = 0D;
                foreach (OfficeFontFallbackRun run in runs) {
                    var bounds = MeasureTextLineInkBounds(run.Text, size, run.FamilyName, style);
                    if (bounds.HasInk) Include(bounds.Left + cursor, bounds.Top, bounds.Right + cursor, bounds.Bottom);
                    cursor += MeasureText(run.Text, size, run.FamilyName, style);
                }
                return Store();
            }
        }
        if (string.IsNullOrWhiteSpace(text)) return Store();
        IOfficeFontProgram? font = ResolveTextFont(text, family, style, size, out OfficeFontStyle resolvedStyle);
        bool bold = (style & OfficeFontStyle.Bold) != 0, italic = (style & OfficeFontStyle.Italic) != 0;
        if (font != null) {
            var contours = TransformTextContours(GetResolvedTextContours(text, font, 0D,
                ResolveRasterTextLineOutlineTop(font, size, 0D), size), size,
                italic && (resolvedStyle & OfficeFontStyle.Italic) == 0, 0D, 0D, 0D, false, false);
            double boldOffset = bold && (resolvedStyle & OfficeFontStyle.Bold) == 0 ? OfficeSyntheticTextStyle.BoldOffset(size) : 0D;
            foreach (var contour in contours) {
                _cancellationToken.ThrowIfCancellationRequested();
                foreach (OfficePoint point in contour) Include(point.X, point.Y, point.X + boldOffset, point.Y);
            }
        } else {
            var bounds = MeasureStrokeTextBounds(text, 0D, 0D, size, bold, italic);
            if (bounds.HasInk) Include(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
        }
        return Store();

        void Include(double x1, double y1, double x2, double y2) {
            hasInk = true;
            left = Math.Min(left, x1); top = Math.Min(top, y1);
            right = Math.Max(right, x2); bottom = Math.Max(bottom, y2);
        }
        (double Left, double Top, double Right, double Bottom, bool HasInk) Store() {
            if (cache.Count >= MaxTextMeasurementCacheEntries) cache.Clear();
            return cache[key] = (left, top, right, bottom, hasInk);
        }
    }

    internal IEnumerable<(double Left, double Top, double Right, double Bottom)> TextLinePaintBounds(
        string text, double size, string? family, OfficeFontStyle style,
        OfficeTextDecorationStyle underline, OfficeTextDecorationStyle strikethrough) {
        if (string.IsNullOrEmpty(text)) yield break;
        size = Math.Max(1D, size);
        var ink = MeasureTextLineInkBounds(text, size, family, style);
        if (ink.HasInk) yield return (ink.Left, ink.Top, ink.Right, ink.Bottom);
        double width = MeasureText(text, size, family, style);
        if (width <= 0D) yield break;
        if (underline != OfficeTextDecorationStyle.None) yield return TextLineDecorationBounds(width, size, underline, size * .86D);
        if (strikethrough != OfficeTextDecorationStyle.None) yield return TextLineDecorationBounds(width, size, strikethrough, size * .52D);
    }

    internal static (double Left, double Top, double Right, double Bottom) TextLineDecorationBounds(
        double width, double size, OfficeTextDecorationStyle style, double centerY) {
        // Match DrawTransformedTextDecoration's stroke, double separation and wave amplitude.
        double thickness = Math.Max(1D, size / 16D), vertical = thickness / 2D;
        if (style == OfficeTextDecorationStyle.Double) vertical += Math.Max(2D, thickness * 1.8D) / 2D;
        if (style == OfficeTextDecorationStyle.Wavy) vertical += Math.Max(1D, thickness);
        return (-thickness / 2D, centerY - vertical, width + thickness / 2D, centerY + vertical);
    }
}
