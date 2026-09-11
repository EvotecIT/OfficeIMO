using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Bounds are in canvas coordinates and include shaped outlines, fallback faces,
    // horizontal advance scaling and synthetic styles used by DrawPositionedText.
    internal (double Left, double Top, double Right, double Bottom) MeasurePositionedTextBounds(
        string text, double x, double y, double width, double height, double size,
        OfficeFontInfo fontInfo, double advance, OfficeTextAlignment alignment,
        OfficeTextFeatureSettings features, string palette, double baselineSize) {
        _cancellationToken.ThrowIfCancellationRequested();
        double left = x, top = y, right = x + width, bottom = y + height;
        if (_fonts != null) {
            IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, fontInfo.FamilyName, fontInfo.Style);
            if (ShouldUseFallbackRuns(runs, fontInfo.FamilyName)) {
                double measured = MeasureText(text, size, fontInfo.FamilyName, fontInfo.Style);
                if (measured <= 0D) return (left, top, right, bottom);
                double cursor = ResolveTextX(x, Math.Max(1D, width), advance, alignment);
                foreach (OfficeFontFallbackRun run in runs) {
                    double runAdvance = MeasureText(run.Text, size, run.FamilyName, fontInfo.Style) * advance / measured;
                    var bounds = MeasurePositionedTextBounds(run.Text, cursor, y, Math.Max(.01D, runAdvance), height,
                        size, new OfficeFontInfo(run.FamilyName, fontInfo.Size, fontInfo.Style), Math.Max(.01D, runAdvance),
                        OfficeTextAlignment.Left, features, palette, baselineSize);
                    left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
                    right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
                    cursor += runAdvance;
                }
                return (left, top, right, bottom);
            }
        }
        IOfficeFontProgram? font = ResolveTextFont(text, fontInfo.FamilyName, fontInfo.Style, out OfficeFontStyle resolvedStyle);
        if (font == null) return (left - size, top - size, right + size, bottom + size);
        double naturalAdvance = MeasureResolvedText(text, font, size, features);
        double scaleX = naturalAdvance > 0D ? advance / naturalAdvance : 1D;
        double textX = ResolveTextX(x, Math.Max(1D, width), advance, alignment);
        double textTop = y + ResolveRasterTextTop(font, size, height, baselineSize);
        OfficeFontStyle simulated = fontInfo.Style & ~resolvedStyle;
        Include(GetResolvedTextContours(text, font, textX, textTop, size, features));
        if (TryGetResolvedColorTextContours(text, font, textX, textTop, size, features, palette,
            OfficeColor.Black, out List<OfficeColorGlyphContours> layers)) {
            foreach (OfficeColorGlyphContours layer in layers) Include(layer.Contours);
        }
        return (left, top, right, bottom);

        void Include(List<List<OfficePoint>> contours) {
            if (Math.Abs(scaleX - 1D) > .0001D) ScaleContoursX(contours, textX, scaleX);
            if ((simulated & OfficeFontStyle.Italic) != 0) SlantContours(contours, textTop, size);
            double boldOffset = (simulated & OfficeFontStyle.Bold) != 0 ? .45D : 0D;
            foreach (List<OfficePoint> contour in contours) {
                _cancellationToken.ThrowIfCancellationRequested();
                foreach (OfficePoint point in contour) {
                    left = Math.Min(left, point.X); top = Math.Min(top, point.Y);
                    right = Math.Max(right, point.X + boldOffset); bottom = Math.Max(bottom, point.Y);
                }
            }
        }
    }
}
