using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Bounds are in canvas coordinates and include shaped outlines, fallback faces,
    // horizontal advance scaling and synthetic styles used by DrawPositionedText.
    internal (double Left, double Top, double Right, double Bottom, bool HasInk) MeasurePositionedTextBounds(
        string text, double x, double y, double width, double height, double size,
        OfficeFontInfo fontInfo, double advance, OfficeTextAlignment alignment,
        OfficeTextFeatureSettings features, string palette, double baselineSize,
        OfficeTextDecorationStyle underlineStyle, OfficeTextDecorationStyle strikethroughStyle) {
        _cancellationToken.ThrowIfCancellationRequested();
        double left = x, top = y, right = x + width, bottom = y + height;
        bool hasInk = false;
        if (_fonts != null) {
            IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, fontInfo.FamilyName, fontInfo.Style);
            if (ShouldUseFallbackRuns(runs, fontInfo.FamilyName)) {
                double measured = MeasureText(text, size, fontInfo.FamilyName, fontInfo.Style);
                if (measured <= 0D) return (left, top, right, bottom, false);
                double cursor = ResolveTextX(x, Math.Max(1D, width), advance, alignment);
                foreach (OfficeFontFallbackRun run in runs) {
                    double runAdvance = MeasureText(run.Text, size, run.FamilyName, fontInfo.Style) * advance / measured;
                    var bounds = MeasurePositionedTextBounds(run.Text, cursor, y, Math.Max(.01D, runAdvance), height,
                        size, new OfficeFontInfo(run.FamilyName, fontInfo.Size, fontInfo.Style), Math.Max(.01D, runAdvance),
                        OfficeTextAlignment.Left, features, palette, baselineSize, underlineStyle, strikethroughStyle);
                    hasInk |= bounds.HasInk;
                    left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
                    right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
                    cursor += runAdvance;
                }
                return (left, top, right, bottom, hasInk);
            }
        }
        IOfficeFontProgram? font = ResolveTextFont(text, fontInfo.FamilyName, fontInfo.Style, out OfficeFontStyle resolvedStyle);
        // An unresolved face is not evidence of an empty glyph.
        if (font == null) return (left - size, top - size, right + size, bottom + size, true);
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
        double lineHeight = font.LineHeight(size);
        IncludeDecoration(underlineStyle != OfficeTextDecorationStyle.None ? underlineStyle :
            (fontInfo.Style & OfficeFontStyle.Underline) != 0 ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None,
            textTop + lineHeight * .86D);
        IncludeDecoration(strikethroughStyle != OfficeTextDecorationStyle.None ? strikethroughStyle :
            (fontInfo.Style & OfficeFontStyle.Strikethrough) != 0 ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None,
            textTop + lineHeight * .52D);
        return (left, top, right, bottom, hasInk);

        void IncludeDecoration(OfficeTextDecorationStyle style, double centerY) {
            if (style == OfficeTextDecorationStyle.None) return;
            hasInk = true;
            // Single strokes use the positioned font size. Patterned decorations use
            // the line height, matching DrawTextCore and DrawTransformedTextDecoration.
            double thickness = Math.Max(1D, (style == OfficeTextDecorationStyle.Single ? size : lineHeight) / 16D);
            double vertical = thickness / 2D;
            if (style == OfficeTextDecorationStyle.Double) vertical += Math.Max(2D, thickness * 1.8D) / 2D;
            if (style == OfficeTextDecorationStyle.Wavy) vertical += Math.Max(1D, thickness);
            left = Math.Min(left, textX - thickness / 2D);
            right = Math.Max(right, textX + advance + thickness / 2D);
            top = Math.Min(top, centerY - vertical);
            bottom = Math.Max(bottom, centerY + vertical);
        }

        void Include(List<List<OfficePoint>> contours) {
            if (Math.Abs(scaleX - 1D) > .0001D) ScaleContoursX(contours, textX, scaleX);
            if ((simulated & OfficeFontStyle.Italic) != 0) SlantContours(contours, textTop, size);
            double boldOffset = (simulated & OfficeFontStyle.Bold) != 0 ? .45D : 0D;
            foreach (List<OfficePoint> contour in contours) {
                _cancellationToken.ThrowIfCancellationRequested();
                hasInk |= contour.Count > 0;
                foreach (OfficePoint point in contour) {
                    left = Math.Min(left, point.X); top = Math.Min(top, point.Y);
                    right = Math.Max(right, point.X + boldOffset); bottom = Math.Max(bottom, point.Y);
                }
            }
        }
    }
}
