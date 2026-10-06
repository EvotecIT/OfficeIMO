using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Bounds are in canvas coordinates and include shaped outlines, fallback faces,
    // horizontal advance scaling and synthetic styles used by DrawPositionedText.
    internal (double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured) MeasurePositionedTextBounds(
        string text, double x, double y, double width, double height, double size,
        OfficeFontInfo fontInfo, double advance, OfficeTextAlignment alignment,
        OfficeTextFeatureSettings features, string palette, double baselineSize,
        OfficeTextDecorationStyle underlineStyle, OfficeTextDecorationStyle strikethroughStyle,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto, bool inkOnly = false,
        OfficeTransform? inkTransform = null, OfficeColor? color = null, OfficeColor? decorationColor = null) {
        using var faceScope = PushTextFace(fontInfo.Face);
        _cancellationToken.ThrowIfCancellationRequested();
        double left = inkOnly ? double.PositiveInfinity : x, top = inkOnly ? double.PositiveInfinity : y;
        double right = inkOnly ? double.NegativeInfinity : x + width, bottom = inkOnly ? double.NegativeInfinity : y + height;
        bool hasInk = false, isMeasured = true;
        OfficeColor foreground = color ?? OfficeColor.Black;
        if (inkOnly && (text.Length == 0 || foreground.A == 0)) return (0D, 0D, 0D, 0D, false, true);
        if (_fonts != null) {
            IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, fontInfo.FamilyName, fontInfo.Face);
            if (ShouldUseFallbackRuns(runs, fontInfo.FamilyName)) {
                double measured = MeasurePositionedText(text, size, fontInfo.FamilyName, fontInfo.Style, features, textDirection);
                if (measured <= 0D) return (left, top, right, bottom, false, true);
                double cursor = ResolveTextX(x, Math.Max(1D, width), advance, alignment);
                foreach ((OfficeFontFallbackRun run, OfficeTextDirection runDirection) in PlanVisualFallbackRuns(text, fontInfo.FamilyName, fontInfo.Style, textDirection)) {
                    double runAdvance = MeasurePositionedText(run.Text, size, run.FamilyName, fontInfo.Style, features, runDirection) * advance / measured;
                    var bounds = MeasurePositionedTextBounds(run.Text, cursor, y, Math.Max(.01D, runAdvance), height,
                        size, fontInfo.WithFamilyName(run.FamilyName), Math.Max(.01D, runAdvance),
                        OfficeTextAlignment.Left, features, palette, baselineSize, underlineStyle, strikethroughStyle,
                        runDirection, inkOnly, inkTransform, color, decorationColor);
                    isMeasured &= bounds.IsMeasured;
                    hasInk |= bounds.HasInk;
                    left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
                    right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
                    cursor += runAdvance;
                }
                return (left, top, right, bottom, hasInk, isMeasured);
            }
        }
        IOfficeFontProgram? font = ResolveTextFont(text, fontInfo.FamilyName, fontInfo.Style, out OfficeFontStyle resolvedStyle);
        // An unresolved face is not evidence of an empty glyph.
        if (font == null) return inkOnly ? (0D, 0D, 0D, 0D, false, false)
            : (left - size, top - size, right + size, bottom + size, true, false);
        double naturalAdvance = MeasureResolvedText(text, font, size, features, textDirection);
        double scaleX = naturalAdvance > 0D ? advance / naturalAdvance : 1D;
        double textX = ResolveTextX(x, Math.Max(1D, width), advance, alignment);
        double textTop = y + ResolveRasterTextTop(font, size, height, baselineSize);
        OfficeFontStyle simulated = fontInfo.Style & ~resolvedStyle;
        bool hasColorLayers = TryGetResolvedColorTextContours(text, font, textX, textTop, size, features, palette,
            foreground, out List<OfficeColorGlyphContours> layers, textDirection);
        // Buffer sizing retains the historical union; ink follows the painter's choice of color or base outlines.
        if (!inkOnly || !hasColorLayers) Include(GetResolvedTextContours(text, font, textX, textTop, size, features, textDirection));
        if (hasColorLayers) foreach (OfficeColorGlyphContours layer in layers)
            if (!inkOnly || layer.Color.A != 0) Include(layer.Contours);
        double lineHeight = font.LineHeight(size);
        IncludeDecoration(underlineStyle != OfficeTextDecorationStyle.None ? underlineStyle :
            (fontInfo.Style & OfficeFontStyle.Underline) != 0 ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None,
            textTop + lineHeight * .86D);
        IncludeDecoration(strikethroughStyle != OfficeTextDecorationStyle.None ? strikethroughStyle :
            (fontInfo.Style & OfficeFontStyle.Strikethrough) != 0 ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None,
            textTop + lineHeight * .52D);
        return (left, top, right, bottom, hasInk, isMeasured);

        void IncludeDecoration(OfficeTextDecorationStyle style, double centerY) {
            if (style == OfficeTextDecorationStyle.None || (inkOnly && (decorationColor ?? foreground).A == 0)) return;
            hasInk = true;
            // Single strokes use the positioned font size. Patterned decorations use
            // the line height, matching DrawTextCore and DrawTransformedTextDecoration.
            double thickness = Math.Max(1D, (style == OfficeTextDecorationStyle.Single ? size : lineHeight) / 16D);
            double vertical = thickness / 2D;
            if (style == OfficeTextDecorationStyle.Double) vertical += Math.Max(2D, thickness * 1.8D) / 2D;
            if (style == OfficeTextDecorationStyle.Wavy) vertical += Math.Max(1D, thickness);
            IncludePoint(new OfficePoint(textX - thickness / 2D, centerY - vertical));
            IncludePoint(new OfficePoint(textX + advance + thickness / 2D, centerY - vertical));
            IncludePoint(new OfficePoint(textX - thickness / 2D, centerY + vertical));
            IncludePoint(new OfficePoint(textX + advance + thickness / 2D, centerY + vertical));
        }

        void Include(List<List<OfficePoint>> contours) {
            if (Math.Abs(scaleX - 1D) > .0001D) ScaleContoursX(contours, textX, scaleX);
            if ((simulated & OfficeFontStyle.Italic) != 0) SlantContours(contours, textTop, size);
            double boldOffset = (simulated & OfficeFontStyle.Bold) != 0 ? size / 24D : 0D;
            foreach (List<OfficePoint> contour in contours) {
                _cancellationToken.ThrowIfCancellationRequested();
                hasInk |= contour.Count > 0;
                foreach (OfficePoint point in contour) {
                    IncludePoint(point);
                    if (boldOffset > 0D) IncludePoint(new OfficePoint(point.X + boldOffset, point.Y));
                }
            }
        }

        void IncludePoint(OfficePoint point) {
            if (inkOnly && inkTransform.HasValue) point = inkTransform.Value.TransformPoint(point);
            left = Math.Min(left, point.X); top = Math.Min(top, point.Y);
            right = Math.Max(right, point.X); bottom = Math.Max(bottom, point.Y);
        }
    }
}
