using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Bounds are in canvas coordinates and include shaped outlines, fallback faces,
    // horizontal advance scaling and synthetic styles used by DrawPositionedText.
    internal (double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured, bool IsClipped) MeasurePositionedTextBounds(
        string text, double x, double y, double width, double height, double size,
        OfficeFontInfo fontInfo, double advance, OfficeTextAlignment alignment,
        OfficeTextFeatureSettings features, string palette, double baselineSize,
        OfficeTextDecorationStyle underlineStyle, OfficeTextDecorationStyle strikethroughStyle,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto, bool inkOnly = false,
        OfficeTransform? inkTransform = null, OfficeColor? color = null, OfficeColor? decorationColor = null,
        IReadOnlyList<OfficeTextInkClip>? inkClips = null) {
        using var faceScope = PushTextFace(fontInfo.Face);
        _cancellationToken.ThrowIfCancellationRequested();
        double left = inkOnly ? double.PositiveInfinity : x, top = inkOnly ? double.PositiveInfinity : y;
        double right = inkOnly ? double.NegativeInfinity : x + width, bottom = inkOnly ? double.NegativeInfinity : y + height;
        bool hasInk = false, isMeasured = true, isClipped = false;
        long remainingClipWork = 4_000_000;
        OfficeColor foreground = color ?? OfficeColor.Black;
        if (inkOnly && (text.Length == 0 || foreground.A == 0)) return (0D, 0D, 0D, 0D, false, true, false);
        if (_fonts != null) {
            IReadOnlyList<OfficeFontFallbackRun> runs = _fonts.PlanFallbackRuns(text, fontInfo.FamilyName, fontInfo.Face);
            if (ShouldUseFallbackRuns(runs, fontInfo.FamilyName)) {
                double measured = MeasurePositionedText(text, size, fontInfo.FamilyName, fontInfo.Style, features, textDirection);
                if (measured <= 0D) return (left, top, right, bottom, false, true, false);
                double cursor = ResolveTextX(x, Math.Max(1D, width), advance, alignment);
                foreach ((OfficeFontFallbackRun run, OfficeTextDirection runDirection) in PlanVisualFallbackRuns(text, fontInfo.FamilyName, fontInfo.Style, textDirection)) {
                    double runAdvance = MeasurePositionedText(run.Text, size, run.FamilyName, fontInfo.Style, features, runDirection) * advance / measured;
                    var bounds = MeasurePositionedTextBounds(run.Text, cursor, y, Math.Max(.01D, runAdvance), height,
                        size, fontInfo.WithFamilyName(run.FamilyName), Math.Max(.01D, runAdvance),
                        OfficeTextAlignment.Left, features, palette, baselineSize, underlineStyle, strikethroughStyle,
                        runDirection, inkOnly, inkTransform, color, decorationColor, inkClips);
                    isMeasured &= bounds.IsMeasured;
                    isClipped |= bounds.IsClipped;
                    hasInk |= bounds.HasInk;
                    if (!inkOnly || bounds.HasInk) {
                        left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
                        right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
                    }
                    cursor += runAdvance;
                }
                return (left, top, right, bottom, hasInk, isMeasured, isClipped);
            }
        }
        IOfficeFontProgram? font = ResolveTextFont(text, fontInfo.FamilyName, fontInfo.Style, out OfficeFontStyle resolvedStyle);
        // An unresolved face is not evidence of an empty glyph.
        if (font == null) return inkOnly ? (0D, 0D, 0D, 0D, false, false, false)
            : (left - size, top - size, right + size, bottom + size, true, false, false);
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
        return (left, top, right, bottom, hasInk, isMeasured, isClipped);

        void IncludeDecoration(OfficeTextDecorationStyle style, double centerY) {
            if (style == OfficeTextDecorationStyle.None || (inkOnly && (decorationColor ?? foreground).A == 0)) return;
            // Single strokes use the positioned font size. Patterned decorations use
            // the line height, matching DrawTextCore and DrawTransformedTextDecoration.
            double thickness = Math.Max(1D, (style == OfficeTextDecorationStyle.Single ? size : lineHeight) / 16D);
            double vertical = thickness / 2D;
            if (style == OfficeTextDecorationStyle.Double) vertical += Math.Max(2D, thickness * 1.8D) / 2D;
            if (style == OfficeTextDecorationStyle.Wavy) vertical += Math.Max(1D, thickness);
            IncludeContour(new List<OfficePoint> {
                new OfficePoint(textX - thickness / 2D, centerY - vertical),
                new OfficePoint(textX + advance + thickness / 2D, centerY - vertical),
                new OfficePoint(textX + advance + thickness / 2D, centerY + vertical),
                new OfficePoint(textX - thickness / 2D, centerY + vertical)
            });
        }

        void Include(List<List<OfficePoint>> contours) {
            if (Math.Abs(scaleX - 1D) > .0001D) ScaleContoursX(contours, textX, scaleX);
            if ((simulated & OfficeFontStyle.Italic) != 0) SlantContours(contours, textTop, size);
            double boldOffset = (simulated & OfficeFontStyle.Bold) != 0 ? size / 24D : 0D;
            if (inkOnly) {
                IncludeFilledContours(contours, 0D);
                if (boldOffset > 0D) IncludeFilledContours(contours, boldOffset);
            } else foreach (List<OfficePoint> contour in contours) {
                _cancellationToken.ThrowIfCancellationRequested();
                IncludeContour(contour);
                if (boldOffset > 0D) IncludeContour(contour, boldOffset);
            }
        }

        void IncludeFilledContours(List<List<OfficePoint>> contours, double offsetX) {
            var prepared = new List<List<OfficePoint>>(contours.Count);
            foreach (var contour in contours) prepared.Add(PrepareContour(contour, offsetX));
            var bounds = MeasureFilledContourBounds(prepared, OfficeFillRule.NonZero, inkClips);
            isMeasured &= bounds.IsMeasured;
            isClipped |= bounds.IsClipped;
            if (!bounds.HasInk) return;
            hasInk = true;
            left = Math.Min(left, bounds.Left); top = Math.Min(top, bounds.Top);
            right = Math.Max(right, bounds.Right); bottom = Math.Max(bottom, bounds.Bottom);
        }

        List<OfficePoint> PrepareContour(List<OfficePoint> contour, double offsetX) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (inkOnly && (inkTransform.HasValue || inkClips != null)) {
                var transformed = new List<OfficePoint>(contour.Count);
                foreach (OfficePoint point in contour) {
                    var shifted = new OfficePoint(point.X + offsetX, point.Y);
                    transformed.Add(inkTransform.HasValue ? inkTransform.Value.TransformPoint(shifted) : shifted);
                }
                contour = transformed; offsetX = 0D;
            }
            if (inkOnly && inkClips != null) foreach (OfficeTextInkClip clip in inkClips)
                if (clip.FilledContours == null) contour = clip.Apply(contour, ref isClipped, ref remainingClipWork, _cancellationToken);
            if (offsetX != 0D) {
                var shifted = new List<OfficePoint>(contour.Count);
                foreach (OfficePoint point in contour) shifted.Add(new OfficePoint(point.X + offsetX, point.Y));
                contour = shifted;
            }
            return contour;
        }

        void IncludeContour(List<OfficePoint> contour, double offsetX = 0D) {
            if (inkOnly) {
                IncludeFilledContours(new List<List<OfficePoint>> { contour }, offsetX);
                return;
            }
            hasInk |= contour.Count > 0;
            foreach (OfficePoint point in contour) {
                left = Math.Min(left, point.X + offsetX); top = Math.Min(top, point.Y);
                right = Math.Max(right, point.X + offsetX); bottom = Math.Max(bottom, point.Y);
            }
        }
    }
}
