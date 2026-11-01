using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private const double TextRotationEpsilon = 0.000001D;
    private const double ItalicShear = 0.22D;
    private const int MaxTextMeasurementCacheEntries = 4096;
    private Dictionary<TextMeasurementKey, double>? _textMeasurementCache;

    /// <summary>Measures text width with the managed font fallback used by this canvas.</summary>
    public double MeasureText(string? text, double fontSize = 12D) {
        return MeasureText(text, fontSize, null);
    }

    /// <summary>Measures text width with the requested font family when it can be resolved without platform font APIs.</summary>
    public double MeasureText(string? text, double fontSize, string? fontFamily) {
        return MeasureText(text, fontSize, fontFamily, OfficeFontStyle.Regular);
    }

    /// <summary>Measures text width with a requested family and bold/italic scoped face when available.</summary>
    public double MeasureText(string? text, double fontSize, string? fontFamily, OfficeFontStyle style) {
        if (string.IsNullOrEmpty(text)) {
            return 0D;
        }

        double size = Math.Max(1D, fontSize);
        var key = new TextMeasurementKey(text!, size, fontFamily, style, PreservePaintedGlyphOrder, RequestedTextFace(style));
        Dictionary<TextMeasurementKey, double> cache = _textMeasurementCache ??= new Dictionary<TextMeasurementKey, double>();
        if (cache.TryGetValue(key, out double cached)) {
            return cached;
        }

        if (_fonts != null) {
            IReadOnlyList<OfficeFontFallbackRun> fallbackRuns =
                _fonts.PlanFallbackRuns(text, fontFamily, RequestedTextFace(style));
            if (ShouldUseFallbackRuns(fallbackRuns, fontFamily)) {
                double aggregate = 0D;
                foreach (OfficeFontFallbackRun run in fallbackRuns) {
                    aggregate += MeasureText(run.Text, size, run.FamilyRequest, style);
                }
                if (cache.Count >= MaxTextMeasurementCacheEntries) cache.Clear();
                cache[key] = aggregate;
                return aggregate;
            }
        }

        IOfficeFontProgram? font = ResolveTextFont(text!, fontFamily, style, size);
        double measured = font != null
            ? MeasureResolvedText(text!, font, size)
            : MeasureFallbackText(text!, size);
        if (cache.Count >= MaxTextMeasurementCacheEntries) {
            cache.Clear();
        }

        cache[key] = measured;
        return measured;
    }

    internal double MeasurePositionedText(
        string? text,
        double fontSize,
        string? fontFamily,
        OfficeFontStyle style,
        OfficeTextFeatureSettings? featureSettings,
        OfficeTextDirection textDirection) {
        if (string.IsNullOrEmpty(text)) return 0D;
        double size = Math.Max(0.1D, fontSize);
        if (_fonts != null) {
            IReadOnlyList<OfficeFontFallbackRun> fallbackRuns = _fonts.PlanFallbackRuns(text, fontFamily, RequestedTextFace(style));
            if (ShouldUseFallbackRuns(fallbackRuns, fontFamily)) {
                double aggregate = 0D;
                foreach ((OfficeFontFallbackRun run, OfficeTextDirection runDirection) in PlanVisualFallbackRuns(text!, fontFamily, style, textDirection)) {
                    aggregate += MeasurePositionedText(
                        run.Text, size, run.FamilyRequest, style, featureSettings, runDirection);
                }
                return aggregate;
            }
        }

        IOfficeFontProgram? font = ResolveTextFont(text!, fontFamily, style, size);
        return font != null
            ? MeasureResolvedText(text!, font, size, featureSettings, textDirection)
            : MeasureFallbackText(text!, size);
    }

    /// <summary>Draws text inside a rectangle using a managed TrueType font when available.</summary>
    public void DrawText(
        string? text,
        double x,
        double y,
        double width,
        double height,
        OfficeColor color,
        double fontSize = 12D,
        OfficeTextAlignment alignment = OfficeTextAlignment.Left,
        OfficeFontStyle style = OfficeFontStyle.Regular,
        string? fontFamily = null) =>
        DrawTextCore(text, x, y, width, height, color, fontSize, alignment, style, fontFamily, OfficeTextOverflowBehavior.Ellipsis, null);

    internal void DrawBaselineText(
        string? text,
        double x,
        double y,
        double width,
        double height,
        OfficeColor color,
        double fontSize,
        OfficeTextAlignment alignment,
        OfficeFontStyle style,
        string? fontFamily) =>
        DrawTextCore(
            text, x, y, width, height, color, fontSize, alignment, style, fontFamily,
            OfficeTextOverflowBehavior.Ellipsis, textAdvanceWidth: null,
            baselineFontSize: fontSize);

    internal void DrawPositionedText(
        string? text,
        double x,
        double y,
        double width,
        double height,
        OfficeColor color,
        double fontSize,
        OfficeTextAlignment alignment,
        OfficeFontStyle style,
        string? fontFamily,
        double textAdvanceWidth) =>
        DrawTextCore(text, x, y, width, height, color, fontSize, alignment, style, fontFamily, OfficeTextOverflowBehavior.Clip, textAdvanceWidth);

    internal void DrawPositionedText(
        string? text,
        double x,
        double y,
        double width,
        double height,
        OfficeColor color,
        double fontSize,
        OfficeTextAlignment alignment,
        OfficeFontStyle style,
        string? fontFamily,
        double textAdvanceWidth,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeColor? decorationColor = null,
        OfficeTextFeatureSettings? featureSettings = null,
        string? fontPalette = null,
        double? baselineFontSize = null,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto) =>
        DrawTextCore(text, x, y, width, height, color, fontSize, alignment, style, fontFamily, OfficeTextOverflowBehavior.Clip, textAdvanceWidth, underlineStyle, strikethroughStyle, decorationColor, featureSettings, fontPalette, baselineFontSize, textDirection);

    private void DrawTextCore(
        string? text,
        double x,
        double y,
        double width,
        double height,
        OfficeColor color,
        double fontSize,
        OfficeTextAlignment alignment,
        OfficeFontStyle style,
        string? fontFamily,
        OfficeTextOverflowBehavior overflowBehavior,
        double? textAdvanceWidth,
        OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None,
        OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None,
        OfficeColor? decorationColor = null,
        OfficeTextFeatureSettings? featureSettings = null,
        string? fontPalette = null,
        double? baselineFontSize = null,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto) {
        if (string.IsNullOrEmpty(text) || color.A == 0 || width <= 0D || height <= 0D) {
            return;
        }
        if (!Enum.IsDefined(typeof(OfficeTextOverflowBehavior), overflowBehavior)) {
            throw new ArgumentOutOfRangeException(nameof(overflowBehavior));
        }
        if (textAdvanceWidth.HasValue && (textAdvanceWidth.Value <= 0D || double.IsNaN(textAdvanceWidth.Value) || double.IsInfinity(textAdvanceWidth.Value))) {
            throw new ArgumentOutOfRangeException(nameof(textAdvanceWidth));
        }

        string value = text!;
        bool retainOverflow = overflowBehavior == OfficeTextOverflowBehavior.Clip;
        double size = ResolveRasterTextSize(fontSize, height, textAdvanceWidth.HasValue);
        if (!retainOverflow) {
            value = FitRasterText(value, size, Math.Max(1D, width - 6D), fontFamily, style, featureSettings, textDirection);
            if (value.Length == 0) return;
        }
        if (TryDrawMixedText(
            value,
            x,
            y,
            width,
            height,
            color,
            fontSize,
            alignment,
            style,
            fontFamily,
            overflowBehavior,
            textAdvanceWidth,
            underlineStyle,
            strikethroughStyle,
            decorationColor,
            featureSettings,
            fontPalette,
            baselineFontSize,
            textDirection)) {
            return;
        }
        IOfficeFontProgram? font = ResolveTextFont(value, fontFamily, style, size, out OfficeFontStyle resolvedStyle);
        OfficeFontStyle simulatedStyle = style & ~resolvedStyle;
        if (font != null) {
            double measured = MeasureResolvedText(value, font, size, featureSettings, textDirection);
            double availableWidth = Math.Max(retainOverflow ? .01D : 1D, retainOverflow ? width : width - 6D);

            double resolvedAdvance = textAdvanceWidth.HasValue && string.Equals(value, text, StringComparison.Ordinal)
                ? textAdvanceWidth.Value
                : measured;
            double top = y + ResolveRasterTextTop(font, size, height, baselineFontSize);
            if (baselineFontSize.HasValue) top += SnapBaselineOffset(y + baselineFontSize.Value);
            double textX = ResolveTextX(retainOverflow ? x : x + 3D, availableWidth, resolvedAdvance, alignment);
            double horizontalScale = measured > 0D && Math.Abs(resolvedAdvance - measured) > 0.0001D
                ? resolvedAdvance / measured
                : 1D;
            if (TryGetResolvedColorTextContours(value, font, textX, top, size, featureSettings, fontPalette, color, out List<OfficeColorGlyphContours> colorLayers, textDirection)) {
                foreach (OfficeColorGlyphContours layer in colorLayers) {
                    if (Math.Abs(horizontalScale - 1D) > 0.0001D) ScaleContoursX(layer.Contours, textX, horizontalScale);
                    if ((simulatedStyle & OfficeFontStyle.Italic) == OfficeFontStyle.Italic) SlantContours(layer.Contours, top, size);
                    FillTextContours(layer.Contours, layer.Color, (simulatedStyle & OfficeFontStyle.Bold) != 0 ? OfficeSyntheticTextStyle.BoldOffset(size) : 0D);
                }
            } else {
                List<List<OfficePoint>> contours = GetResolvedTextContours(value, font, textX, top, size, featureSettings, textDirection);
                if (Math.Abs(horizontalScale - 1D) > 0.0001D) ScaleContoursX(contours, textX, horizontalScale);
                if ((simulatedStyle & OfficeFontStyle.Italic) == OfficeFontStyle.Italic) SlantContours(contours, top, size);
                PaintTextContours(contours, color, (simulatedStyle & OfficeFontStyle.Bold) != 0 ? OfficeSyntheticTextStyle.BoldOffset(size) : 0D);
            }

            OfficeTextDecorationStyle resolvedUnderlineStyle = underlineStyle != OfficeTextDecorationStyle.None
                ? underlineStyle
                : (style & OfficeFontStyle.Underline) == OfficeFontStyle.Underline
                    ? OfficeTextDecorationStyle.Single
                    : OfficeTextDecorationStyle.None;
            OfficeTextDecorationStyle resolvedStrikethroughStyle = strikethroughStyle != OfficeTextDecorationStyle.None
                ? strikethroughStyle
                : (style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough
                    ? OfficeTextDecorationStyle.Single
                    : OfficeTextDecorationStyle.None;
            if (resolvedUnderlineStyle == OfficeTextDecorationStyle.Single) {
                double underlineY = top + (font.LineHeight(size) * 0.86D);
                DrawLine(textX, underlineY, textX + resolvedAdvance, underlineY, decorationColor ?? color, Math.Max(1D, size / 16D));
            } else if (resolvedUnderlineStyle != OfficeTextDecorationStyle.None) {
                DrawTextLineDecorations(textX, resolvedAdvance, top, font.LineHeight(size), decorationColor ?? color, 0D, 0D, 0D, resolvedUnderlineStyle, OfficeTextDecorationStyle.None, false, false);
            }

            if (resolvedStrikethroughStyle == OfficeTextDecorationStyle.Single) {
                double strikeY = top + (font.LineHeight(size) * 0.52D);
                DrawLine(textX, strikeY, textX + resolvedAdvance, strikeY, decorationColor ?? color, Math.Max(1D, size / 16D));
            } else if (resolvedStrikethroughStyle != OfficeTextDecorationStyle.None) {
                DrawTextLineDecorations(textX, resolvedAdvance, top, font.LineHeight(size), decorationColor ?? color, 0D, 0D, 0D, OfficeTextDecorationStyle.None, resolvedStrikethroughStyle, false, false);
            }
            return;
        }

        if (_textInkObserver != null) { _textInkObserver(null); return; }
        DrawFallbackText(
            value,
            x,
            y,
            width,
            height,
            color,
            size,
            alignment,
            style,
            overflowBehavior,
            textAdvanceWidth,
            underlineStyle,
            strikethroughStyle,
            decorationColor,
            baselineFontSize);
    }

    private void DrawTextLineDecorations(
        double x,
        double width,
        double top,
        double fontHeight,
        OfficeColor color,
        double rotationRadians,
        double rotationCenterX,
        double rotationCenterY,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        bool flipHorizontal,
        bool flipVertical) {
        if (width <= 0D || color.A == 0) {
            return;
        }

        if (underlineStyle != OfficeTextDecorationStyle.None) {
            DrawTransformedTextDecoration(x, width, top + (fontHeight * 0.86D), color, fontHeight, rotationRadians, rotationCenterX, rotationCenterY, underlineStyle, flipHorizontal, flipVertical);
        }

        if (strikethroughStyle != OfficeTextDecorationStyle.None) {
            DrawTransformedTextDecoration(x, width, top + (fontHeight * 0.52D), color, fontHeight, rotationRadians, rotationCenterX, rotationCenterY, strikethroughStyle, flipHorizontal, flipVertical);
        }
    }

    private void DrawTransformedTextDecoration(
        double x,
        double width,
        double y,
        OfficeColor color,
        double fontHeight,
        double rotationRadians,
        double rotationCenterX,
        double rotationCenterY,
        OfficeTextDecorationStyle style,
        bool flipHorizontal,
        bool flipVertical) {
        double thickness = Math.Max(1D, fontHeight / 16D);
        double separation = Math.Max(2D, thickness * 1.8D);
        if (InspectTextDecoration(x, width, y, fontHeight, style, rotationRadians, rotationCenterX,
            rotationCenterY, flipHorizontal, flipVertical)) return;

        if (style == OfficeTextDecorationStyle.Double) {
            DrawTransformedTextDecoration(x, width, y - (separation / 2D), color, fontHeight, rotationRadians, rotationCenterX, rotationCenterY, OfficeTextDecorationStyle.Single, flipHorizontal, flipVertical);
            DrawTransformedTextDecoration(x, width, y + (separation / 2D), color, fontHeight, rotationRadians, rotationCenterX, rotationCenterY, OfficeTextDecorationStyle.Single, flipHorizontal, flipVertical);
            return;
        }

        if (style == OfficeTextDecorationStyle.Wavy) {
            double wavelength = Math.Max(4D, fontHeight * 0.36D);
            int steps = Math.Max(4, (int)Math.Ceiling(width / Math.Max(1D, wavelength / 4D)));
            var points = new List<OfficePoint>(steps + 1);
            for (int index = 0; index <= steps; index++) {
                double offset = width * index / steps;
                double waveY = y + (Math.Sin(offset * Math.PI * 2D / wavelength) * Math.Max(1D, thickness));
                points.Add(TransformFramePoint(new OfficePoint(x + offset, waveY), rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical));
            }
            DrawPolyline(points, color, thickness);
            return;
        }

        OfficePoint start = TransformFramePoint(new OfficePoint(x, y), rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
        OfficePoint end = TransformFramePoint(new OfficePoint(x + width, y), rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
        if (style == OfficeTextDecorationStyle.Dashed) {
            DrawDashedLine(start.X, start.Y, end.X, end.Y, color, thickness, Math.Max(2D, fontHeight * 0.22D), Math.Max(1D, fontHeight * 0.12D));
        } else if (style == OfficeTextDecorationStyle.Dotted) {
            DrawDashedLine(start.X, start.Y, end.X, end.Y, color, thickness, thickness, Math.Max(thickness, fontHeight * 0.12D));
        } else {
            DrawLine(start.X, start.Y, end.X, end.Y, color, thickness);
        }
    }

    private void DrawAffineTextLineDecorations(
        double x,
        double width,
        double top,
        double fontHeight,
        OfficeColor color,
        OfficeTransform transform,
        bool underline,
        bool strikethrough) {
        if (width <= 0D || color.A == 0) {
            return;
        }

        if (underline) {
            DrawAffineTextDecorationLine(x, width, top + (fontHeight * 0.86D), color, fontHeight, transform);
        }

        if (strikethrough) {
            DrawAffineTextDecorationLine(x, width, top + (fontHeight * 0.52D), color, fontHeight, transform);
        }
    }

    private void DrawAffineTextDecorationLine(double x, double width, double y, OfficeColor color, double fontHeight, OfficeTransform transform) {
        OfficePoint start = transform.TransformPoint(new OfficePoint(x, y));
        OfficePoint end = transform.TransformPoint(new OfficePoint(x + width, y));
        DrawLine(start.X, start.Y, end.X, end.Y, color, Math.Max(1D, fontHeight / 16D));
    }

    private void DrawTransformedTextDecorationLine(
        double x,
        double width,
        double y,
        OfficeColor color,
        double fontHeight,
        double rotationRadians,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical) {
        OfficePoint start = TransformFramePoint(new OfficePoint(x, y), rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
        OfficePoint end = TransformFramePoint(new OfficePoint(x + width, y), rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
        DrawLine(start.X, start.Y, end.X, end.Y, color, Math.Max(1D, fontHeight / 16D));
    }

    private static double ResolveTextX(double left, double width, double measured, OfficeTextAlignment alignment) {
        if (alignment == OfficeTextAlignment.Right) {
            return left + Math.Max(0D, width - measured);
        }

        if (alignment == OfficeTextAlignment.Center) {
            return left + Math.Max(0D, (width - measured) / 2D);
        }

        return left;
    }

    private static void SlantContours(List<List<OfficePoint>> contours, double top, double fontSize) {
        double baseY = top + fontSize;
        for (int i = 0; i < contours.Count; i++) {
            List<OfficePoint> contour = contours[i];
            for (int j = 0; j < contour.Count; j++) {
                OfficePoint point = contour[j];
                contour[j] = new OfficePoint(point.X + OfficeSyntheticTextStyle.ItalicOffset(baseY, point.Y), point.Y);
            }
        }
    }

    private static void ScaleContoursX(List<List<OfficePoint>> contours, double originX, double scaleX) {
        for (int i = 0; i < contours.Count; i++) {
            List<OfficePoint> contour = contours[i];
            for (int j = 0; j < contour.Count; j++) {
                OfficePoint point = contour[j];
                contour[j] = new OfficePoint(originX + ((point.X - originX) * scaleX), point.Y);
            }
        }
    }

    private static void OffsetContours(List<List<OfficePoint>> contours, double offsetX, double offsetY) {
        for (int i = 0; i < contours.Count; i++) {
            List<OfficePoint> contour = contours[i];
            for (int j = 0; j < contour.Count; j++) {
                OfficePoint point = contour[j];
                contour[j] = new OfficePoint(point.X + offsetX, point.Y + offsetY);
            }
        }
    }

    private static double MeasureFallbackText(string text, double fontSize) {
        return MeasureStrokeText(text, fontSize);
    }

    private static double ResolveAnchoredTextX(double anchorX, double width, OfficeTextAlignment alignment) {
        if (alignment == OfficeTextAlignment.Right) {
            return anchorX - width;
        }

        if (alignment == OfficeTextAlignment.Center) {
            return anchorX - (width / 2D);
        }

        return anchorX;
    }

}
