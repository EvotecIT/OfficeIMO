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
        var key = new TextMeasurementKey(text!, size, fontFamily, style);
        Dictionary<TextMeasurementKey, double> cache = _textMeasurementCache ??= new Dictionary<TextMeasurementKey, double>();
        if (cache.TryGetValue(key, out double cached)) {
            return cached;
        }

        OfficeTrueTypeFont? font = ResolveTextFont(text!, fontFamily, style);
        double measured = font != null
            ? MeasureResolvedText(text!, font, size)
            : MeasureFallbackText(text!, size);
        if (cache.Count >= MaxTextMeasurementCacheEntries) {
            cache.Clear();
        }

        cache[key] = measured;
        return measured;
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
        double? textAdvanceWidth) {
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
        double size = Math.Max(6D, Math.Min(fontSize, height - 2D));
        OfficeTrueTypeFont? font = ResolveTextFont(value, fontFamily, style, out OfficeFontStyle resolvedStyle);
        OfficeFontStyle simulatedStyle = style & ~resolvedStyle;
        if (font != null) {
            double measured = MeasureResolvedText(value, font, size);
            double availableWidth = Math.Max(1D, retainOverflow ? width : width - 6D);
            if (!retainOverflow) {
                while (measured > availableWidth && value.Length > 0) {
                    value = OfficeTextElements.RemoveLast(value);
                    if (value.Length == 0) break;
                    measured = MeasureResolvedText(value + "...", font, size);
                }

                if (value.Length == 0 && MeasureResolvedText("...", font, size) > availableWidth) return;
                if (!string.Equals(value, text, StringComparison.Ordinal)) {
                    value += "...";
                    measured = MeasureResolvedText(value, font, size);
                }
            }

            double resolvedAdvance = textAdvanceWidth.HasValue && string.Equals(value, text, StringComparison.Ordinal)
                ? textAdvanceWidth.Value
                : measured;
            double top = y + Math.Max(1D, (height - font.LineHeight(size)) / 2D);
            double textX = ResolveTextX(retainOverflow ? x : x + 3D, availableWidth, resolvedAdvance, alignment);
            List<List<OfficePoint>> contours = GetResolvedTextContours(value, font, textX, top, size);
            if (measured > 0D && Math.Abs(resolvedAdvance - measured) > 0.0001D) {
                ScaleContoursX(contours, textX, resolvedAdvance / measured);
            }
            if ((simulatedStyle & OfficeFontStyle.Italic) == OfficeFontStyle.Italic) {
                SlantContours(contours, top, size);
            }

            FillContours(contours, color, OfficeFillRule.EvenOdd);
            if ((simulatedStyle & OfficeFontStyle.Bold) == OfficeFontStyle.Bold) {
                OffsetContours(contours, 0.45D, 0D);
                FillContours(contours, color, OfficeFillRule.EvenOdd);
            }

            if ((style & OfficeFontStyle.Underline) == OfficeFontStyle.Underline) {
                double underlineY = top + (font.LineHeight(size) * 0.86D);
                DrawLine(textX, underlineY, textX + resolvedAdvance, underlineY, color, Math.Max(1D, size / 16D));
            }

            if ((style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough) {
                double strikeY = top + (font.LineHeight(size) * 0.52D);
                DrawLine(textX, strikeY, textX + resolvedAdvance, strikeY, color, Math.Max(1D, size / 16D));
            }
            return;
        }

        DrawFallbackText(
            value,
            retainOverflow ? x : x + 3D,
            retainOverflow ? y : y + 3D,
            retainOverflow ? width : width - 6D,
            retainOverflow ? height : height - 6D,
            color,
            alignment,
            overflowBehavior);
    }

    /// <summary>
    /// Draws a single anchored text line with optional bold, italic, alignment, and rotation.
    /// </summary>
    /// <param name="text">Text to render.</param>
    /// <param name="anchorX">Horizontal anchor. Left, center, or right interpretation depends on <paramref name="alignment"/>.</param>
    /// <param name="top">Top coordinate of the text line box.</param>
    /// <param name="height">Text line height in canvas pixels.</param>
    /// <param name="color">Text color.</param>
    /// <param name="bold">Whether to simulate bold rendering.</param>
    /// <param name="italic">Whether to simulate italic rendering.</param>
    /// <param name="alignment">Horizontal alignment relative to <paramref name="anchorX"/>.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="underline">Whether to draw an underline using the measured text width.</param>
    /// <param name="strikethrough">Whether to draw a strikethrough using the measured text width.</param>
    /// <param name="fontFamily">Requested font family fallback list.</param>
    /// <param name="flipHorizontal">Whether to mirror the rendered line horizontally around the rotation center before rotation.</param>
    /// <param name="flipVertical">Whether to mirror the rendered line vertically around the rotation center before rotation.</param>
    public void DrawTextLine(
        string? text,
        double anchorX,
        double top,
        double height,
        OfficeColor color,
        bool bold = false,
        bool italic = false,
        OfficeTextAlignment alignment = OfficeTextAlignment.Center,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        bool underline = false,
        bool strikethrough = false,
        string? fontFamily = null,
        bool flipHorizontal = false,
        bool flipVertical = false) {
        if (string.IsNullOrEmpty(text) || color.A == 0 || height <= 0D) {
            return;
        }

        string value = text!;
        double fontHeight = Math.Max(1D, height);
        OfficeFontStyle fontStyle = (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        OfficeTrueTypeFont? font = ResolveTextFont(value, fontFamily, fontStyle, out OfficeFontStyle resolvedStyle);
        bool simulateBold = bold && (resolvedStyle & OfficeFontStyle.Bold) != OfficeFontStyle.Bold;
        bool simulateItalic = italic && (resolvedStyle & OfficeFontStyle.Italic) != OfficeFontStyle.Italic;
        double width = MeasureText(value, fontHeight, fontFamily, fontStyle);
        double x = ResolveAnchoredTextX(anchorX, width, alignment);
        double rotationRadians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        if (font != null) {
            double bottom = top + fontHeight;
            IReadOnlyList<List<OfficePoint>> contours = TransformTextContours(
                GetResolvedTextContours(value, font, x, top, fontHeight),
                bottom,
                simulateItalic,
                rotationRadians,
                rotationCenterX,
                rotationCenterY,
                flipHorizontal,
                flipVertical);
            FillContours(contours, color, OfficeFillRule.EvenOdd);
            if (simulateBold) {
                contours = TransformTextContours(
                    GetResolvedTextContours(value, font, x + Math.Max(1D, fontHeight / 22D), top, fontHeight),
                    bottom,
                    simulateItalic,
                    rotationRadians,
                    rotationCenterX,
                    rotationCenterY,
                    flipHorizontal,
                    flipVertical);
                FillContours(contours, color, OfficeFillRule.EvenOdd);
            }

            DrawTextLineDecorations(x, width, top, fontHeight, color, rotationRadians, rotationCenterX, rotationCenterY, underline, strikethrough, flipHorizontal, flipVertical);
            return;
        }

        DrawStrokeText(value, anchorX, top + (fontHeight / 2D), fontHeight, color, bold, italic, alignment, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
        DrawTextLineDecorations(x, width, top, fontHeight, color, rotationRadians, rotationCenterX, rotationCenterY, underline, strikethrough, flipHorizontal, flipVertical);
    }

    /// <summary>
    /// Draws a single text line through an arbitrary affine transform.
    /// </summary>
    /// <param name="text">Text to draw.</param>
    /// <param name="anchorX">Untransformed text anchor X coordinate.</param>
    /// <param name="top">Untransformed top coordinate.</param>
    /// <param name="height">Untransformed text height.</param>
    /// <param name="color">Text fill color.</param>
    /// <param name="transform">Affine transform applied to text contours.</param>
    /// <param name="bold">Whether to draw a bold approximation.</param>
    /// <param name="italic">Whether to skew text contours before applying the transform.</param>
    /// <param name="alignment">Anchor alignment.</param>
    /// <param name="underline">Whether to draw an underline.</param>
    /// <param name="strikethrough">Whether to draw a strikethrough.</param>
    /// <param name="fontFamily">Requested font family fallback list.</param>
    public void DrawTextLineTransformed(
        string? text,
        double anchorX,
        double top,
        double height,
        OfficeColor color,
        OfficeTransform transform,
        bool bold = false,
        bool italic = false,
        OfficeTextAlignment alignment = OfficeTextAlignment.Center,
        bool underline = false,
        bool strikethrough = false,
        string? fontFamily = null) {
        if (string.IsNullOrEmpty(text) || color.A == 0 || height <= 0D) {
            return;
        }

        string value = text!;
        double fontHeight = Math.Max(1D, height);
        OfficeFontStyle fontStyle = (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        OfficeTrueTypeFont? font = ResolveTextFont(value, fontFamily, fontStyle, out OfficeFontStyle resolvedStyle);
        bool simulateBold = bold && (resolvedStyle & OfficeFontStyle.Bold) != OfficeFontStyle.Bold;
        bool simulateItalic = italic && (resolvedStyle & OfficeFontStyle.Italic) != OfficeFontStyle.Italic;
        double width = MeasureText(value, fontHeight, fontFamily, fontStyle);
        double x = ResolveAnchoredTextX(anchorX, width, alignment);
        if (font != null) {
            IReadOnlyList<List<OfficePoint>> contours = TransformTextContours(
                GetResolvedTextContours(value, font, x, top, fontHeight),
                top + fontHeight,
                simulateItalic,
                transform);
            FillContours(contours, color, OfficeFillRule.EvenOdd);
            if (simulateBold) {
                contours = TransformTextContours(
                    GetResolvedTextContours(value, font, x + Math.Max(1D, fontHeight / 22D), top, fontHeight),
                    top + fontHeight,
                    simulateItalic,
                    transform);
                FillContours(contours, color, OfficeFillRule.EvenOdd);
            }

            DrawAffineTextLineDecorations(x, width, top, fontHeight, color, transform, underline, strikethrough);
            return;
        }

        DrawAffineStrokeText(value, anchorX, top + (fontHeight / 2D), fontHeight, color, bold, italic, alignment, transform);
        DrawAffineTextLineDecorations(x, width, top, fontHeight, color, transform, underline, strikethrough);
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
        bool underline,
        bool strikethrough,
        bool flipHorizontal,
        bool flipVertical) {
        if (width <= 0D || color.A == 0) {
            return;
        }

        if (underline) {
            DrawTransformedTextDecorationLine(x, width, top + (fontHeight * 0.86D), color, fontHeight, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
        }

        if (strikethrough) {
            DrawTransformedTextDecorationLine(x, width, top + (fontHeight * 0.52D), color, fontHeight, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
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

    private void DrawFallbackText(string text, double x, double y, double width, double height, OfficeColor color, OfficeTextAlignment alignment, OfficeTextOverflowBehavior overflowBehavior) {
        if (string.IsNullOrEmpty(text) || color.A == 0 || width <= 0D || height <= 0D) {
            return;
        }

        string value = text;
        double fontHeight = Math.Max(1D, height);
        if (overflowBehavior == OfficeTextOverflowBehavior.Ellipsis) {
            while (MeasureStrokeText(value, fontHeight) > width && value.Length > 0) {
                value = OfficeTextElements.RemoveLast(value);
            }
        }

        double anchorX = alignment == OfficeTextAlignment.Right
            ? x + width
            : alignment == OfficeTextAlignment.Center ? x + (width / 2D) : x;
        DrawStrokeText(value, anchorX, y + (fontHeight / 2D), fontHeight, color, false, false, alignment, 0D, x, y, flipHorizontal: false, flipVertical: false);
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
                contour[j] = new OfficePoint(point.X + ((baseY - point.Y) * 0.18D), point.Y);
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

    private OfficeTrueTypeFont? ResolveTextFont(string? text, string? fontFamily, OfficeFontStyle style = OfficeFontStyle.Regular) =>
        ResolveTextFont(text, fontFamily, style, out _);

    private OfficeTrueTypeFont? ResolveTextFont(string? text, string? fontFamily, OfficeFontStyle style, out OfficeFontStyle resolvedStyle) {
        resolvedStyle = OfficeFontStyle.Regular;
        if (_fonts != null) {
            OfficeTrueTypeFont? scoped = string.IsNullOrEmpty(text)
                ? _fonts.Resolve(fontFamily, style, out resolvedStyle)
                : _fonts.ResolveForText(text!, fontFamily, style, out resolvedStyle);
            if (scoped != null) {
                return scoped;
            }
        }

        if (string.IsNullOrWhiteSpace(fontFamily)) {
            return _font;
        }

        return OfficeTrueTypeFont.TryLoadFontFamily(fontFamily) ?? _font;
    }

    private readonly struct TextMeasurementKey : IEquatable<TextMeasurementKey> {
        internal TextMeasurementKey(string text, double fontSize, string? fontFamily, OfficeFontStyle style) {
            Text = text;
            FontSize = fontSize;
            FontFamily = fontFamily ?? string.Empty;
            Style = OfficeFontFace.NormalizeStyle(style);
        }

        private string Text { get; }
        private double FontSize { get; }
        private string FontFamily { get; }
        private OfficeFontStyle Style { get; }

        public bool Equals(TextMeasurementKey other) =>
            FontSize.Equals(other.FontSize) &&
            string.Equals(Text, other.Text, StringComparison.Ordinal) &&
            string.Equals(FontFamily, other.FontFamily, StringComparison.Ordinal) &&
            Style == other.Style;

        public override bool Equals(object? obj) =>
            obj is TextMeasurementKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = (Text != null ? StringComparer.Ordinal.GetHashCode(Text) : 0);
                hash = (hash * 397) ^ FontSize.GetHashCode();
                hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(FontFamily);
                hash = (hash * 397) ^ Style.GetHashCode();
                return hash;
            }
        }
    }

    private void DrawStrokeText(
        string text,
        double anchorX,
        double centerY,
        double height,
        OfficeColor color,
        bool bold,
        bool italic,
        OfficeTextAlignment alignment,
        double rotationRadians,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical) {
        if (string.IsNullOrEmpty(text) || color.A == 0 || height <= 0D) {
            return;
        }

        double cell = Math.Max(1D, height / 7D);
        double gap = cell * 0.9D;
        double width = MeasureStrokeText(text, height);
        double x = ResolveAnchoredTextX(anchorX, width, alignment);
        double top = centerY - (height / 2D);
        double bottom = top + Math.Max(1D, height);
        foreach (char c in text) {
            DrawStrokeGlyph(c, x, top, cell, color, bold, italic, bottom, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
            x += (GlyphWidth(c) * cell) + gap;
        }
    }

    private void DrawAffineStrokeText(
        string text,
        double anchorX,
        double centerY,
        double height,
        OfficeColor color,
        bool bold,
        bool italic,
        OfficeTextAlignment alignment,
        OfficeTransform transform) {
        if (string.IsNullOrEmpty(text) || color.A == 0 || height <= 0D) {
            return;
        }

        double cell = Math.Max(1D, height / 7D);
        double gap = cell * 0.9D;
        double width = MeasureStrokeText(text, height);
        double x = ResolveAnchoredTextX(anchorX, width, alignment);
        double top = centerY - (height / 2D);
        double bottom = top + Math.Max(1D, height);
        double strokeScale = GetAffineStrokeScale(transform);
        foreach (char c in text) {
            DrawAffineStrokeGlyph(c, x, top, cell, color, bold, italic, bottom, transform, strokeScale);
            x += (GlyphWidth(c) * cell) + gap;
        }
    }

    private void DrawStrokeGlyph(
        char c,
        double x,
        double y,
        double cell,
        OfficeColor color,
        bool bold,
        bool italic,
        double bottom,
        double rotationRadians,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical) {
        string[] rows = GlyphRows(c);
        double strokeWidth = Math.Max(1D, bold ? cell * 0.38D : cell * 0.26D);
        for (int row = 0; row < rows.Length; row++) {
            string bits = rows[row];
            for (int col = 0; col < bits.Length; col++) {
                if (bits[col] != '1') {
                    continue;
                }

                OfficePoint current = TransformTextPoint(GlyphPoint(x, y, cell, col, row), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                bool connected = false;
                if (col + 1 < bits.Length && bits[col + 1] == '1') {
                    OfficePoint nextPoint = TransformTextPoint(GlyphPoint(x, y, cell, col + 1, row), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                    DrawLine(current.X, current.Y, nextPoint.X, nextPoint.Y, color, strokeWidth);
                    connected = true;
                }

                if (row + 1 < rows.Length) {
                    string next = rows[row + 1];
                    if (col < next.Length && next[col] == '1') {
                        OfficePoint nextPoint = TransformTextPoint(GlyphPoint(x, y, cell, col, row + 1), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                        DrawLine(current.X, current.Y, nextPoint.X, nextPoint.Y, color, strokeWidth);
                        connected = true;
                    }

                    if (col > 0 && col - 1 < next.Length && next[col - 1] == '1') {
                        OfficePoint nextPoint = TransformTextPoint(GlyphPoint(x, y, cell, col - 1, row + 1), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                        DrawLine(current.X, current.Y, nextPoint.X, nextPoint.Y, color, strokeWidth);
                        connected = true;
                    }

                    if (col + 1 < next.Length && next[col + 1] == '1') {
                        OfficePoint nextPoint = TransformTextPoint(GlyphPoint(x, y, cell, col + 1, row + 1), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                        DrawLine(current.X, current.Y, nextPoint.X, nextPoint.Y, color, strokeWidth);
                        connected = true;
                    }
                }

                if (!connected) {
                    DrawEllipse(current.X, current.Y, strokeWidth / 2D, strokeWidth / 2D, color, OfficeColor.Transparent, 0D);
                }
            }
        }
    }

    private void DrawAffineStrokeGlyph(
        char c,
        double x,
        double y,
        double cell,
        OfficeColor color,
        bool bold,
        bool italic,
        double bottom,
        OfficeTransform transform,
        double strokeScale) {
        string[] rows = GlyphRows(c);
        double strokeWidth = Math.Max(1D, (bold ? cell * 0.38D : cell * 0.26D) * strokeScale);
        for (int row = 0; row < rows.Length; row++) {
            string bits = rows[row];
            for (int col = 0; col < bits.Length; col++) {
                if (bits[col] != '1') {
                    continue;
                }

                OfficePoint current = TransformAffineTextPoint(GlyphPoint(x, y, cell, col, row), bottom, italic, transform);
                bool connected = false;
                if (col + 1 < bits.Length && bits[col + 1] == '1') {
                    OfficePoint nextPoint = TransformAffineTextPoint(GlyphPoint(x, y, cell, col + 1, row), bottom, italic, transform);
                    DrawLine(current.X, current.Y, nextPoint.X, nextPoint.Y, color, strokeWidth);
                    connected = true;
                }

                if (row + 1 < rows.Length) {
                    string next = rows[row + 1];
                    if (col < next.Length && next[col] == '1') {
                        OfficePoint nextPoint = TransformAffineTextPoint(GlyphPoint(x, y, cell, col, row + 1), bottom, italic, transform);
                        DrawLine(current.X, current.Y, nextPoint.X, nextPoint.Y, color, strokeWidth);
                        connected = true;
                    }

                    if (col > 0 && col - 1 < next.Length && next[col - 1] == '1') {
                        OfficePoint nextPoint = TransformAffineTextPoint(GlyphPoint(x, y, cell, col - 1, row + 1), bottom, italic, transform);
                        DrawLine(current.X, current.Y, nextPoint.X, nextPoint.Y, color, strokeWidth);
                        connected = true;
                    }

                    if (col + 1 < next.Length && next[col + 1] == '1') {
                        OfficePoint nextPoint = TransformAffineTextPoint(GlyphPoint(x, y, cell, col + 1, row + 1), bottom, italic, transform);
                        DrawLine(current.X, current.Y, nextPoint.X, nextPoint.Y, color, strokeWidth);
                        connected = true;
                    }
                }

                if (!connected) {
                    DrawEllipse(current.X, current.Y, strokeWidth / 2D, strokeWidth / 2D, color, OfficeColor.Transparent, 0D);
                }
            }
        }
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

    private static IReadOnlyList<List<OfficePoint>> TransformTextContours(IReadOnlyList<List<OfficePoint>> contours, double bottom, bool italic, double rotationRadians, double rotationCenterX, double rotationCenterY, bool flipHorizontal, bool flipVertical) {
        if ((!italic && Math.Abs(rotationRadians) < TextRotationEpsilon && !flipHorizontal && !flipVertical) || contours.Count == 0) {
            return contours;
        }

        List<List<OfficePoint>> transformed = new(contours.Count);
        foreach (List<OfficePoint> contour in contours) {
            List<OfficePoint> points = new(contour.Count);
            foreach (OfficePoint point in contour) {
                points.Add(TransformTextPoint(point, bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical));
            }

            transformed.Add(points);
        }

        return transformed;
    }

    private static IReadOnlyList<List<OfficePoint>> TransformTextContours(IReadOnlyList<List<OfficePoint>> contours, double bottom, bool italic, OfficeTransform transform) {
        if ((!italic && transform == OfficeTransform.Identity) || contours.Count == 0) {
            return contours;
        }

        List<List<OfficePoint>> transformed = new(contours.Count);
        foreach (List<OfficePoint> contour in contours) {
            List<OfficePoint> points = new(contour.Count);
            foreach (OfficePoint point in contour) {
                OfficePoint skewed = italic ? new OfficePoint(point.X + ((bottom - point.Y) * ItalicShear), point.Y) : point;
                points.Add(transform.TransformPoint(skewed));
            }

            transformed.Add(points);
        }

        return transformed;
    }

    private static OfficePoint TransformTextPoint(OfficePoint point, double bottom, bool italic, double rotationRadians, double rotationCenterX, double rotationCenterY, bool flipHorizontal, bool flipVertical) {
        if (!italic && Math.Abs(rotationRadians) < TextRotationEpsilon && !flipHorizontal && !flipVertical) {
            return point;
        }

        OfficePoint skewed = italic ? new OfficePoint(point.X + ((bottom - point.Y) * ItalicShear), point.Y) : point;
        return TransformFramePoint(skewed, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
    }

    private static OfficePoint TransformFramePoint(OfficePoint point, double rotationRadians, double centerX, double centerY, bool flipHorizontal, bool flipVertical) {
        OfficePoint transformed = point;
        if (flipHorizontal) {
            transformed = new OfficePoint((2D * centerX) - transformed.X, transformed.Y);
        }

        if (flipVertical) {
            transformed = new OfficePoint(transformed.X, (2D * centerY) - transformed.Y);
        }

        return Math.Abs(rotationRadians) < TextRotationEpsilon
            ? transformed
            : OfficeGeometry.RotatePoint(transformed, centerX, centerY, rotationRadians);
    }

    private static OfficePoint TransformAffineTextPoint(OfficePoint point, double bottom, bool italic, OfficeTransform transform) {
        OfficePoint skewed = italic ? new OfficePoint(point.X + ((bottom - point.Y) * ItalicShear), point.Y) : point;
        return transform.TransformPoint(skewed);
    }

    private static double GetAffineStrokeScale(OfficeTransform transform) {
        double xScale = Math.Sqrt((transform.M11 * transform.M11) + (transform.M12 * transform.M12));
        double yScale = Math.Sqrt((transform.M21 * transform.M21) + (transform.M22 * transform.M22));
        double scale = Math.Max(xScale, yScale);
        return !double.IsNaN(scale) && !double.IsInfinity(scale) && scale > 0D ? scale : 1D;
    }

    private static OfficePoint GlyphPoint(double x, double y, double cell, int col, int row) {
        return new OfficePoint(x + ((col + 0.5D) * cell), y + ((row + 0.5D) * cell));
    }

    private static double MeasureStrokeText(string text, double height) {
        if (string.IsNullOrEmpty(text)) {
            return 0D;
        }

        double cell = Math.Max(1D, height / 7D);
        double gap = cell * 0.9D;
        double width = 0D;
        foreach (char c in text) {
            width += (GlyphWidth(c) * cell) + gap;
        }

        return width > 0D ? width - gap : 0D;
    }

    private static int GlyphWidth(char c) => c == ' ' ? 3 : 5;

    private static string[] GlyphRows(char c) {
        switch (char.ToUpperInvariant(c)) {
            case 'A': return new[] { "01110", "10001", "10001", "11111", "10001", "10001", "10001" };
            case 'B': return new[] { "11110", "10001", "10001", "11110", "10001", "10001", "11110" };
            case 'C': return new[] { "01111", "10000", "10000", "10000", "10000", "10000", "01111" };
            case 'D': return new[] { "11110", "10001", "10001", "10001", "10001", "10001", "11110" };
            case 'E': return new[] { "11111", "10000", "10000", "11110", "10000", "10000", "11111" };
            case 'F': return new[] { "11111", "10000", "10000", "11110", "10000", "10000", "10000" };
            case 'G': return new[] { "01111", "10000", "10000", "10111", "10001", "10001", "01110" };
            case 'H': return new[] { "10001", "10001", "10001", "11111", "10001", "10001", "10001" };
            case 'I': return new[] { "11111", "00100", "00100", "00100", "00100", "00100", "11111" };
            case 'J': return new[] { "00111", "00010", "00010", "00010", "10010", "10010", "01100" };
            case 'K': return new[] { "10001", "10010", "10100", "11000", "10100", "10010", "10001" };
            case 'L': return new[] { "10000", "10000", "10000", "10000", "10000", "10000", "11111" };
            case 'M': return new[] { "10001", "11011", "10101", "10101", "10001", "10001", "10001" };
            case 'N': return new[] { "10001", "11001", "10101", "10011", "10001", "10001", "10001" };
            case 'O': return new[] { "01110", "10001", "10001", "10001", "10001", "10001", "01110" };
            case 'P': return new[] { "11110", "10001", "10001", "11110", "10000", "10000", "10000" };
            case 'Q': return new[] { "01110", "10001", "10001", "10001", "10101", "10010", "01101" };
            case 'R': return new[] { "11110", "10001", "10001", "11110", "10100", "10010", "10001" };
            case 'S': return new[] { "01111", "10000", "10000", "01110", "00001", "00001", "11110" };
            case 'T': return new[] { "11111", "00100", "00100", "00100", "00100", "00100", "00100" };
            case 'U': return new[] { "10001", "10001", "10001", "10001", "10001", "10001", "01110" };
            case 'V': return new[] { "10001", "10001", "10001", "10001", "10001", "01010", "00100" };
            case 'W': return new[] { "10001", "10001", "10001", "10101", "10101", "10101", "01010" };
            case 'X': return new[] { "10001", "10001", "01010", "00100", "01010", "10001", "10001" };
            case 'Y': return new[] { "10001", "10001", "01010", "00100", "00100", "00100", "00100" };
            case 'Z': return new[] { "11111", "00001", "00010", "00100", "01000", "10000", "11111" };
            case '0': return new[] { "01110", "10001", "10011", "10101", "11001", "10001", "01110" };
            case '1': return new[] { "00100", "01100", "00100", "00100", "00100", "00100", "01110" };
            case '2': return new[] { "01110", "10001", "00001", "00010", "00100", "01000", "11111" };
            case '3': return new[] { "11110", "00001", "00001", "01110", "00001", "00001", "11110" };
            case '4': return new[] { "00010", "00110", "01010", "10010", "11111", "00010", "00010" };
            case '5': return new[] { "11111", "10000", "10000", "11110", "00001", "00001", "11110" };
            case '6': return new[] { "01110", "10000", "10000", "11110", "10001", "10001", "01110" };
            case '7': return new[] { "11111", "00001", "00010", "00100", "01000", "01000", "01000" };
            case '8': return new[] { "01110", "10001", "10001", "01110", "10001", "10001", "01110" };
            case '9': return new[] { "01110", "10001", "10001", "01111", "00001", "00001", "01110" };
            case '-': return new[] { "00000", "00000", "00000", "11111", "00000", "00000", "00000" };
            case '_': return new[] { "00000", "00000", "00000", "00000", "00000", "00000", "11111" };
            case '+': return new[] { "00000", "00100", "00100", "11111", "00100", "00100", "00000" };
            case '=': return new[] { "00000", "00000", "11111", "00000", "11111", "00000", "00000" };
            case '/': return new[] { "00001", "00001", "00010", "00100", "01000", "10000", "10000" };
            case '\\': return new[] { "10000", "10000", "01000", "00100", "00010", "00001", "00001" };
            case '.': return new[] { "00000", "00000", "00000", "00000", "00000", "01100", "01100" };
            case ',': return new[] { "00000", "00000", "00000", "00000", "00000", "01100", "01000" };
            case ':': return new[] { "00000", "01100", "01100", "00000", "01100", "01100", "00000" };
            case ';': return new[] { "00000", "01100", "01100", "00000", "01100", "01000", "10000" };
            case '!': return new[] { "00100", "00100", "00100", "00100", "00100", "00000", "00100" };
            case '?': return new[] { "01110", "10001", "00001", "00010", "00100", "00000", "00100" };
            case '&': return new[] { "01100", "10010", "10100", "01000", "10101", "10010", "01101" };
            case '%': return new[] { "11001", "11010", "00010", "00100", "01000", "01011", "10011" };
            case '#': return new[] { "01010", "01010", "11111", "01010", "11111", "01010", "01010" };
            case '(': return new[] { "00010", "00100", "01000", "01000", "01000", "00100", "00010" };
            case ')': return new[] { "01000", "00100", "00010", "00010", "00010", "00100", "01000" };
            case '[': return new[] { "01110", "01000", "01000", "01000", "01000", "01000", "01110" };
            case ']': return new[] { "01110", "00010", "00010", "00010", "00010", "00010", "01110" };
            case '<': return new[] { "00010", "00100", "01000", "10000", "01000", "00100", "00010" };
            case '>': return new[] { "01000", "00100", "00010", "00001", "00010", "00100", "01000" };
            case '|': return new[] { "00100", "00100", "00100", "00100", "00100", "00100", "00100" };
            case '\'': return new[] { "01100", "00100", "01000", "00000", "00000", "00000", "00000" };
            case '"': return new[] { "01010", "01010", "01010", "00000", "00000", "00000", "00000" };
            case ' ': return new[] { "000", "000", "000", "000", "000", "000", "000" };
            default: return new[] { "11111", "10001", "00001", "00010", "00100", "00000", "00100" };
        }
    }
}
