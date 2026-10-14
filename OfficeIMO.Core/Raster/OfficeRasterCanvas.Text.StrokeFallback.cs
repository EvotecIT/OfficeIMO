using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private void DrawFallbackText(
        string text, double x, double y, double width, double height, OfficeColor color,
        double size, OfficeTextAlignment alignment, OfficeFontStyle style,
        OfficeTextOverflowBehavior overflowBehavior, double? textAdvanceWidth,
        OfficeTextDecorationStyle underlineStyle, OfficeTextDecorationStyle strikethroughStyle,
        OfficeColor? decorationColor, double? baselineFontSize) {
        bool retainOverflow = overflowBehavior == OfficeTextOverflowBehavior.Clip;
        double availableWidth = Math.Max(1D, retainOverflow ? width : width - 6D);
        string value = text;
        double measured = MeasureStrokeText(value, size);
        if (measured <= 0D) return;
        double advance = textAdvanceWidth.HasValue && string.Equals(value, text, StringComparison.Ordinal)
            ? textAdvanceWidth.Value : measured;
        double textX = ResolveTextX(retainOverflow ? x : x + 3D, availableWidth, advance, alignment);
        double top = y + ResolveStrokeTextTop(size, height, baselineFontSize);
        double horizontalScale = advance / measured;
        var transform = new OfficeTransform(horizontalScale, 0D, 0D, 1D, textX * (1D - horizontalScale), 0D);
        DrawAffineStrokeText(value, textX, top + size / 2D, size, color,
            (style & OfficeFontStyle.Bold) != 0, (style & OfficeFontStyle.Italic) != 0,
            OfficeTextAlignment.Left, transform);
        DrawTextLineDecorations(textX, advance, top, size, decorationColor ?? color, 0D, 0D, 0D,
            underlineStyle != OfficeTextDecorationStyle.None ? underlineStyle :
                (style & OfficeFontStyle.Underline) != 0 ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None,
            strikethroughStyle != OfficeTextDecorationStyle.None ? strikethroughStyle :
                (style & OfficeFontStyle.Strikethrough) != 0 ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None,
            false, false);
    }

    private static double ResolveStrokeTextTop(double size, double height, double? baselineFontSize) =>
        baselineFontSize.HasValue ? baselineFontSize.Value - size * .84D : Math.Max(1D, (height - size) / 2D);

    // Bound the same stroke points, shear and scaled round caps used for painting.
    // Small text keeps the fallback's one-pixel cells, so its ink can exceed one em.
    private static (double Left, double Top, double Right, double Bottom, bool HasInk) MeasureStrokeTextBounds(
        string text, double x, double top, double size, bool bold, bool italic, double horizontalScale = 1D) {
        double cell = Math.Max(1D, size / 7D);
        double gap = cell * .9D;
        double strokeRadius = Math.Max(1D, (bold ? cell * .38D : cell * .26D) * Math.Max(1D, horizontalScale)) / 2D;
        double left = double.PositiveInfinity, right = double.NegativeInfinity;
        double inkTop = double.PositiveInfinity, bottom = double.NegativeInfinity;
        double cursor = x;
        foreach (char c in text) {
            string[] rows = GlyphRows(c);
            for (int row = 0; row < rows.Length; row++) {
                for (int col = 0; col < rows[row].Length; col++) {
                    if (rows[row][col] != '1') continue;
                    OfficePoint point = GlyphPoint(cursor, top, cell, col, row);
                    if (italic) point = new OfficePoint(point.X + ((top + Math.Max(1D, size) - point.Y) * ItalicShear), point.Y);
                    double px = x + (point.X - x) * horizontalScale;
                    left = Math.Min(left, px - strokeRadius); right = Math.Max(right, px + strokeRadius);
                    inkTop = Math.Min(inkTop, point.Y - strokeRadius); bottom = Math.Max(bottom, point.Y + strokeRadius);
                }
            }
            cursor += GlyphWidth(c) * cell + gap;
        }
        return (left, inkTop, right, bottom, !double.IsPositiveInfinity(left));
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
        var strokes = new List<OfficeFlattenedPathContour>();
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
                    strokes.Add(new OfficeFlattenedPathContour(new[] { current, nextPoint }, false));
                    connected = true;
                }

                if (row + 1 < rows.Length) {
                    string next = rows[row + 1];
                    if (col < next.Length && next[col] == '1') {
                        OfficePoint nextPoint = TransformTextPoint(GlyphPoint(x, y, cell, col, row + 1), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                        strokes.Add(new OfficeFlattenedPathContour(new[] { current, nextPoint }, false));
                        connected = true;
                    }

                    if (col > 0 && col - 1 < next.Length && next[col - 1] == '1') {
                        OfficePoint nextPoint = TransformTextPoint(GlyphPoint(x, y, cell, col - 1, row + 1), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                        strokes.Add(new OfficeFlattenedPathContour(new[] { current, nextPoint }, false));
                        connected = true;
                    }

                    if (col + 1 < next.Length && next[col + 1] == '1') {
                        OfficePoint nextPoint = TransformTextPoint(GlyphPoint(x, y, cell, col + 1, row + 1), bottom, italic, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                        strokes.Add(new OfficeFlattenedPathContour(new[] { current, nextPoint }, false));
                        connected = true;
                    }
                }

                if (!connected) {
                    strokes.Add(new OfficeFlattenedPathContour(new[] { current }, false));
                }
            }
        }
        StrokeContours(strokes, strokeWidth, OfficeStrokeLineCap.Round, OfficeStrokeLineJoin.Round, 4D, null, 0D, (_, _) => color);
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
        var strokes = new List<OfficeFlattenedPathContour>();
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
                    strokes.Add(new OfficeFlattenedPathContour(new[] { current, nextPoint }, false));
                    connected = true;
                }

                if (row + 1 < rows.Length) {
                    string next = rows[row + 1];
                    if (col < next.Length && next[col] == '1') {
                        OfficePoint nextPoint = TransformAffineTextPoint(GlyphPoint(x, y, cell, col, row + 1), bottom, italic, transform);
                        strokes.Add(new OfficeFlattenedPathContour(new[] { current, nextPoint }, false));
                        connected = true;
                    }

                    if (col > 0 && col - 1 < next.Length && next[col - 1] == '1') {
                        OfficePoint nextPoint = TransformAffineTextPoint(GlyphPoint(x, y, cell, col - 1, row + 1), bottom, italic, transform);
                        strokes.Add(new OfficeFlattenedPathContour(new[] { current, nextPoint }, false));
                        connected = true;
                    }

                    if (col + 1 < next.Length && next[col + 1] == '1') {
                        OfficePoint nextPoint = TransformAffineTextPoint(GlyphPoint(x, y, cell, col + 1, row + 1), bottom, italic, transform);
                        strokes.Add(new OfficeFlattenedPathContour(new[] { current, nextPoint }, false));
                        connected = true;
                    }
                }

                if (!connected) {
                    strokes.Add(new OfficeFlattenedPathContour(new[] { current }, false));
                }
            }
        }
        StrokeContours(strokes, strokeWidth, OfficeStrokeLineCap.Round, OfficeStrokeLineJoin.Round, 4D, null, 0D, (_, _) => color);
    }

}
