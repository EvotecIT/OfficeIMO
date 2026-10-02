using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
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
