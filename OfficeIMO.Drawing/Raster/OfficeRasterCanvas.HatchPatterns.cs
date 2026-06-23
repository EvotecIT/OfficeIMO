using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>
    /// Draws a shared hatch pattern over the supplied rectangle.
    /// </summary>
    /// <remarks>Diagonal hatches may extend outside the rectangle; callers can clip when strict bounds are required.</remarks>
    public void DrawHatchPatternRectangle(double x, double y, double width, double height, OfficeColor color, double step, double lineWidth, OfficeHatchPatternKind pattern) {
        if (color.A == 0 || width <= 0D || height <= 0D || step <= 0D || lineWidth <= 0D) {
            return;
        }

        switch (pattern) {
            case OfficeHatchPatternKind.Horizontal:
                DrawHorizontalHatchPattern(x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.Vertical:
                DrawVerticalHatchPattern(x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.DiagonalDown:
                DrawDiagonalDownHatchPattern(x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.DiagonalUp:
                DrawDiagonalUpHatchPattern(x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.Grid:
                DrawHorizontalHatchPattern(x, y, width, height, color, step, lineWidth);
                DrawVerticalHatchPattern(x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.Trellis:
                DrawDiagonalDownHatchPattern(x, y, width, height, color, step, lineWidth);
                DrawDiagonalUpHatchPattern(x, y, width, height, color, step, lineWidth);
                break;
            default:
                DrawDottedHatchPattern(x, y, width, height, color, step, lineWidth);
                break;
        }
    }

    private void DrawHorizontalHatchPattern(double x, double y, double width, double height, OfficeColor color, double step, double lineWidth) {
        for (double yy = y + step; yy < y + height; yy += step) {
            DrawStyledLine(x, yy, x + width, yy, color, lineWidth);
        }
    }

    private void DrawVerticalHatchPattern(double x, double y, double width, double height, OfficeColor color, double step, double lineWidth) {
        for (double xx = x + step; xx < x + width; xx += step) {
            DrawStyledLine(xx, y, xx, y + height, color, lineWidth);
        }
    }

    private void DrawDiagonalDownHatchPattern(double x, double y, double width, double height, OfficeColor color, double step, double lineWidth) {
        for (double xx = x - height; xx < x + width; xx += step) {
            DrawStyledLine(xx, y, xx + height, y + height, color, lineWidth);
        }
    }

    private void DrawDiagonalUpHatchPattern(double x, double y, double width, double height, OfficeColor color, double step, double lineWidth) {
        for (double xx = x; xx < x + width + height; xx += step) {
            DrawStyledLine(xx, y, xx - height, y + height, color, lineWidth);
        }
    }

    private void DrawDottedHatchPattern(double x, double y, double width, double height, OfficeColor color, double step, double dotSize) {
        double size = Math.Max(1D, dotSize);
        for (double yy = y + step / 2D; yy < y + height; yy += step) {
            for (double xx = x + step / 2D; xx < x + width; xx += step) {
                FillRectangle(xx, yy, size, size, color);
            }
        }
    }
}
