using System;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgFormatting {
    /// <summary>
    /// Appends SVG elements for a shared hatch pattern over the supplied rectangle.
    /// </summary>
    /// <returns>The supplied builder for call chaining.</returns>
    /// <remarks>Diagonal hatches may extend outside the rectangle; callers can clip when strict bounds are required.</remarks>
    public static StringBuilder AppendHatchPatternRectangle(this StringBuilder builder, double x, double y, double width, double height, OfficeColor color, double step, double lineWidth, OfficeHatchPatternKind pattern) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (color.A == 0 || width <= 0D || height <= 0D || step <= 0D || lineWidth <= 0D) {
            return builder;
        }

        if (OfficeStipplePattern.TryCreate(pattern, out OfficeStipplePattern stipplePattern)) {
            AppendSvgStipplePattern(builder, x, y, width, height, color, step, lineWidth, stipplePattern);
            return builder;
        }

        switch (pattern) {
            case OfficeHatchPatternKind.Horizontal:
                AppendSvgHorizontalHatchPattern(builder, x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.Vertical:
                AppendSvgVerticalHatchPattern(builder, x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.DiagonalDown:
                AppendSvgDiagonalDownHatchPattern(builder, x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.DiagonalUp:
                AppendSvgDiagonalUpHatchPattern(builder, x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.Grid:
                AppendSvgHorizontalHatchPattern(builder, x, y, width, height, color, step, lineWidth);
                AppendSvgVerticalHatchPattern(builder, x, y, width, height, color, step, lineWidth);
                break;
            case OfficeHatchPatternKind.Trellis:
                AppendSvgDiagonalDownHatchPattern(builder, x, y, width, height, color, step, lineWidth);
                AppendSvgDiagonalUpHatchPattern(builder, x, y, width, height, color, step, lineWidth);
                break;
            default:
                AppendSvgDottedHatchPattern(builder, x, y, width, height, color, step, lineWidth);
                break;
        }

        return builder;
    }

    private static void AppendSvgHorizontalHatchPattern(StringBuilder builder, double x, double y, double width, double height, OfficeColor color, double step, double lineWidth) {
        for (double yy = y + step; yy < y + height; yy += step) {
            AppendSvgHatchLine(builder, x, yy, x + width, yy, color, lineWidth);
        }
    }

    private static void AppendSvgVerticalHatchPattern(StringBuilder builder, double x, double y, double width, double height, OfficeColor color, double step, double lineWidth) {
        for (double xx = x + step; xx < x + width; xx += step) {
            AppendSvgHatchLine(builder, xx, y, xx, y + height, color, lineWidth);
        }
    }

    private static void AppendSvgDiagonalDownHatchPattern(StringBuilder builder, double x, double y, double width, double height, OfficeColor color, double step, double lineWidth) {
        for (double xx = x - height; xx < x + width; xx += step) {
            AppendSvgHatchLine(builder, xx, y, xx + height, y + height, color, lineWidth);
        }
    }

    private static void AppendSvgDiagonalUpHatchPattern(StringBuilder builder, double x, double y, double width, double height, OfficeColor color, double step, double lineWidth) {
        for (double xx = x; xx < x + width + height; xx += step) {
            AppendSvgHatchLine(builder, xx, y, xx - height, y + height, color, lineWidth);
        }
    }

    private static void AppendSvgDottedHatchPattern(StringBuilder builder, double x, double y, double width, double height, OfficeColor color, double step, double dotSize) {
        double size = Math.Max(1D, dotSize);
        for (double yy = y + step / 2D; yy < y + height; yy += step) {
            for (double xx = x + step / 2D; xx < x + width; xx += step) {
                builder.Append("<rect");
                builder.AppendNumberAttribute("x", xx)
                    .AppendNumberAttribute("y", yy)
                    .AppendNumberAttribute("width", size)
                    .AppendNumberAttribute("height", size)
                    .AppendPaintAttribute("fill", color)
                    .Append("/>");
            }
        }
    }

    private static void AppendSvgStipplePattern(StringBuilder builder, double x, double y, double width, double height, OfficeColor color, double step, double dotSize, OfficeStipplePattern pattern) {
        double size = Math.Max(1D, dotSize);
        double tileSize = Math.Max(step, size * pattern.Size);
        for (double tileY = y; tileY < y + height; tileY += tileSize) {
            for (double tileX = x; tileX < x + width; tileX += tileSize) {
                for (int cellY = 0; cellY < pattern.Size; cellY++) {
                    for (int cellX = 0; cellX < pattern.Size; cellX++) {
                        if (pattern.IsFilled(cellX, cellY)) {
                            builder.Append("<rect");
                            builder.AppendNumberAttribute("x", tileX + cellX * size)
                                .AppendNumberAttribute("y", tileY + cellY * size)
                                .AppendNumberAttribute("width", size)
                                .AppendNumberAttribute("height", size)
                                .AppendPaintAttribute("fill", color)
                                .Append("/>");
                        }
                    }
                }
            }
        }
    }

    private static void AppendSvgHatchLine(StringBuilder builder, double x1, double y1, double x2, double y2, OfficeColor color, double lineWidth) {
        builder.Append("<line");
        builder.AppendNumberAttribute("x1", x1)
            .AppendNumberAttribute("y1", y1)
            .AppendNumberAttribute("x2", x2)
            .AppendNumberAttribute("y2", y2)
            .AppendPaintAttribute("stroke", color)
            .AppendNumberAttribute("stroke-width", lineWidth)
            .Append("/>");
    }
}
