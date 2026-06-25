using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free renderer for compact conditional-formatting icon shapes.
/// </summary>
public static class OfficeConditionalIconRenderer {
    /// <summary>
    /// Draws a conditional-formatting icon on a raster canvas.
    /// </summary>
    public static void DrawRaster(OfficeRasterCanvas canvas, double x, double y, double size, OfficeConditionalIconKind kind, double scale = 1D) {
        if (canvas == null) {
            throw new ArgumentNullException(nameof(canvas));
        }

        if (size <= 0D) {
            return;
        }

        OfficeColor fill = GetFillColor(kind);
        OfficeColor stroke = GetStrokeColor(kind);
        double strokeWidth = Math.Max(1D, scale);
        switch (kind) {
            case OfficeConditionalIconKind.GreenCheck:
                canvas.DrawLine(x + size * 0.22D, y + size * 0.54D, x + size * 0.42D, y + size * 0.74D, fill, Math.Max(2D, size * 0.14D));
                canvas.DrawLine(x + size * 0.42D, y + size * 0.74D, x + size * 0.80D, y + size * 0.28D, fill, Math.Max(2D, size * 0.14D));
                break;
            case OfficeConditionalIconKind.YellowExclamation:
                DrawRasterCircle(canvas, x, y, size, fill, stroke, strokeWidth);
                canvas.DrawLine(x + size * 0.5D, y + size * 0.25D, x + size * 0.5D, y + size * 0.60D, OfficeColor.White, Math.Max(2D, size * 0.12D));
                canvas.FillEllipse(x + size * 0.44D, y + size * 0.72D, size * 0.12D, size * 0.12D, OfficeColor.White);
                break;
            case OfficeConditionalIconKind.RedCross:
                canvas.DrawLine(x + size * 0.25D, y + size * 0.25D, x + size * 0.75D, y + size * 0.75D, fill, Math.Max(2D, size * 0.14D));
                canvas.DrawLine(x + size * 0.75D, y + size * 0.25D, x + size * 0.25D, y + size * 0.75D, fill, Math.Max(2D, size * 0.14D));
                break;
            case OfficeConditionalIconKind.GreenCircle:
            case OfficeConditionalIconKind.LightGreenCircle:
            case OfficeConditionalIconKind.YellowCircle:
            case OfficeConditionalIconKind.OrangeCircle:
            case OfficeConditionalIconKind.RedCircle:
                DrawRasterCircle(canvas, x, y, size, fill, stroke, strokeWidth);
                break;
            default:
                IReadOnlyList<OfficePoint> points = CreateArrowPoints(x, y, size, kind);
                canvas.FillPolygon(points, fill);
                canvas.DrawPolygon(points, stroke, strokeWidth);
                break;
        }
    }

    /// <summary>
    /// Appends SVG markup for a conditional-formatting icon.
    /// </summary>
    public static StringBuilder AppendSvg(StringBuilder builder, double x, double y, double size, OfficeConditionalIconKind kind, double scale = 1D) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (size <= 0D) {
            return builder;
        }

        OfficeColor fill = GetFillColor(kind);
        OfficeColor stroke = GetStrokeColor(kind);
        double strokeWidth = Math.Max(1D, scale);
        switch (kind) {
            case OfficeConditionalIconKind.GreenCheck:
                AppendSvgLine(builder, x + size * 0.22D, y + size * 0.54D, x + size * 0.42D, y + size * 0.74D, fill, size * 0.14D);
                AppendSvgLine(builder, x + size * 0.42D, y + size * 0.74D, x + size * 0.80D, y + size * 0.28D, fill, size * 0.14D);
                break;
            case OfficeConditionalIconKind.YellowExclamation:
                AppendSvgCircle(builder, x, y, size, fill, stroke, Math.Max(1D, size / 14D));
                AppendSvgLine(builder, x + size * 0.5D, y + size * 0.25D, x + size * 0.5D, y + size * 0.60D, OfficeColor.White, size * 0.12D);
                builder.AppendCircleElement(x + size * 0.5D, y + size * 0.78D, size * 0.06D, OfficeColor.White);
                break;
            case OfficeConditionalIconKind.RedCross:
                AppendSvgLine(builder, x + size * 0.25D, y + size * 0.25D, x + size * 0.75D, y + size * 0.75D, fill, size * 0.14D);
                AppendSvgLine(builder, x + size * 0.75D, y + size * 0.25D, x + size * 0.25D, y + size * 0.75D, fill, size * 0.14D);
                break;
            case OfficeConditionalIconKind.GreenCircle:
            case OfficeConditionalIconKind.LightGreenCircle:
            case OfficeConditionalIconKind.YellowCircle:
            case OfficeConditionalIconKind.OrangeCircle:
            case OfficeConditionalIconKind.RedCircle:
                AppendSvgCircle(builder, x, y, size, fill, stroke, Math.Max(1D, size / 14D));
                break;
            default:
                IReadOnlyList<OfficePoint> points = CreateArrowPoints(x, y, size, kind);
                var attributes = new StringBuilder()
                    .AppendPaintAttribute("fill", fill)
                    .AppendPaintAttribute("stroke", stroke)
                    .AppendNumberAttribute("stroke-width", strokeWidth)
                    .ToString();
                builder.AppendPathElement(OfficeSvgFormatting.FormatMoveLinePathData(points, closePath: true), attributes);
                break;
        }

        return builder;
    }

    private static void DrawRasterCircle(OfficeRasterCanvas canvas, double x, double y, double size, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        canvas.FillEllipse(x, y, size, size, fill);
        canvas.DrawEllipse(x, y, size, size, stroke, strokeWidth);
    }

    private static void AppendSvgCircle(StringBuilder builder, double x, double y, double size, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        var attributes = new StringBuilder()
            .AppendPaintAttribute("fill", fill)
            .AppendPaintAttribute("stroke", stroke)
            .AppendNumberAttribute("stroke-width", strokeWidth)
            .ToString();
        builder.AppendCircleElement(x + size / 2D, y + size / 2D, size / 2D, attributes);
    }

    private static void AppendSvgLine(StringBuilder builder, double x1, double y1, double x2, double y2, OfficeColor color, double width) {
        builder.AppendLineElement(x1, y1, x2, y2, color, Math.Max(1D, width), OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap.Round);
    }

    private static IReadOnlyList<OfficePoint> CreateArrowPoints(double x, double y, double size, OfficeConditionalIconKind kind) {
        double s = size;
        if (kind == OfficeConditionalIconKind.RedDownArrow || kind == OfficeConditionalIconKind.YellowDownArrow) {
            return new[] {
                new OfficePoint(x + s * 0.36D, y + s * 0.08D),
                new OfficePoint(x + s * 0.64D, y + s * 0.08D),
                new OfficePoint(x + s * 0.64D, y + s * 0.54D),
                new OfficePoint(x + s * 0.86D, y + s * 0.54D),
                new OfficePoint(x + s * 0.50D, y + s * 0.92D),
                new OfficePoint(x + s * 0.14D, y + s * 0.54D),
                new OfficePoint(x + s * 0.36D, y + s * 0.54D)
            };
        }

        if (kind == OfficeConditionalIconKind.YellowSideArrow) {
            return new[] {
                new OfficePoint(x + s * 0.10D, y + s * 0.36D),
                new OfficePoint(x + s * 0.56D, y + s * 0.36D),
                new OfficePoint(x + s * 0.56D, y + s * 0.14D),
                new OfficePoint(x + s * 0.92D, y + s * 0.50D),
                new OfficePoint(x + s * 0.56D, y + s * 0.86D),
                new OfficePoint(x + s * 0.56D, y + s * 0.64D),
                new OfficePoint(x + s * 0.10D, y + s * 0.64D)
            };
        }

        return new[] {
            new OfficePoint(x + s * 0.36D, y + s * 0.92D),
            new OfficePoint(x + s * 0.64D, y + s * 0.92D),
            new OfficePoint(x + s * 0.64D, y + s * 0.46D),
            new OfficePoint(x + s * 0.86D, y + s * 0.46D),
            new OfficePoint(x + s * 0.50D, y + s * 0.08D),
            new OfficePoint(x + s * 0.14D, y + s * 0.46D),
            new OfficePoint(x + s * 0.36D, y + s * 0.46D)
        };
    }

    private static OfficeColor GetFillColor(OfficeConditionalIconKind kind) =>
        kind switch {
            OfficeConditionalIconKind.GreenUpArrow or OfficeConditionalIconKind.GreenCheck or OfficeConditionalIconKind.GreenCircle => OfficeColor.FromRgb(22, 163, 74),
            OfficeConditionalIconKind.LightGreenCircle => OfficeColor.FromRgb(132, 204, 22),
            OfficeConditionalIconKind.YellowUpArrow or OfficeConditionalIconKind.YellowSideArrow or OfficeConditionalIconKind.YellowDownArrow or OfficeConditionalIconKind.YellowExclamation or OfficeConditionalIconKind.YellowCircle => OfficeColor.FromRgb(245, 158, 11),
            OfficeConditionalIconKind.OrangeCircle => OfficeColor.FromRgb(249, 115, 22),
            _ => OfficeColor.FromRgb(220, 38, 38)
        };

    private static OfficeColor GetStrokeColor(OfficeConditionalIconKind kind) =>
        kind switch {
            OfficeConditionalIconKind.GreenUpArrow or OfficeConditionalIconKind.GreenCheck or OfficeConditionalIconKind.GreenCircle => OfficeColor.FromRgb(21, 128, 61),
            OfficeConditionalIconKind.LightGreenCircle => OfficeColor.FromRgb(77, 124, 15),
            OfficeConditionalIconKind.YellowUpArrow or OfficeConditionalIconKind.YellowSideArrow or OfficeConditionalIconKind.YellowDownArrow or OfficeConditionalIconKind.YellowExclamation or OfficeConditionalIconKind.YellowCircle => OfficeColor.FromRgb(180, 83, 9),
            OfficeConditionalIconKind.OrangeCircle => OfficeColor.FromRgb(194, 65, 12),
            _ => OfficeColor.FromRgb(185, 28, 28)
        };
}
