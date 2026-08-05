using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free renderer for compact conditional-formatting icon shapes.
/// </summary>
public static class OfficeConditionalIconRenderer {
    private static readonly OfficeColor IconShadowColor = OfficeColor.FromRgba(15, 23, 42, 42);
    private static readonly OfficeColor IconHighlightColor = OfficeColor.FromRgba(255, 255, 255, 72);

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
                DrawRasterLineShadow(canvas, x + size * 0.22D, y + size * 0.54D, x + size * 0.42D, y + size * 0.74D, size);
                DrawRasterLineShadow(canvas, x + size * 0.42D, y + size * 0.74D, x + size * 0.80D, y + size * 0.28D, size);
                canvas.DrawLine(x + size * 0.22D, y + size * 0.54D, x + size * 0.42D, y + size * 0.74D, fill, Math.Max(2D, size * 0.14D));
                canvas.DrawLine(x + size * 0.42D, y + size * 0.74D, x + size * 0.80D, y + size * 0.28D, fill, Math.Max(2D, size * 0.14D));
                break;
            case OfficeConditionalIconKind.YellowExclamation:
                DrawRasterCircle(canvas, x, y, size, fill, stroke, strokeWidth);
                canvas.DrawLine(x + size * 0.5D, y + size * 0.25D, x + size * 0.5D, y + size * 0.60D, OfficeColor.White, Math.Max(2D, size * 0.12D));
                canvas.FillEllipse(x + size * 0.44D, y + size * 0.72D, size * 0.12D, size * 0.12D, OfficeColor.White);
                break;
            case OfficeConditionalIconKind.RedCross:
                DrawRasterLineShadow(canvas, x + size * 0.25D, y + size * 0.25D, x + size * 0.75D, y + size * 0.75D, size);
                DrawRasterLineShadow(canvas, x + size * 0.75D, y + size * 0.25D, x + size * 0.25D, y + size * 0.75D, size);
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
            case OfficeConditionalIconKind.RatingOne:
            case OfficeConditionalIconKind.RatingTwo:
            case OfficeConditionalIconKind.RatingThree:
            case OfficeConditionalIconKind.RatingFour:
            case OfficeConditionalIconKind.RatingFive:
                DrawRasterRatingStar(canvas, x, y, size, fill, stroke, strokeWidth);
                break;
            case OfficeConditionalIconKind.QuarterEmpty:
            case OfficeConditionalIconKind.QuarterOne:
            case OfficeConditionalIconKind.QuarterTwo:
            case OfficeConditionalIconKind.QuarterThree:
            case OfficeConditionalIconKind.QuarterFull:
                DrawRasterQuarterPie(canvas, x, y, size, GetQuarterFillCount(kind), fill, stroke, strokeWidth);
                break;
            case OfficeConditionalIconKind.GreenFlag:
            case OfficeConditionalIconKind.YellowFlag:
            case OfficeConditionalIconKind.RedFlag:
                DrawRasterFlag(canvas, x, y, size, fill, stroke, strokeWidth);
                break;
            default:
                IReadOnlyList<OfficePoint> points = CreateArrowPoints(x, y, size, kind);
                canvas.FillPolygon(OffsetPoints(points, size * 0.055D, size * 0.065D), IconShadowColor);
                canvas.FillPolygon(points, fill);
                canvas.DrawPolygon(points, stroke, strokeWidth);
                DrawRasterArrowHighlight(canvas, x, y, size, kind);
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
                AppendSvgLineShadow(builder, x + size * 0.22D, y + size * 0.54D, x + size * 0.42D, y + size * 0.74D, size);
                AppendSvgLineShadow(builder, x + size * 0.42D, y + size * 0.74D, x + size * 0.80D, y + size * 0.28D, size);
                AppendSvgLine(builder, x + size * 0.22D, y + size * 0.54D, x + size * 0.42D, y + size * 0.74D, fill, size * 0.14D);
                AppendSvgLine(builder, x + size * 0.42D, y + size * 0.74D, x + size * 0.80D, y + size * 0.28D, fill, size * 0.14D);
                break;
            case OfficeConditionalIconKind.YellowExclamation:
                AppendSvgCircle(builder, x, y, size, fill, stroke, Math.Max(1D, size / 14D));
                AppendSvgLine(builder, x + size * 0.5D, y + size * 0.25D, x + size * 0.5D, y + size * 0.60D, OfficeColor.White, size * 0.12D);
                builder.AppendCircleElement(x + size * 0.5D, y + size * 0.78D, size * 0.06D, OfficeColor.White);
                break;
            case OfficeConditionalIconKind.RedCross:
                AppendSvgLineShadow(builder, x + size * 0.25D, y + size * 0.25D, x + size * 0.75D, y + size * 0.75D, size);
                AppendSvgLineShadow(builder, x + size * 0.75D, y + size * 0.25D, x + size * 0.25D, y + size * 0.75D, size);
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
            case OfficeConditionalIconKind.RatingOne:
            case OfficeConditionalIconKind.RatingTwo:
            case OfficeConditionalIconKind.RatingThree:
            case OfficeConditionalIconKind.RatingFour:
            case OfficeConditionalIconKind.RatingFive:
                AppendSvgRatingStar(builder, x, y, size, fill, stroke, strokeWidth);
                break;
            case OfficeConditionalIconKind.QuarterEmpty:
            case OfficeConditionalIconKind.QuarterOne:
            case OfficeConditionalIconKind.QuarterTwo:
            case OfficeConditionalIconKind.QuarterThree:
            case OfficeConditionalIconKind.QuarterFull:
                AppendSvgQuarterPie(builder, x, y, size, GetQuarterFillCount(kind), fill, stroke, strokeWidth);
                break;
            case OfficeConditionalIconKind.GreenFlag:
            case OfficeConditionalIconKind.YellowFlag:
            case OfficeConditionalIconKind.RedFlag:
                AppendSvgFlag(builder, x, y, size, fill, stroke, strokeWidth);
                break;
            default:
                IReadOnlyList<OfficePoint> points = CreateArrowPoints(x, y, size, kind);
                var shadowAttributes = new StringBuilder()
                    .AppendPaintAttribute("fill", IconShadowColor)
                    .ToString();
                builder.AppendPathElement(OfficeSvgFormatting.FormatMoveLinePathData(OffsetPoints(points, size * 0.055D, size * 0.065D), closePath: true), shadowAttributes);
                var attributes = new StringBuilder()
                    .AppendPaintAttribute("fill", fill)
                    .AppendPaintAttribute("stroke", stroke)
                    .AppendNumberAttribute("stroke-width", strokeWidth)
                    .ToString();
                builder.AppendPathElement(OfficeSvgFormatting.FormatMoveLinePathData(points, closePath: true), attributes);
                AppendSvgArrowHighlight(builder, x, y, size, kind);
                break;
        }

        return builder;
    }

    private static void DrawRasterCircle(OfficeRasterCanvas canvas, double x, double y, double size, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        double shadowOffset = Math.Max(0.75D, size * 0.045D);
        canvas.FillEllipse(x + shadowOffset, y + shadowOffset, size, size, IconShadowColor);
        canvas.FillEllipse(x, y, size, size, fill);
        canvas.FillEllipse(x + size * 0.18D, y + size * 0.14D, size * 0.44D, size * 0.26D, IconHighlightColor);
        canvas.DrawEllipse(x, y, size, size, stroke, strokeWidth);
    }

    private static void DrawRasterFlag(OfficeRasterCanvas canvas, double x, double y, double size, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        double poleX = x + size * 0.24D;
        double poleTop = y + size * 0.12D;
        double poleBottom = y + size * 0.90D;
        double poleWidth = Math.Max(1D, size * 0.105D);
        IReadOnlyList<OfficePoint> points = CreateFlagPoints(x, y, size);
        canvas.DrawLine(poleX + size * 0.055D, poleTop + size * 0.065D, poleX + size * 0.055D, poleBottom + size * 0.065D, IconShadowColor, poleWidth);
        canvas.FillPolygon(OffsetPoints(points, size * 0.055D, size * 0.065D), IconShadowColor);
        canvas.DrawLine(poleX, poleTop, poleX, poleBottom, stroke, poleWidth);
        canvas.FillPolygon(points, fill);
        canvas.DrawPolygon(points, stroke, strokeWidth);
        canvas.FillPolygon(CreateFlagHighlightPoints(x, y, size), IconHighlightColor);
    }

    private static void DrawRasterRatingStar(OfficeRasterCanvas canvas, double x, double y, double size, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        IReadOnlyList<OfficePoint> points = CreateRatingStarPoints(x, y, size);
        canvas.FillPolygon(OffsetPoints(points, size * 0.055D, size * 0.065D), IconShadowColor);
        canvas.FillPolygon(points, fill);
        canvas.FillEllipse(x + size * 0.28D, y + size * 0.22D, size * 0.18D, size * 0.12D, IconHighlightColor);
        canvas.DrawPolygon(points, stroke, Math.Max(1D, strokeWidth));
    }

    private static void DrawRasterQuarterPie(OfficeRasterCanvas canvas, double x, double y, double size, int quarters, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        OfficeColor emptyFill = OfficeColor.FromRgb(241, 245, 249);
        double shadowOffset = Math.Max(0.75D, size * 0.045D);
        canvas.FillEllipse(x + shadowOffset, y + shadowOffset, size, size, IconShadowColor);
        canvas.FillEllipse(x, y, size, size, emptyFill);
        if (quarters >= 4) {
            canvas.FillEllipse(x, y, size, size, fill);
        } else if (quarters > 0) {
            canvas.FillPolygon(CreateQuarterPiePoints(x, y, size, quarters), fill);
        }

        canvas.FillEllipse(x + size * 0.18D, y + size * 0.14D, size * 0.44D, size * 0.26D, IconHighlightColor);
        canvas.DrawEllipse(x, y, size, size, stroke, strokeWidth);
    }

    private static void AppendSvgCircle(StringBuilder builder, double x, double y, double size, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        double shadowOffset = Math.Max(0.75D, size * 0.045D);
        builder.AppendCircleElement(x + size / 2D + shadowOffset, y + size / 2D + shadowOffset, size / 2D, IconShadowColor);
        var attributes = new StringBuilder()
            .AppendPaintAttribute("fill", fill)
            .AppendPaintAttribute("stroke", stroke)
            .AppendNumberAttribute("stroke-width", strokeWidth)
            .ToString();
        builder.AppendCircleElement(x + size / 2D, y + size / 2D, size / 2D, attributes);
        builder.AppendEllipseElement(x + size * 0.4D, y + size * 0.27D, size * 0.22D, size * 0.13D, IconHighlightColor);
    }

    private static void AppendSvgFlag(StringBuilder builder, double x, double y, double size, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        double poleX = x + size * 0.24D;
        double poleTop = y + size * 0.12D;
        double poleBottom = y + size * 0.90D;
        double poleWidth = Math.Max(1D, size * 0.105D);
        var shadowAttributes = new StringBuilder()
            .AppendPaintAttribute("fill", IconShadowColor)
            .ToString();
        builder.AppendLineElement(poleX + size * 0.055D, poleTop + size * 0.065D, poleX + size * 0.055D, poleBottom + size * 0.065D, IconShadowColor, poleWidth, OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap.Round);
        builder.AppendPolygonElement(OffsetPoints(CreateFlagPoints(x, y, size), size * 0.055D, size * 0.065D), shadowAttributes);
        builder.AppendLineElement(poleX, poleTop, poleX, poleBottom, stroke, poleWidth, OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap.Round);
        var flagAttributes = new StringBuilder()
            .AppendPaintAttribute("fill", fill)
            .AppendPaintAttribute("stroke", stroke)
            .AppendNumberAttribute("stroke-width", strokeWidth)
            .AppendAttribute("stroke-linejoin", "round")
            .ToString();
        builder.AppendPolygonElement(CreateFlagPoints(x, y, size), flagAttributes);
        builder.AppendPolygonElement(
            CreateFlagHighlightPoints(x, y, size),
            new StringBuilder().AppendPaintAttribute("fill", IconHighlightColor).ToString());
    }

    private static void AppendSvgRatingStar(StringBuilder builder, double x, double y, double size, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        IReadOnlyList<OfficePoint> points = CreateRatingStarPoints(x, y, size);
        var shadowAttributes = new StringBuilder()
            .AppendPaintAttribute("fill", IconShadowColor)
            .ToString();
        builder.AppendPolygonElement(OffsetPoints(points, size * 0.055D, size * 0.065D), shadowAttributes);
        var attributes = new StringBuilder()
            .AppendPaintAttribute("fill", fill)
            .AppendPaintAttribute("stroke", stroke)
            .AppendNumberAttribute("stroke-width", Math.Max(1D, strokeWidth))
            .AppendAttribute("stroke-linejoin", "round")
            .ToString();
        builder.AppendPolygonElement(points, attributes);
        builder.AppendEllipseElement(x + size * 0.37D, y + size * 0.28D, size * 0.09D, size * 0.06D, IconHighlightColor);
    }

    private static void AppendSvgQuarterPie(StringBuilder builder, double x, double y, double size, int quarters, OfficeColor fill, OfficeColor stroke, double strokeWidth) {
        OfficeColor emptyFill = OfficeColor.FromRgb(241, 245, 249);
        double shadowOffset = Math.Max(0.75D, size * 0.045D);
        builder.AppendCircleElement(x + size / 2D + shadowOffset, y + size / 2D + shadowOffset, size / 2D, IconShadowColor);
        builder.AppendCircleElement(
            x + size / 2D,
            y + size / 2D,
            size / 2D,
            new StringBuilder()
                .AppendPaintAttribute("fill", emptyFill)
                .AppendPaintAttribute("stroke", stroke)
                .AppendNumberAttribute("stroke-width", strokeWidth)
                .ToString());

        if (quarters >= 4) {
            builder.AppendCircleElement(x + size / 2D, y + size / 2D, size / 2D - (strokeWidth * 0.5D), fill);
        } else if (quarters > 0) {
            var attributes = new StringBuilder()
                .AppendPaintAttribute("fill", fill)
                .ToString();
            double inset = strokeWidth * 0.5D;
            builder.AppendPolygonElement(CreateQuarterPiePoints(x + inset, y + inset, size - strokeWidth, quarters), attributes);
        }

        builder.AppendEllipseElement(x + size * 0.4D, y + size * 0.27D, size * 0.22D, size * 0.13D, IconHighlightColor);
    }

    private static void AppendSvgLine(StringBuilder builder, double x1, double y1, double x2, double y2, OfficeColor color, double width) {
        builder.AppendLineElement(x1, y1, x2, y2, color, Math.Max(1D, width), OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap.Round);
    }

    private static void DrawRasterLineShadow(OfficeRasterCanvas canvas, double x1, double y1, double x2, double y2, double size) {
        double offset = Math.Max(0.75D, size * 0.045D);
        canvas.DrawLine(x1 + offset, y1 + offset, x2 + offset, y2 + offset, IconShadowColor, Math.Max(2D, size * 0.14D));
    }

    private static void AppendSvgLineShadow(StringBuilder builder, double x1, double y1, double x2, double y2, double size) {
        double offset = Math.Max(0.75D, size * 0.045D);
        AppendSvgLine(builder, x1 + offset, y1 + offset, x2 + offset, y2 + offset, IconShadowColor, size * 0.14D);
    }

    private static void DrawRasterArrowHighlight(OfficeRasterCanvas canvas, double x, double y, double size, OfficeConditionalIconKind kind) {
        if (kind == OfficeConditionalIconKind.YellowSideArrow) {
            canvas.DrawLine(x + size * 0.20D, y + size * 0.43D, x + size * 0.55D, y + size * 0.43D, IconHighlightColor, Math.Max(1D, size * 0.08D));
            return;
        }

        if (kind == OfficeConditionalIconKind.RedDownArrow || kind == OfficeConditionalIconKind.YellowDownArrow) {
            canvas.DrawLine(x + size * 0.42D, y + size * 0.18D, x + size * 0.42D, y + size * 0.53D, IconHighlightColor, Math.Max(1D, size * 0.08D));
            return;
        }

        canvas.DrawLine(x + size * 0.42D, y + size * 0.88D, x + size * 0.42D, y + size * 0.48D, IconHighlightColor, Math.Max(1D, size * 0.08D));
    }

    private static void AppendSvgArrowHighlight(StringBuilder builder, double x, double y, double size, OfficeConditionalIconKind kind) {
        if (kind == OfficeConditionalIconKind.YellowSideArrow) {
            AppendSvgLine(builder, x + size * 0.20D, y + size * 0.43D, x + size * 0.55D, y + size * 0.43D, IconHighlightColor, size * 0.08D);
            return;
        }

        if (kind == OfficeConditionalIconKind.RedDownArrow || kind == OfficeConditionalIconKind.YellowDownArrow) {
            AppendSvgLine(builder, x + size * 0.42D, y + size * 0.18D, x + size * 0.42D, y + size * 0.53D, IconHighlightColor, size * 0.08D);
            return;
        }

        AppendSvgLine(builder, x + size * 0.42D, y + size * 0.88D, x + size * 0.42D, y + size * 0.48D, IconHighlightColor, size * 0.08D);
    }

    private static IReadOnlyList<OfficePoint> OffsetPoints(IReadOnlyList<OfficePoint> points, double offsetX, double offsetY) {
        var shifted = new OfficePoint[points.Count];
        for (int i = 0; i < points.Count; i++) {
            shifted[i] = new OfficePoint(points[i].X + offsetX, points[i].Y + offsetY);
        }

        return shifted;
    }

    private static IReadOnlyList<OfficePoint> CreateQuarterPiePoints(double x, double y, double size, int quarters) {
        double inset = Math.Max(0D, size * 0.04D);
        double radius = Math.Max(1D, (size / 2D) - inset);
        double centerX = x + size / 2D;
        double centerY = y + size / 2D;
        double start = -Math.PI / 2D;
        double sweep = Math.PI * 0.5D * Math.Max(0, Math.Min(4, quarters));
        var points = new List<OfficePoint> {
            new OfficePoint(centerX, centerY),
            new OfficePoint(centerX + Math.Cos(start) * radius, centerY + Math.Sin(start) * radius)
        };
        points.AddRange(OfficeGeometry.CreateEllipticalArcPoints(centerX, centerY, radius, radius, start, sweep, Math.Max(3, quarters * 4)));
        return points;
    }

    private static IReadOnlyList<OfficePoint> CreateFlagPoints(double x, double y, double size) =>
        new[] {
            new OfficePoint(x + size * 0.25D, y + size * 0.14D),
            new OfficePoint(x + size * 0.78D, y + size * 0.18D),
            new OfficePoint(x + size * 0.68D, y + size * 0.38D),
            new OfficePoint(x + size * 0.82D, y + size * 0.58D),
            new OfficePoint(x + size * 0.25D, y + size * 0.52D)
        };

    private static IReadOnlyList<OfficePoint> CreateFlagHighlightPoints(double x, double y, double size) =>
        new[] {
            new OfficePoint(x + size * 0.32D, y + size * 0.20D),
            new OfficePoint(x + size * 0.62D, y + size * 0.22D),
            new OfficePoint(x + size * 0.56D, y + size * 0.31D),
            new OfficePoint(x + size * 0.32D, y + size * 0.30D)
        };

    private static IReadOnlyList<OfficePoint> CreateRatingStarPoints(double x, double y, double size) {
        const int pointCount = 10;
        double centerX = x + size * 0.5D;
        double centerY = y + size * 0.52D;
        double outerRadius = size * 0.44D;
        double innerRadius = size * 0.205D;
        double startAngle = -Math.PI / 2D;
        var points = new OfficePoint[pointCount];
        for (int i = 0; i < pointCount; i++) {
            double radius = i % 2 == 0 ? outerRadius : innerRadius;
            double angle = startAngle + (Math.PI * 2D * i / pointCount);
            points[i] = new OfficePoint(centerX + Math.Cos(angle) * radius, centerY + Math.Sin(angle) * radius);
        }

        return points;
    }

    private static int GetQuarterFillCount(OfficeConditionalIconKind kind) =>
        kind switch {
            OfficeConditionalIconKind.QuarterOne => 1,
            OfficeConditionalIconKind.QuarterTwo => 2,
            OfficeConditionalIconKind.QuarterThree => 3,
            OfficeConditionalIconKind.QuarterFull => 4,
            _ => 0
        };

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
            OfficeConditionalIconKind.GreenUpArrow or OfficeConditionalIconKind.GreenCheck or OfficeConditionalIconKind.GreenCircle or OfficeConditionalIconKind.RatingFive or OfficeConditionalIconKind.QuarterFull or OfficeConditionalIconKind.GreenFlag => OfficeColor.FromRgb(22, 163, 74),
            OfficeConditionalIconKind.LightGreenCircle or OfficeConditionalIconKind.RatingFour or OfficeConditionalIconKind.QuarterThree => OfficeColor.FromRgb(132, 204, 22),
            OfficeConditionalIconKind.YellowUpArrow or OfficeConditionalIconKind.YellowSideArrow or OfficeConditionalIconKind.YellowDownArrow or OfficeConditionalIconKind.YellowExclamation or OfficeConditionalIconKind.YellowCircle or OfficeConditionalIconKind.RatingThree or OfficeConditionalIconKind.QuarterTwo or OfficeConditionalIconKind.YellowFlag => OfficeColor.FromRgb(245, 158, 11),
            OfficeConditionalIconKind.OrangeCircle or OfficeConditionalIconKind.RatingTwo or OfficeConditionalIconKind.QuarterOne => OfficeColor.FromRgb(249, 115, 22),
            OfficeConditionalIconKind.QuarterEmpty => OfficeColor.FromRgb(148, 163, 184),
            _ => OfficeColor.FromRgb(220, 38, 38)
        };

    private static OfficeColor GetStrokeColor(OfficeConditionalIconKind kind) =>
        kind switch {
            OfficeConditionalIconKind.GreenUpArrow or OfficeConditionalIconKind.GreenCheck or OfficeConditionalIconKind.GreenCircle or OfficeConditionalIconKind.RatingFive or OfficeConditionalIconKind.QuarterFull or OfficeConditionalIconKind.GreenFlag => OfficeColor.FromRgb(21, 128, 61),
            OfficeConditionalIconKind.LightGreenCircle or OfficeConditionalIconKind.RatingFour or OfficeConditionalIconKind.QuarterThree => OfficeColor.FromRgb(77, 124, 15),
            OfficeConditionalIconKind.YellowUpArrow or OfficeConditionalIconKind.YellowSideArrow or OfficeConditionalIconKind.YellowDownArrow or OfficeConditionalIconKind.YellowExclamation or OfficeConditionalIconKind.YellowCircle or OfficeConditionalIconKind.RatingThree or OfficeConditionalIconKind.QuarterTwo or OfficeConditionalIconKind.YellowFlag => OfficeColor.FromRgb(180, 83, 9),
            OfficeConditionalIconKind.OrangeCircle or OfficeConditionalIconKind.RatingTwo or OfficeConditionalIconKind.QuarterOne => OfficeColor.FromRgb(194, 65, 12),
            OfficeConditionalIconKind.QuarterEmpty => OfficeColor.FromRgb(100, 116, 139),
            _ => OfficeColor.FromRgb(185, 28, 28)
        };
}
