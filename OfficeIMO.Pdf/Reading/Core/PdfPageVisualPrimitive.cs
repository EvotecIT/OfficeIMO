using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal enum PdfPageVisualPrimitiveKind {
    Rectangle,
    Line,
    Path
}

internal readonly struct PdfPageVisualPrimitive {
    private PdfPageVisualPrimitive(
        PdfPageVisualPrimitiveKind kind,
        double x,
        double y,
        double width,
        double height,
        double x1,
        double y1,
        double x2,
        double y2,
        IReadOnlyList<OfficePathCommand> pathCommands,
        OfficeColor? fillColor,
        OfficeLinearGradient? fillGradient,
        OfficeRadialGradient? fillRadialGradient,
        OfficeColor? strokeColor,
        OfficeLinearGradient? strokeGradient,
        OfficeRadialGradient? strokeRadialGradient,
        double strokeWidth,
        OfficeStrokeDashStyle strokeDashStyle,
        OfficeStrokeLineCap? strokeLineCap,
        OfficeStrokeLineJoin? strokeLineJoin,
        double? fillOpacity,
        double? strokeOpacity,
        OfficeFillRule fillRule,
        PdfPageClipPath? clipPath,
        double paintOrder = 0D,
        PdfPageTilingPatternPaint? fillTilingPattern = null,
        PdfPageTilingPatternPaint? strokeTilingPattern = null) {
        Kind = kind;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        PathCommands = pathCommands;
        FillColor = fillColor;
        FillGradient = fillGradient;
        FillRadialGradient = fillRadialGradient;
        StrokeColor = strokeColor;
        StrokeGradient = strokeGradient;
        StrokeRadialGradient = strokeRadialGradient;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        StrokeLineCap = strokeLineCap;
        StrokeLineJoin = strokeLineJoin;
        FillOpacity = fillOpacity;
        StrokeOpacity = strokeOpacity;
        FillRule = fillRule;
        ClipPath = clipPath;
        PaintOrder = paintOrder;
        FillTilingPattern = fillTilingPattern;
        StrokeTilingPattern = strokeTilingPattern;
    }

    public static PdfPageVisualPrimitive Rectangle(double x, double y, double width, double height, OfficeColor? fillColor, OfficeColor? strokeColor, double strokeWidth, OfficeStrokeDashStyle strokeDashStyle, OfficeStrokeLineCap? strokeLineCap, OfficeStrokeLineJoin? strokeLineJoin, double? fillOpacity, double? strokeOpacity, PdfPageClipPath? clipPath, double paintOrder = 0D) =>
        new PdfPageVisualPrimitive(PdfPageVisualPrimitiveKind.Rectangle, x, y, width, height, x, y, x + width, y + height, Array.Empty<OfficePathCommand>(), fillColor, null, null, strokeColor, null, null, strokeWidth, strokeDashStyle, strokeLineCap, strokeLineJoin, fillOpacity, strokeOpacity, OfficeFillRule.EvenOdd, clipPath, paintOrder);

    public static PdfPageVisualPrimitive Rectangle(double x, double y, double width, double height, OfficeColor? fillColor, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? strokeColor, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth, OfficeStrokeDashStyle strokeDashStyle, OfficeStrokeLineCap? strokeLineCap, OfficeStrokeLineJoin? strokeLineJoin, double? fillOpacity, double? strokeOpacity, PdfPageClipPath? clipPath, double paintOrder = 0D, PdfPageTilingPatternPaint? fillTilingPattern = null, PdfPageTilingPatternPaint? strokeTilingPattern = null) =>
        new PdfPageVisualPrimitive(PdfPageVisualPrimitiveKind.Rectangle, x, y, width, height, x, y, x + width, y + height, Array.Empty<OfficePathCommand>(), fillColor, fillGradient, fillRadialGradient, strokeColor, strokeGradient, strokeRadialGradient, strokeWidth, strokeDashStyle, strokeLineCap, strokeLineJoin, fillOpacity, strokeOpacity, OfficeFillRule.EvenOdd, clipPath, paintOrder, fillTilingPattern, strokeTilingPattern);

    public static PdfPageVisualPrimitive ShadedRectangle(double x, double y, double width, double height, OfficeLinearGradient fillGradient, double? fillOpacity, PdfPageClipPath? clipPath, double paintOrder = 0D) =>
        new PdfPageVisualPrimitive(PdfPageVisualPrimitiveKind.Rectangle, x, y, width, height, x, y, x + width, y + height, Array.Empty<OfficePathCommand>(), null, fillGradient, null, null, null, null, 0D, OfficeStrokeDashStyle.Solid, null, null, fillOpacity, null, OfficeFillRule.EvenOdd, clipPath, paintOrder);

    public static PdfPageVisualPrimitive ShadedRectangle(double x, double y, double width, double height, OfficeRadialGradient fillRadialGradient, double? fillOpacity, PdfPageClipPath? clipPath, double paintOrder = 0D) =>
        new PdfPageVisualPrimitive(PdfPageVisualPrimitiveKind.Rectangle, x, y, width, height, x, y, x + width, y + height, Array.Empty<OfficePathCommand>(), null, null, fillRadialGradient, null, null, null, 0D, OfficeStrokeDashStyle.Solid, null, null, fillOpacity, null, OfficeFillRule.EvenOdd, clipPath, paintOrder);

    public static PdfPageVisualPrimitive Line(double x1, double y1, double x2, double y2, OfficeColor? strokeColor, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth, OfficeStrokeDashStyle strokeDashStyle, OfficeStrokeLineCap? strokeLineCap, OfficeStrokeLineJoin? strokeLineJoin, double? strokeOpacity, PdfPageClipPath? clipPath, double paintOrder = 0D, PdfPageTilingPatternPaint? strokeTilingPattern = null) {
        double left = Math.Min(x1, x2);
        double top = Math.Min(y1, y2);
        return new PdfPageVisualPrimitive(PdfPageVisualPrimitiveKind.Line, left, top, Math.Abs(x2 - x1), Math.Abs(y2 - y1), x1, y1, x2, y2, Array.Empty<OfficePathCommand>(), null, null, null, strokeColor, strokeGradient, strokeRadialGradient, strokeWidth, strokeDashStyle, strokeLineCap, strokeLineJoin, null, strokeOpacity, OfficeFillRule.EvenOdd, clipPath, paintOrder, null, strokeTilingPattern);
    }

    public static bool TryCreatePath(IReadOnlyList<OfficePathCommand> pathCommands, OfficeColor? fillColor, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor? strokeColor, OfficeLinearGradient? strokeGradient, OfficeRadialGradient? strokeRadialGradient, double strokeWidth, OfficeStrokeDashStyle strokeDashStyle, OfficeStrokeLineCap? strokeLineCap, OfficeStrokeLineJoin? strokeLineJoin, double? fillOpacity, double? strokeOpacity, OfficeFillRule fillRule, PdfPageClipPath? clipPath, double paintOrder, PdfPageTilingPatternPaint? fillTilingPattern, PdfPageTilingPatternPaint? strokeTilingPattern, bool retainPathCommands, out PdfPageVisualPrimitive primitive) {
        primitive = default;
        if (pathCommands.Count == 0 || pathCommands[0].Kind != OfficePathCommandKind.MoveTo) {
            return false;
        }

        bool hasPoint = false;
        bool hasDraw = false;
        double left = 0D;
        double top = 0D;
        double right = 0D;
        double bottom = 0D;
        for (int i = 0; i < pathCommands.Count; i++) {
            OfficePathCommand command = pathCommands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    break;
                case OfficePathCommandKind.LineTo:
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    Include(command.ControlPoint1, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    Include(command.ControlPoint1, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    Include(command.ControlPoint2, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    Include(command.Point, ref hasPoint, ref left, ref top, ref right, ref bottom);
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.Close:
                    break;
            }
        }

        double width = right - left;
        double height = bottom - top;
        if (!hasDraw || width <= 0D || height <= 0D) {
            return false;
        }

        IReadOnlyList<OfficePathCommand> retainedCommands = retainPathCommands
            ? new List<OfficePathCommand>(pathCommands)
            : pathCommands;
        primitive = new PdfPageVisualPrimitive(PdfPageVisualPrimitiveKind.Path, left, top, width, height, left, top, right, bottom, retainedCommands, fillColor, fillGradient, fillRadialGradient, strokeColor, strokeGradient, strokeRadialGradient, strokeWidth, strokeDashStyle, strokeLineCap, strokeLineJoin, fillOpacity, strokeOpacity, fillRule, clipPath, paintOrder, fillTilingPattern, strokeTilingPattern);
        return true;
    }

    public PdfPageVisualPrimitiveKind Kind { get; }

    public double X { get; }

    public double Y { get; }

    public double Width { get; }

    public double Height { get; }

    public double X1 { get; }

    public double Y1 { get; }

    public double X2 { get; }

    public double Y2 { get; }

    public IReadOnlyList<OfficePathCommand> PathCommands { get; }

    public OfficeColor? FillColor { get; }

    public OfficeLinearGradient? FillGradient { get; }

    public OfficeRadialGradient? FillRadialGradient { get; }

    public OfficeColor? StrokeColor { get; }

    public OfficeLinearGradient? StrokeGradient { get; }

    public OfficeRadialGradient? StrokeRadialGradient { get; }

    public double StrokeWidth { get; }

    public OfficeStrokeDashStyle StrokeDashStyle { get; }

    public OfficeStrokeLineCap? StrokeLineCap { get; }

    public OfficeStrokeLineJoin? StrokeLineJoin { get; }

    public double? FillOpacity { get; }

    public double? StrokeOpacity { get; }

    public OfficeFillRule FillRule { get; }

    public PdfPageClipPath? ClipPath { get; }

    public double PaintOrder { get; }

    public PdfPageTilingPatternPaint? FillTilingPattern { get; }

    public PdfPageTilingPatternPaint? StrokeTilingPattern { get; }

    private static void Include(OfficePoint point, ref bool hasPoint, ref double left, ref double top, ref double right, ref double bottom) {
        if (!hasPoint) {
            left = right = point.X;
            top = bottom = point.Y;
            hasPoint = true;
            return;
        }

        if (point.X < left) {
            left = point.X;
        }

        if (point.Y < top) {
            top = point.Y;
        }

        if (point.X > right) {
            right = point.X;
        }

        if (point.Y > bottom) {
            bottom = point.Y;
        }
    }
}
