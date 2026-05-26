namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void DrawRowFill(StringBuilder sb, PdfColor color, double x, double y, double w, double h) {
        new ContentStreamBuilder(sb)
            .SaveState()
            .FillColor(color)
            .Rectangle(x, y, w, h)
            .FillPath()
            .RestoreState();
    }

    private static void DrawRowRect(StringBuilder sb, PdfColor color, double widthStroke, double x, double y, double w, double h) {
        new ContentStreamBuilder(sb)
            .SaveState()
            .StrokeColor(color)
            .LineWidth(widthStroke)
            .Rectangle(x, y, w, h)
            .StrokePath()
            .RestoreState();
    }

    private static void DrawRectangle(StringBuilder sb, PdfColor? fillColor, PdfColor? strokeColor, double strokeWidth, OfficeIMO.Drawing.OfficeStrokeDashStyle strokeDashStyle, OfficeIMO.Drawing.OfficeStrokeLineCap? strokeLineCap, OfficeIMO.Drawing.OfficeStrokeLineJoin? strokeLineJoin, double x, double y, double w, double h) {
        if (!fillColor.HasValue && (!strokeColor.HasValue || strokeWidth <= 0)) {
            return;
        }

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (fillColor.HasValue) {
            content.FillColor(fillColor.Value);
        }

        bool stroke = strokeColor.HasValue && strokeWidth > 0;
        if (stroke) {
            content
                .StrokeColor(strokeColor!.Value)
                .LineWidth(strokeWidth);
            ApplyStrokeStyle(content, strokeDashStyle, strokeWidth, strokeLineCap, strokeLineJoin);
        }

        content.Rectangle(x, y, w, h);
        if (fillColor.HasValue && stroke) {
            content.FillStrokePath();
        } else if (fillColor.HasValue) {
            content.FillPath();
        } else {
            content.StrokePath();
        }

        content.RestoreState();
    }

    private static void DrawLine(StringBuilder sb, PdfColor? strokeColor, double strokeWidth, OfficeIMO.Drawing.OfficeStrokeDashStyle strokeDashStyle, OfficeIMO.Drawing.OfficeStrokeLineCap? strokeLineCap, OfficeIMO.Drawing.OfficeStrokeLineJoin? strokeLineJoin, System.Collections.Generic.IReadOnlyList<OfficeIMO.Drawing.OfficePoint> points, double x, double y, double h) {
        if (points.Count != 2 || !strokeColor.HasValue || strokeWidth <= 0) {
            return;
        }

        var content = new ContentStreamBuilder(sb)
            .SaveState()
            .StrokeColor(strokeColor.Value)
            .LineWidth(strokeWidth);
        ApplyStrokeStyle(content, strokeDashStyle, strokeWidth, strokeLineCap, strokeLineJoin);
        content
            .MoveTo(x + points[0].X, y + h - points[0].Y)
            .LineTo(x + points[1].X, y + h - points[1].Y)
            .StrokePath()
            .RestoreState();
    }

    private static void DrawRoundedRectangle(StringBuilder sb, PdfColor? fillColor, PdfColor? strokeColor, double strokeWidth, OfficeIMO.Drawing.OfficeStrokeDashStyle strokeDashStyle, OfficeIMO.Drawing.OfficeStrokeLineCap? strokeLineCap, OfficeIMO.Drawing.OfficeStrokeLineJoin? strokeLineJoin, double x, double y, double w, double h, double cornerRadius) {
        if (cornerRadius <= 0) {
            DrawRectangle(sb, fillColor, strokeColor, strokeWidth, strokeDashStyle, strokeLineCap, strokeLineJoin, x, y, w, h);
            return;
        }

        if (!fillColor.HasValue && (!strokeColor.HasValue || strokeWidth <= 0)) {
            return;
        }

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (fillColor.HasValue) {
            content.FillColor(fillColor.Value);
        }

        bool stroke = strokeColor.HasValue && strokeWidth > 0;
        if (stroke) {
            content
                .StrokeColor(strokeColor!.Value)
                .LineWidth(strokeWidth);
            ApplyStrokeStyle(content, strokeDashStyle, strokeWidth, strokeLineCap, strokeLineJoin);
        }

        AppendRoundedPath(content, x, y, w, h, cornerRadius);
        PaintPath(content, fillColor.HasValue, stroke, closePath: true);
        content.RestoreState();
    }

    private static void DrawEllipse(StringBuilder sb, PdfColor? fillColor, PdfColor? strokeColor, double strokeWidth, OfficeIMO.Drawing.OfficeStrokeDashStyle strokeDashStyle, OfficeIMO.Drawing.OfficeStrokeLineCap? strokeLineCap, OfficeIMO.Drawing.OfficeStrokeLineJoin? strokeLineJoin, double x, double y, double w, double h) {
        if (!fillColor.HasValue && (!strokeColor.HasValue || strokeWidth <= 0)) {
            return;
        }

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (fillColor.HasValue) {
            content.FillColor(fillColor.Value);
        }

        bool stroke = strokeColor.HasValue && strokeWidth > 0;
        if (stroke) {
            content
                .StrokeColor(strokeColor!.Value)
                .LineWidth(strokeWidth);
            ApplyStrokeStyle(content, strokeDashStyle, strokeWidth, strokeLineCap, strokeLineJoin);
        }

        AppendEllipsePath(content, x, y, w, h);
        PaintPath(content, fillColor.HasValue, stroke, closePath: false);
        content.RestoreState();
    }

    private static void DrawPolygon(StringBuilder sb, PdfColor? fillColor, PdfColor? strokeColor, double strokeWidth, OfficeIMO.Drawing.OfficeStrokeDashStyle strokeDashStyle, OfficeIMO.Drawing.OfficeStrokeLineCap? strokeLineCap, OfficeIMO.Drawing.OfficeStrokeLineJoin? strokeLineJoin, System.Collections.Generic.IReadOnlyList<OfficeIMO.Drawing.OfficePoint> points, double x, double y, double h) {
        if (points.Count < 3 || (!fillColor.HasValue && (!strokeColor.HasValue || strokeWidth <= 0))) {
            return;
        }

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (fillColor.HasValue) {
            content.FillColor(fillColor.Value);
        }

        bool stroke = strokeColor.HasValue && strokeWidth > 0;
        if (stroke) {
            content
                .StrokeColor(strokeColor!.Value)
                .LineWidth(strokeWidth);
            ApplyStrokeStyle(content, strokeDashStyle, strokeWidth, strokeLineCap, strokeLineJoin);
        }

        content.MoveTo(x + points[0].X, y + h - points[0].Y);
        for (int i = 1; i < points.Count; i++) {
            content.LineTo(x + points[i].X, y + h - points[i].Y);
        }

        PaintPath(content, fillColor.HasValue, stroke, closePath: true);
        content.RestoreState();
    }

    private static void DrawPath(StringBuilder sb, PdfColor? fillColor, PdfColor? strokeColor, double strokeWidth, OfficeIMO.Drawing.OfficeStrokeDashStyle strokeDashStyle, OfficeIMO.Drawing.OfficeStrokeLineCap? strokeLineCap, OfficeIMO.Drawing.OfficeStrokeLineJoin? strokeLineJoin, System.Collections.Generic.IReadOnlyList<OfficeIMO.Drawing.OfficePathCommand> commands, double x, double y, double h) {
        if (commands.Count == 0 || (!fillColor.HasValue && (!strokeColor.HasValue || strokeWidth <= 0))) {
            return;
        }

        var content = new ContentStreamBuilder(sb)
            .SaveState();
        if (fillColor.HasValue) {
            content.FillColor(fillColor.Value);
        }

        bool stroke = strokeColor.HasValue && strokeWidth > 0;
        if (stroke) {
            content
                .StrokeColor(strokeColor!.Value)
                .LineWidth(strokeWidth);
            ApplyStrokeStyle(content, strokeDashStyle, strokeWidth, strokeLineCap, strokeLineJoin);
        }

        AppendPathCommands(content, commands, x, y, h);
        PaintPath(content, fillColor.HasValue, stroke, closePath: false);
        content.RestoreState();
    }

    private static void AppendRoundedPath(ContentStreamBuilder content, double x, double y, double w, double h, double cornerRadius) {
        double r = Math.Min(cornerRadius, Math.Min(w, h) / 2D);
        double c = r * 0.5522847498307936;
        double x2 = x + w;
        double y2 = y + h;

        content
            .MoveTo(x + r, y)
            .LineTo(x2 - r, y)
            .CubicTo(x2 - r + c, y, x2, y + r - c, x2, y + r)
            .LineTo(x2, y2 - r)
            .CubicTo(x2, y2 - r + c, x2 - r + c, y2, x2 - r, y2)
            .LineTo(x + r, y2)
            .CubicTo(x + r - c, y2, x, y2 - r + c, x, y2 - r)
            .LineTo(x, y + r)
            .CubicTo(x, y + r - c, x + r - c, y, x + r, y);
    }

    private static void AppendEllipsePath(ContentStreamBuilder content, double x, double y, double w, double h) {
        double rx = w / 2;
        double ry = h / 2;
        double cx = x + rx;
        double cy = y + ry;
        const double kappa = 0.5522847498307936;
        double ox = rx * kappa;
        double oy = ry * kappa;

        content
            .MoveTo(cx + rx, cy)
            .CubicTo(cx + rx, cy + oy, cx + ox, cy + ry, cx, cy + ry)
            .CubicTo(cx - ox, cy + ry, cx - rx, cy + oy, cx - rx, cy)
            .CubicTo(cx - rx, cy - oy, cx - ox, cy - ry, cx, cy - ry)
            .CubicTo(cx + ox, cy - ry, cx + rx, cy - oy, cx + rx, cy);
    }

    private static void AppendPathCommands(ContentStreamBuilder content, System.Collections.Generic.IReadOnlyList<OfficeIMO.Drawing.OfficePathCommand> commands, double x, double y, double h) {
        for (int i = 0; i < commands.Count; i++) {
            var command = commands[i];
            switch (command.Kind) {
                case OfficeIMO.Drawing.OfficePathCommandKind.MoveTo:
                    content.MoveTo(x + command.Point.X, y + h - command.Point.Y);
                    break;
                case OfficeIMO.Drawing.OfficePathCommandKind.LineTo:
                    content.LineTo(x + command.Point.X, y + h - command.Point.Y);
                    break;
                case OfficeIMO.Drawing.OfficePathCommandKind.CubicBezierTo:
                    content.CubicTo(
                        x + command.ControlPoint1.X,
                        y + h - command.ControlPoint1.Y,
                        x + command.ControlPoint2.X,
                        y + h - command.ControlPoint2.Y,
                        x + command.Point.X,
                        y + h - command.Point.Y);
                    break;
                case OfficeIMO.Drawing.OfficePathCommandKind.Close:
                    content.ClosePath();
                    break;
            }
        }
    }

    private static void AppendLocalPathCommands(ContentStreamBuilder content, System.Collections.Generic.IReadOnlyList<OfficeIMO.Drawing.OfficePathCommand> commands) {
        for (int i = 0; i < commands.Count; i++) {
            var command = commands[i];
            switch (command.Kind) {
                case OfficeIMO.Drawing.OfficePathCommandKind.MoveTo:
                    content.MoveTo(command.Point.X, command.Point.Y);
                    break;
                case OfficeIMO.Drawing.OfficePathCommandKind.LineTo:
                    content.LineTo(command.Point.X, command.Point.Y);
                    break;
                case OfficeIMO.Drawing.OfficePathCommandKind.CubicBezierTo:
                    content.CubicTo(
                        command.ControlPoint1.X,
                        command.ControlPoint1.Y,
                        command.ControlPoint2.X,
                        command.ControlPoint2.Y,
                        command.Point.X,
                        command.Point.Y);
                    break;
                case OfficeIMO.Drawing.OfficePathCommandKind.Close:
                    content.ClosePath();
                    break;
            }
        }
    }

    private static void PaintPath(ContentStreamBuilder content, bool fill, bool stroke, bool closePath) {
        if (closePath) {
            content.ClosePath();
        }

        if (fill && stroke) {
            content.FillStrokePath();
        } else if (fill) {
            content.FillPath();
        } else if (stroke) {
            content.StrokePath();
        } else {
            content.EndPath();
        }
    }

    private static void DrawGradientShape(StringBuilder sb, OfficeIMO.Drawing.OfficeShape shape, string shadingName, double x, double y) {
        if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
            return;
        }

        new ContentStreamBuilder(sb)
            .SaveState();
        AppendShapeClipPath(sb, shape, x, y);
        new ContentStreamBuilder(sb)
            .Shading(shadingName)
            .RestoreState();
    }

    private static void DrawTransformedShape(StringBuilder sb, OfficeIMO.Drawing.OfficeShape shape, PdfColor? fillColor, PdfColor? strokeColor, string? shadingName, double x, double y) {
        bool stroke = strokeColor.HasValue && shape.StrokeWidth > 0;
        bool gradient = !string.IsNullOrEmpty(shadingName) && shape.Kind != OfficeIMO.Drawing.OfficeShapeKind.Line;
        if (!shape.Transform.HasValue || (!fillColor.HasValue && !gradient && !stroke) || (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line && !stroke)) {
            return;
        }

        new ContentStreamBuilder(sb)
            .SaveState();
        ApplyLocalTransform(sb, shape.Transform.Value, x, y, shape.Height);
        if (shape.ClipPath != null) {
            AppendLocalClipPath(sb, shape.ClipPath);
        }

        if (gradient) {
            new ContentStreamBuilder(sb)
                .SaveState();
            AppendLocalShapeClipPath(sb, shape);
            new ContentStreamBuilder(sb)
                .Shading(shadingName!)
                .RestoreState();
        }

        var content = new ContentStreamBuilder(sb);
        bool fill = fillColor.HasValue && !gradient;
        if (fill) {
            content.FillColor(fillColor.GetValueOrDefault());
        }

        if (stroke) {
            content
                .StrokeColor(strokeColor.GetValueOrDefault())
                .LineWidth(shape.StrokeWidth);
            ApplyStrokeStyle(content, shape.StrokeDashStyle, shape.StrokeWidth, shape.StrokeLineCap, shape.StrokeLineJoin);
        }

        switch (shape.Kind) {
            case OfficeIMO.Drawing.OfficeShapeKind.Line:
                if (shape.Points.Count == 2) {
                    content
                        .MoveTo(shape.Points[0].X, shape.Points[0].Y)
                        .LineTo(shape.Points[1].X, shape.Points[1].Y);
                }

                PaintPath(content, fill: false, stroke: true, closePath: false);
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.RoundedRectangle:
                if (shape.CornerRadius <= 0) {
                    content.Rectangle(0, 0, shape.Width, shape.Height);
                    PaintPath(content, fill, stroke, closePath: false);
                } else {
                    AppendRoundedPath(content, 0, 0, shape.Width, shape.Height, shape.CornerRadius);
                    PaintPath(content, fill, stroke, closePath: true);
                }

                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Rectangle:
                content.Rectangle(0, 0, shape.Width, shape.Height);
                PaintPath(content, fill, stroke, closePath: false);
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Ellipse:
                AppendEllipsePath(content, 0, 0, shape.Width, shape.Height);
                PaintPath(content, fill, stroke, closePath: false);
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Polygon:
                AppendLocalPathCommands(content, ConvertPolygonToPath(shape.Points));
                PaintPath(content, fill, stroke, closePath: true);
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Path:
                AppendLocalPathCommands(content, shape.PathCommands);
                PaintPath(content, fill, stroke, closePath: false);
                break;
        }

        content.RestoreState();
    }

    private static void AppendShapeClipPath(StringBuilder sb, OfficeIMO.Drawing.OfficeShape shape, double x, double y) {
        var content = new ContentStreamBuilder(sb);
        switch (shape.Kind) {
            case OfficeIMO.Drawing.OfficeShapeKind.RoundedRectangle:
                if (shape.CornerRadius <= 0) {
                    content.Rectangle(x, y, shape.Width, shape.Height);
                } else {
                    AppendRoundedPath(content, x, y, shape.Width, shape.Height, shape.CornerRadius);
                    content.ClosePath();
                }
                content.ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Rectangle:
                content.Rectangle(x, y, shape.Width, shape.Height).ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Ellipse:
                AppendEllipsePath(content, x, y, shape.Width, shape.Height);
                content.ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Polygon:
                AppendPathCommands(content, ConvertPolygonToPath(shape.Points), x, y, shape.Height);
                content.ClosePath().ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Path:
                AppendPathCommands(content, shape.PathCommands, x, y, shape.Height);
                content.ClipPath().EndPath();
                break;
        }
    }

    private static void AppendLocalShapeClipPath(StringBuilder sb, OfficeIMO.Drawing.OfficeShape shape) {
        var content = new ContentStreamBuilder(sb);
        switch (shape.Kind) {
            case OfficeIMO.Drawing.OfficeShapeKind.RoundedRectangle:
                if (shape.CornerRadius <= 0) {
                    content.Rectangle(0, 0, shape.Width, shape.Height);
                } else {
                    AppendRoundedPath(content, 0, 0, shape.Width, shape.Height, shape.CornerRadius);
                    content.ClosePath();
                }
                content.ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Rectangle:
                content.Rectangle(0, 0, shape.Width, shape.Height).ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Ellipse:
                AppendEllipsePath(content, 0, 0, shape.Width, shape.Height);
                content.ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Polygon:
                AppendLocalPathCommands(content, ConvertPolygonToPath(shape.Points));
                content.ClosePath().ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeShapeKind.Path:
                AppendLocalPathCommands(content, shape.PathCommands);
                content.ClipPath().EndPath();
                break;
        }
    }

    private static System.Collections.Generic.List<OfficeIMO.Drawing.OfficePathCommand> ConvertPolygonToPath(System.Collections.Generic.IReadOnlyList<OfficeIMO.Drawing.OfficePoint> points) {
        var commands = new System.Collections.Generic.List<OfficeIMO.Drawing.OfficePathCommand>(points.Count);
        if (points.Count == 0) {
            return commands;
        }

        commands.Add(OfficeIMO.Drawing.OfficePathCommand.MoveTo(points[0]));
        for (int i = 1; i < points.Count; i++) {
            commands.Add(OfficeIMO.Drawing.OfficePathCommand.LineTo(points[i]));
        }

        return commands;
    }

    private static void AppendClipPath(StringBuilder sb, OfficeIMO.Drawing.OfficeClipPath clipPath, double x, double y, double shapeHeight) {
        var content = new ContentStreamBuilder(sb);
        switch (clipPath.Kind) {
            case OfficeIMO.Drawing.OfficeClipPathKind.Rectangle:
                content.Rectangle(x, y + shapeHeight - clipPath.Height, clipPath.Width, clipPath.Height).ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeClipPathKind.RoundedRectangle:
                if (clipPath.CornerRadius <= 0) {
                    content.Rectangle(x, y + shapeHeight - clipPath.Height, clipPath.Width, clipPath.Height);
                } else {
                    AppendRoundedPath(content, x, y + shapeHeight - clipPath.Height, clipPath.Width, clipPath.Height, clipPath.CornerRadius);
                    content.ClosePath();
                }
                content.ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeClipPathKind.Path:
                AppendPathCommands(content, clipPath.Commands, x, y, shapeHeight);
                content.ClipPath().EndPath();
                break;
        }
    }

    private static void AppendLocalClipPath(StringBuilder sb, OfficeIMO.Drawing.OfficeClipPath clipPath) {
        var content = new ContentStreamBuilder(sb);
        switch (clipPath.Kind) {
            case OfficeIMO.Drawing.OfficeClipPathKind.Rectangle:
                content.Rectangle(0, 0, clipPath.Width, clipPath.Height).ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeClipPathKind.RoundedRectangle:
                if (clipPath.CornerRadius <= 0) {
                    content.Rectangle(0, 0, clipPath.Width, clipPath.Height);
                } else {
                    AppendRoundedPath(content, 0, 0, clipPath.Width, clipPath.Height, clipPath.CornerRadius);
                    content.ClosePath();
                }
                content.ClipPath().EndPath();
                break;
            case OfficeIMO.Drawing.OfficeClipPathKind.Path:
                AppendLocalPathCommands(content, clipPath.Commands);
                content.ClipPath().EndPath();
                break;
        }
    }

    private static void AppendEllipsePath(StringBuilder sb, double x, double y, double w, double h) {
        double rx = w / 2;
        double ry = h / 2;
        double cx = x + rx;
        double cy = y + ry;
        const double kappa = 0.5522847498307936;
        double ox = rx * kappa;
        double oy = ry * kappa;

        sb.Append(F(cx + rx)).Append(' ').Append(F(cy)).Append(" m\n");
        sb.Append(F(cx + rx)).Append(' ').Append(F(cy + oy)).Append(' ').Append(F(cx + ox)).Append(' ').Append(F(cy + ry)).Append(' ').Append(F(cx)).Append(' ').Append(F(cy + ry)).Append(" c\n");
        sb.Append(F(cx - ox)).Append(' ').Append(F(cy + ry)).Append(' ').Append(F(cx - rx)).Append(' ').Append(F(cy + oy)).Append(' ').Append(F(cx - rx)).Append(' ').Append(F(cy)).Append(" c\n");
        sb.Append(F(cx - rx)).Append(' ').Append(F(cy - oy)).Append(' ').Append(F(cx - ox)).Append(' ').Append(F(cy - ry)).Append(' ').Append(F(cx)).Append(' ').Append(F(cy - ry)).Append(" c\n");
        sb.Append(F(cx + ox)).Append(' ').Append(F(cy - ry)).Append(' ').Append(F(cx + rx)).Append(' ').Append(F(cy - oy)).Append(' ').Append(F(cx + rx)).Append(' ').Append(F(cy)).Append(" c\n");
    }

    private static void ApplyLocalTransform(StringBuilder sb, OfficeIMO.Drawing.OfficeTransform transform, double x, double y, double h) {
        new ContentStreamBuilder(sb)
            .TransformMatrix(
                NormalizePdfZero(transform.M11),
                NormalizePdfZero(-transform.M12),
                NormalizePdfZero(transform.M21),
                NormalizePdfZero(-transform.M22),
                x + transform.OffsetX,
                y + h - transform.OffsetY);
    }

    private static double NormalizePdfZero(double value) => Math.Abs(value) < 0.000000000001D ? 0D : value;

    private static void ApplyStrokeStyle(ContentStreamBuilder content, OfficeIMO.Drawing.OfficeStrokeDashStyle dashStyle, double strokeWidth, OfficeIMO.Drawing.OfficeStrokeLineCap? strokeLineCap, OfficeIMO.Drawing.OfficeStrokeLineJoin? strokeLineJoin) {
        if (strokeLineCap.HasValue) {
            content.LineCap(LineCapValue(strokeLineCap.GetValueOrDefault()));
        }

        if (strokeLineJoin.HasValue) {
            content.LineJoin(LineJoinValue(strokeLineJoin.GetValueOrDefault()));
        }

        ApplyStrokeDashStyle(content, dashStyle, strokeWidth, strokeLineCap.HasValue);
    }

    private static void ApplyStrokeDashStyle(ContentStreamBuilder content, OfficeIMO.Drawing.OfficeStrokeDashStyle dashStyle, double strokeWidth, bool hasExplicitLineCap) {
        double unit = Math.Max(0.1, strokeWidth);
        switch (dashStyle) {
            case OfficeIMO.Drawing.OfficeStrokeDashStyle.Dash:
                content.StrokeDash(unit * 3, unit * 1.5);
                break;
            case OfficeIMO.Drawing.OfficeStrokeDashStyle.Dot:
                if (!hasExplicitLineCap) {
                    content.LineCap(1);
                }

                content.StrokeDash(unit, unit * 1.5);
                break;
            case OfficeIMO.Drawing.OfficeStrokeDashStyle.DashDot:
                if (!hasExplicitLineCap) {
                    content.LineCap(1);
                }

                content.StrokeDash(unit * 3, unit * 1.5, unit, unit * 1.5);
                break;
        }
    }

    private static int LineCapValue(OfficeIMO.Drawing.OfficeStrokeLineCap lineCap) {
        switch (lineCap) {
            case OfficeIMO.Drawing.OfficeStrokeLineCap.Butt:
                return 0;
            case OfficeIMO.Drawing.OfficeStrokeLineCap.Round:
                return 1;
            case OfficeIMO.Drawing.OfficeStrokeLineCap.Square:
                return 2;
            default:
                throw new System.ArgumentOutOfRangeException(nameof(lineCap), "Unsupported stroke line cap.");
        }
    }

    private static int LineJoinValue(OfficeIMO.Drawing.OfficeStrokeLineJoin lineJoin) {
        switch (lineJoin) {
            case OfficeIMO.Drawing.OfficeStrokeLineJoin.Miter:
                return 0;
            case OfficeIMO.Drawing.OfficeStrokeLineJoin.Round:
                return 1;
            case OfficeIMO.Drawing.OfficeStrokeLineJoin.Bevel:
                return 2;
            default:
                throw new System.ArgumentOutOfRangeException(nameof(lineJoin), "Unsupported stroke line join.");
        }
    }

    private static void DrawVLine(StringBuilder sb, PdfColor color, double widthStroke, double x, double yTop, double yBottom) {
        new ContentStreamBuilder(sb)
            .SaveState()
            .StrokeColor(color)
            .LineWidth(widthStroke)
            .MoveTo(x, yTop)
            .LineTo(x, yBottom)
            .StrokePath()
            .RestoreState();
    }

    private static void DrawHLine(StringBuilder sb, PdfColor color, double widthStroke, double x1, double x2, double y) {
        new ContentStreamBuilder(sb)
            .SaveState()
            .StrokeColor(color)
            .LineWidth(widthStroke)
            .MoveTo(x1, y)
            .LineTo(x2, y)
            .StrokePath()
            .RestoreState();
    }

    private static void DrawCellBorder(StringBuilder sb, PdfCellBorder border, double x, double y, double w, double h) {
        if (!border.Color.HasValue || border.Width <= 0) {
            return;
        }

        if (border.Top && border.Right && border.Bottom && border.Left) {
            DrawRowRect(sb, border.Color.Value, border.Width, x, y, w, h);
            return;
        }

        double x2 = x + w;
        double y2 = y + h;
        if (border.Top) DrawHLine(sb, border.Color.Value, border.Width, x, x2, y2);
        if (border.Right) DrawVLine(sb, border.Color.Value, border.Width, x2, y2, y);
        if (border.Bottom) DrawHLine(sb, border.Color.Value, border.Width, x, x2, y);
        if (border.Left) DrawVLine(sb, border.Color.Value, border.Width, x, y2, y);
    }

    private static void WriteCell(StringBuilder sb, string fontRes, double fontSize, double x, double y, string text, PdfColor? color, PdfOptions opts) {
        var effective = color ?? opts.DefaultTextColor;
        new ContentStreamBuilder(sb)
            .BeginText()
            .Font(fontRes, fontSize)
            .FillColor(effective ?? PdfColor.Black)
            .TextMatrix(x, y)
            .ShowHexText(EncodeWinAnsiHex(text))
            .EndText();
    }

    private static void WriteClippedCell(StringBuilder sb, string fontRes, double fontSize, double x, double y, string text, PdfColor? color, PdfOptions opts, double clipX, double clipY, double clipWidth, double clipHeight) {
        new ContentStreamBuilder(sb)
            .SaveState()
            .Rectangle(clipX, clipY, clipWidth, clipHeight)
            .ClipPath()
            .EndPath();

        WriteCell(sb, fontRes, fontSize, x, y, text, color, opts);

        new ContentStreamBuilder(sb)
            .RestoreState();
    }
}
