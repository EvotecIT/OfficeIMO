using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Exports dependency-free OfficeIMO drawings to SVG for consumers that need a portable visual fallback.
/// </summary>
public static class OfficeDrawingSvgExporter {
    /// <summary>
    /// Converts a drawing to an SVG document.
    /// </summary>
    /// <param name="drawing">Drawing to export.</param>
    /// <returns>SVG markup representing the drawing.</returns>
    public static string ToSvg(OfficeDrawing drawing) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        var sb = new StringBuilder();
        sb.Append("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"")
            .Append(Format(drawing.Width))
            .Append("\" height=\"")
            .Append(Format(drawing.Height))
            .Append("\" viewBox=\"0 0 ")
            .Append(Format(drawing.Width))
            .Append(' ')
            .Append(Format(drawing.Height))
            .Append("\" role=\"img\">");

        int gradientId = 0;
        int clipPathId = 0;
        for (int i = 0; i < drawing.Elements.Count; i++) {
            switch (drawing.Elements[i]) {
                case OfficeDrawingShape drawingShape:
                    string? fillGradientId = null;
                    if (drawingShape.Shape.FillGradient != null) {
                        fillGradientId = "officeimo-gradient-" + (++gradientId).ToString(CultureInfo.InvariantCulture);
                        AppendGradientDefinition(sb, fillGradientId, drawingShape.Shape.FillGradient);
                    }

                    string? shapeClipPathId = null;
                    if (drawingShape.Shape.ClipPath != null) {
                        shapeClipPathId = "officeimo-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
                        AppendClipPathDefinition(sb, shapeClipPathId, drawingShape.Shape.ClipPath);
                    }

                    AppendShape(sb, drawingShape, fillGradientId, shapeClipPathId);
                    break;
                case OfficeDrawingText drawingText:
                    AppendText(sb, drawingText);
                    break;
            }
        }

        sb.Append("</svg>");
        return sb.ToString();
    }

    /// <summary>
    /// Converts a drawing to UTF-8 SVG bytes.
    /// </summary>
    /// <param name="drawing">Drawing to export.</param>
    /// <returns>UTF-8 encoded SVG bytes.</returns>
    public static byte[] ToSvgBytes(OfficeDrawing drawing) => Encoding.UTF8.GetBytes(ToSvg(drawing));

    private static void AppendShape(StringBuilder sb, OfficeDrawingShape drawingShape, string? fillGradientId, string? clipPathId) {
        OfficeShape shape = drawingShape.Shape;
        string paint = BuildPaintAttributes(shape, fillGradientId);
        bool useLocalCoordinates = clipPathId != null || HasNonIdentityTransform(shape.Transform);
        double originX = useLocalCoordinates ? 0D : drawingShape.X;
        double originY = useLocalCoordinates ? 0D : drawingShape.Y;
        string transform = clipPathId == null
            ? BuildTransformAttribute(shape.Transform, drawingShape.X, drawingShape.Y)
            : string.Empty;

        if (clipPathId != null) {
            sb.Append("<g clip-path=\"url(#")
                .Append(Escape(clipPathId))
                .Append(")\"")
                .Append(BuildPlacementTransformAttribute(shape.Transform, drawingShape.X, drawingShape.Y))
                .Append('>');
        }

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
                sb.Append("<rect x=\"").Append(Format(originX))
                    .Append("\" y=\"").Append(Format(originY))
                    .Append("\" width=\"").Append(Format(shape.Width))
                    .Append("\" height=\"").Append(Format(shape.Height))
                    .Append('"').Append(paint).Append(transform).Append("/>");
                break;
            case OfficeShapeKind.RoundedRectangle:
                sb.Append("<rect x=\"").Append(Format(originX))
                    .Append("\" y=\"").Append(Format(originY))
                    .Append("\" width=\"").Append(Format(shape.Width))
                    .Append("\" height=\"").Append(Format(shape.Height))
                    .Append("\" rx=\"").Append(Format(shape.CornerRadius))
                    .Append("\" ry=\"").Append(Format(shape.CornerRadius))
                    .Append('"').Append(paint).Append(transform).Append("/>");
                break;
            case OfficeShapeKind.Ellipse:
                sb.Append("<ellipse cx=\"").Append(Format(originX + shape.Width / 2D))
                    .Append("\" cy=\"").Append(Format(originY + shape.Height / 2D))
                    .Append("\" rx=\"").Append(Format(shape.Width / 2D))
                    .Append("\" ry=\"").Append(Format(shape.Height / 2D))
                    .Append('"').Append(paint).Append(transform).Append("/>");
                break;
            case OfficeShapeKind.Line:
                AppendLine(sb, drawingShape, paint, transform, originX, originY);
                break;
            case OfficeShapeKind.Polygon:
                AppendPolygon(sb, drawingShape, paint, transform, originX, originY);
                break;
            case OfficeShapeKind.Path:
                AppendPath(sb, drawingShape, paint, transform, originX, originY);
                break;
        }

        if (clipPathId != null) {
            sb.Append("</g>");
        }
    }

    private static void AppendLine(StringBuilder sb, OfficeDrawingShape drawingShape, string paint, string transform, double originX, double originY) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.Points.Count != 2) {
            return;
        }

        sb.Append("<line x1=\"").Append(Format(originX + shape.Points[0].X))
            .Append("\" y1=\"").Append(Format(originY + shape.Points[0].Y))
            .Append("\" x2=\"").Append(Format(originX + shape.Points[1].X))
            .Append("\" y2=\"").Append(Format(originY + shape.Points[1].Y))
            .Append('"').Append(paint).Append(transform).Append("/>");
    }

    private static void AppendPolygon(StringBuilder sb, OfficeDrawingShape drawingShape, string paint, string transform, double originX, double originY) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.Points.Count < 3) {
            return;
        }

        sb.Append("<polygon points=\"");
        for (int i = 0; i < shape.Points.Count; i++) {
            if (i > 0) {
                sb.Append(' ');
            }

            sb.Append(Format(originX + shape.Points[i].X))
                .Append(',')
                .Append(Format(originY + shape.Points[i].Y));
        }

        sb.Append('"').Append(paint).Append(transform).Append("/>");
    }

    private static void AppendPath(StringBuilder sb, OfficeDrawingShape drawingShape, string paint, string transform, double originX, double originY) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.PathCommands.Count == 0) {
            return;
        }

        sb.Append("<path d=\"");
        for (int i = 0; i < shape.PathCommands.Count; i++) {
            OfficePathCommand command = shape.PathCommands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    sb.Append('M').Append(Format(originX + command.Point.X)).Append(' ').Append(Format(originY + command.Point.Y));
                    break;
                case OfficePathCommandKind.LineTo:
                    sb.Append('L').Append(Format(originX + command.Point.X)).Append(' ').Append(Format(originY + command.Point.Y));
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    sb.Append('C')
                        .Append(Format(originX + command.ControlPoint1.X)).Append(' ').Append(Format(originY + command.ControlPoint1.Y)).Append(' ')
                        .Append(Format(originX + command.ControlPoint2.X)).Append(' ').Append(Format(originY + command.ControlPoint2.Y)).Append(' ')
                        .Append(Format(originX + command.Point.X)).Append(' ').Append(Format(originY + command.Point.Y));
                    break;
                case OfficePathCommandKind.Close:
                    sb.Append('Z');
                    break;
            }
        }

        sb.Append('"').Append(paint).Append(transform).Append("/>");
    }

    private static void AppendText(StringBuilder sb, OfficeDrawingText text) {
        double x = text.X;
        string anchor = "start";
        if (text.Alignment == OfficeTextAlignment.Center) {
            x += text.Width / 2D;
            anchor = "middle";
        } else if (text.Alignment == OfficeTextAlignment.Right) {
            x += text.Width;
            anchor = "end";
        }

        double fontSize = text.Font.Size > 0 ? text.Font.Size : 10D;
        double y = text.Y + fontSize;

        sb.Append("<text x=\"").Append(Format(x))
            .Append("\" y=\"").Append(Format(y))
            .Append("\" font-family=\"").Append(Escape(text.Font.FamilyName ?? "Arial"))
            .Append("\" font-size=\"").Append(Format(fontSize))
            .Append("\" text-anchor=\"").Append(anchor)
            .Append("\" fill=\"").Append(ToCssColor(text.Color ?? OfficeColor.Black)).Append('"');

        if (text.Font.IsBold || text.Font.IsItalic) {
            if (text.Font.IsBold) {
                sb.Append(" font-weight=\"700\"");
            }

            if (text.Font.IsItalic) {
                sb.Append(" font-style=\"italic\"");
            }
        }

        sb.Append('>');
        string[] lines = text.Text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        double lineHeight = text.LineHeight ?? fontSize * 1.2D;
        for (int i = 0; i < lines.Length; i++) {
            if (i == 0) {
                sb.Append(Escape(lines[i]));
            } else {
                sb.Append("<tspan x=\"").Append(Format(x))
                    .Append("\" dy=\"").Append(Format(lineHeight))
                    .Append("\">").Append(Escape(lines[i])).Append("</tspan>");
            }
        }

        sb.Append("</text>");
    }

    private static void AppendGradientDefinition(StringBuilder sb, string id, OfficeLinearGradient gradient) {
        sb.Append("<defs><linearGradient id=\"")
            .Append(Escape(id))
            .Append("\" x1=\"")
            .Append(Format(gradient.StartX * 100D))
            .Append("%\" y1=\"")
            .Append(Format(gradient.StartY * 100D))
            .Append("%\" x2=\"")
            .Append(Format(gradient.EndX * 100D))
            .Append("%\" y2=\"")
            .Append(Format(gradient.EndY * 100D))
            .Append("%\">");

        for (int i = 0; i < gradient.Stops.Count; i++) {
            OfficeGradientStop stop = gradient.Stops[i];
            sb.Append("<stop offset=\"")
                .Append(Format(stop.Offset * 100D))
                .Append("%\" stop-color=\"")
                .Append(ToCssColor(stop.Color))
                .Append('"');

            double opacity = ToOpacity(stop.Color);
            if (opacity < 1D) {
                sb.Append(" stop-opacity=\"").Append(Format(opacity)).Append('"');
            }

            sb.Append("/>");
        }

        sb.Append("</linearGradient></defs>");
    }

    private static void AppendClipPathDefinition(StringBuilder sb, string id, OfficeClipPath clipPath) {
        sb.Append("<defs><clipPath id=\"")
            .Append(Escape(id))
            .Append("\">");

        switch (clipPath.Kind) {
            case OfficeClipPathKind.Rectangle:
                sb.Append("<rect x=\"0\" y=\"0\" width=\"")
                    .Append(Format(clipPath.Width))
                    .Append("\" height=\"")
                    .Append(Format(clipPath.Height))
                    .Append("\"/>");
                break;
            case OfficeClipPathKind.RoundedRectangle:
                sb.Append("<rect x=\"0\" y=\"0\" width=\"")
                    .Append(Format(clipPath.Width))
                    .Append("\" height=\"")
                    .Append(Format(clipPath.Height))
                    .Append("\" rx=\"")
                    .Append(Format(clipPath.CornerRadius))
                    .Append("\" ry=\"")
                    .Append(Format(clipPath.CornerRadius))
                    .Append("\"/>");
                break;
            case OfficeClipPathKind.Path:
                AppendClipPathPath(sb, clipPath);
                break;
        }

        sb.Append("</clipPath></defs>");
    }

    private static void AppendClipPathPath(StringBuilder sb, OfficeClipPath clipPath) {
        sb.Append("<path d=\"");
        for (int i = 0; i < clipPath.Commands.Count; i++) {
            OfficePathCommand command = clipPath.Commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    sb.Append('M').Append(Format(command.Point.X)).Append(' ').Append(Format(command.Point.Y));
                    break;
                case OfficePathCommandKind.LineTo:
                    sb.Append('L').Append(Format(command.Point.X)).Append(' ').Append(Format(command.Point.Y));
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    sb.Append('C')
                        .Append(Format(command.ControlPoint1.X)).Append(' ').Append(Format(command.ControlPoint1.Y)).Append(' ')
                        .Append(Format(command.ControlPoint2.X)).Append(' ').Append(Format(command.ControlPoint2.Y)).Append(' ')
                        .Append(Format(command.Point.X)).Append(' ').Append(Format(command.Point.Y));
                    break;
                case OfficePathCommandKind.Close:
                    sb.Append('Z');
                    break;
            }
        }

        sb.Append("\"/>");
    }

    private static string BuildPaintAttributes(OfficeShape shape, string? fillGradientId) {
        var sb = new StringBuilder();
        if (fillGradientId != null) {
            sb.Append(" fill=\"url(#").Append(Escape(fillGradientId)).Append(")\"");
            double fillOpacity = shape.FillOpacity ?? 1D;
            if (fillOpacity < 1D) {
                sb.Append(" fill-opacity=\"").Append(Format(fillOpacity)).Append('"');
            }
        } else if (shape.FillColor.HasValue && shape.FillColor.Value.A > 0) {
            sb.Append(" fill=\"").Append(ToCssColor(shape.FillColor.Value)).Append('"');
            double fillOpacity = shape.FillOpacity ?? ToOpacity(shape.FillColor.Value);
            if (fillOpacity < 1D) {
                sb.Append(" fill-opacity=\"").Append(Format(fillOpacity)).Append('"');
            }
        } else {
            sb.Append(" fill=\"none\"");
        }

        if (shape.StrokeColor.HasValue && shape.StrokeWidth > 0 && shape.StrokeColor.Value.A > 0) {
            sb.Append(" stroke=\"").Append(ToCssColor(shape.StrokeColor.Value)).Append('"')
                .Append(" stroke-width=\"").Append(Format(shape.StrokeWidth)).Append('"');
            double strokeOpacity = shape.StrokeOpacity ?? ToOpacity(shape.StrokeColor.Value);
            if (strokeOpacity < 1D) {
                sb.Append(" stroke-opacity=\"").Append(Format(strokeOpacity)).Append('"');
            }

            AppendStrokeStyle(sb, shape);
        } else {
            sb.Append(" stroke=\"none\"");
        }

        return sb.ToString();
    }

    private static void AppendStrokeStyle(StringBuilder sb, OfficeShape shape) {
        switch (shape.StrokeDashStyle) {
            case OfficeStrokeDashStyle.Dash:
                sb.Append(" stroke-dasharray=\"").Append(Format(shape.StrokeWidth * 4D)).Append(' ').Append(Format(shape.StrokeWidth * 2D)).Append('"');
                break;
            case OfficeStrokeDashStyle.Dot:
                sb.Append(" stroke-dasharray=\"").Append(Format(shape.StrokeWidth)).Append(' ').Append(Format(shape.StrokeWidth * 2D)).Append('"');
                break;
            case OfficeStrokeDashStyle.DashDot:
                sb.Append(" stroke-dasharray=\"")
                    .Append(Format(shape.StrokeWidth * 4D)).Append(' ')
                    .Append(Format(shape.StrokeWidth * 2D)).Append(' ')
                    .Append(Format(shape.StrokeWidth)).Append(' ')
                    .Append(Format(shape.StrokeWidth * 2D)).Append('"');
                break;
        }

        if (shape.StrokeLineCap.HasValue) {
            sb.Append(" stroke-linecap=\"").Append(MapLineCap(shape.StrokeLineCap.Value)).Append('"');
        }

        if (shape.StrokeLineJoin.HasValue) {
            sb.Append(" stroke-linejoin=\"").Append(MapLineJoin(shape.StrokeLineJoin.Value)).Append('"');
        }
    }

    private static bool HasNonIdentityTransform(OfficeTransform? transform) => transform.HasValue && transform.Value != OfficeTransform.Identity;

    private static string BuildTransformAttribute(OfficeTransform? transform, double placementX, double placementY) {
        if (!HasNonIdentityTransform(transform)) {
            return string.Empty;
        }

        OfficeTransform value = transform!.Value;
        return BuildMatrixTransformAttribute(value, placementX, placementY);
    }

    private static string BuildPlacementTransformAttribute(OfficeTransform? transform, double placementX, double placementY) {
        OfficeTransform value = transform ?? OfficeTransform.Identity;
        if (value == OfficeTransform.Identity && placementX == 0D && placementY == 0D) {
            return string.Empty;
        }

        return BuildMatrixTransformAttribute(value, placementX, placementY);
    }

    private static string BuildMatrixTransformAttribute(OfficeTransform value, double placementX, double placementY) {
        return " transform=\"matrix(" +
            Format(value.M11) + " " +
            Format(value.M12) + " " +
            Format(value.M21) + " " +
            Format(value.M22) + " " +
            Format(value.OffsetX + placementX) + " " +
            Format(value.OffsetY + placementY) + ")\"";
    }

    private static string MapLineCap(OfficeStrokeLineCap cap) {
        switch (cap) {
            case OfficeStrokeLineCap.Round:
                return "round";
            case OfficeStrokeLineCap.Square:
                return "square";
            default:
                return "butt";
        }
    }

    private static string MapLineJoin(OfficeStrokeLineJoin join) {
        switch (join) {
            case OfficeStrokeLineJoin.Bevel:
                return "bevel";
            case OfficeStrokeLineJoin.Round:
                return "round";
            default:
                return "miter";
        }
    }

    private static string ToCssColor(OfficeColor color) => "#" + color.ToRgbHex();

    private static double ToOpacity(OfficeColor color) => color.A / 255D;

    private static string Format(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    private static string Escape(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        return value!
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }
}
