using System;
using System.Collections.Generic;
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
            .Append("pt\" height=\"")
            .Append(Format(drawing.Height))
            .Append("pt\" viewBox=\"0 0 ")
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
                        sb.AppendLinearGradientDefinition(fillGradientId, drawingShape.Shape.FillGradient);
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
        if (TryCreateShadowShape(drawingShape, out var shadowShape)) {
            AppendShapeGeometry(sb, shadowShape, fillGradientId: null, clipPathId);
        }

        AppendShapeGeometry(sb, drawingShape, fillGradientId, clipPathId);
    }

    private static void AppendShapeGeometry(StringBuilder sb, OfficeDrawingShape drawingShape, string? fillGradientId, string? clipPathId) {
        OfficeShape shape = drawingShape.Shape;
        string paint = BuildPaintAttributes(shape, fillGradientId);
        bool useLocalCoordinates = clipPathId != null || HasNonIdentityTransform(shape.Transform);
        double originX = useLocalCoordinates ? 0D : drawingShape.X;
        double originY = useLocalCoordinates ? 0D : drawingShape.Y;
        string transform = clipPathId == null
            ? BuildTransformAttribute(shape.Transform, drawingShape.X, drawingShape.Y)
            : string.Empty;

        if (clipPathId != null) {
            sb.Append("<g")
                .AppendClipPathReference(clipPathId)
                .Append(BuildPlacementTransformAttribute(shape.Transform, drawingShape.X, drawingShape.Y))
                .Append('>');
        }

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
                sb.AppendRectElement(originX, originY, shape.Width, shape.Height, paint + transform);
                break;
            case OfficeShapeKind.RoundedRectangle:
                sb.AppendRectElement(originX, originY, shape.Width, shape.Height, shape.CornerRadius, shape.CornerRadius, paint + transform);
                break;
            case OfficeShapeKind.Ellipse:
                sb.AppendEllipseElement(
                    originX + shape.Width / 2D,
                    originY + shape.Height / 2D,
                    shape.Width / 2D,
                    shape.Height / 2D,
                    paint + transform);
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

    private static bool TryCreateShadowShape(OfficeDrawingShape drawingShape, out OfficeDrawingShape shadowDrawingShape) {
        OfficeShape shape = drawingShape.Shape;
        OfficeShadow? shadow = shape.Shadow;
        if (shadow == null || shadow.Opacity <= 0D || shadow.Color.A == 0) {
            shadowDrawingShape = drawingShape;
            return false;
        }

        bool hasStroke = shape.Kind == OfficeShapeKind.Line ||
            (shape.StrokeColor.HasValue && shape.StrokeWidth > 0D && shape.StrokeColor.Value.A > 0);
        bool hasFill = shape.Kind != OfficeShapeKind.Line &&
            (shape.FillGradient != null || (shape.FillColor.HasValue && shape.FillColor.Value.A > 0));

        var shadowShape = shape.Clone();
        shadowShape.Shadow = null;
        shadowShape.FillGradient = null;
        shadowShape.FillColor = hasFill || !hasStroke ? shadow.Color : null;
        shadowShape.FillOpacity = shadow.Opacity;
        shadowShape.StrokeColor = hasStroke ? shadow.Color : null;
        shadowShape.StrokeOpacity = shadow.Opacity;

        shadowDrawingShape = new OfficeDrawingShape(
            shadowShape,
            drawingShape.X + shadow.OffsetX,
            drawingShape.Y + shadow.OffsetY);
        return true;
    }

    private static void AppendLine(StringBuilder sb, OfficeDrawingShape drawingShape, string paint, string transform, double originX, double originY) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.Points.Count != 2) {
            return;
        }

        sb.AppendLineElement(
            originX + shape.Points[0].X,
            originY + shape.Points[0].Y,
            originX + shape.Points[1].X,
            originY + shape.Points[1].Y,
            paint + transform);
    }

    private static void AppendPolygon(StringBuilder sb, OfficeDrawingShape drawingShape, string paint, string transform, double originX, double originY) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.Points.Count < 3) {
            return;
        }

        List<OfficePoint> points = OffsetPoints(shape.Points, originX, originY);
        sb.AppendPolygonElement(points, paint + transform);
    }

    private static void AppendPath(StringBuilder sb, OfficeDrawingShape drawingShape, string paint, string transform, double originX, double originY) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.PathCommands.Count == 0) {
            return;
        }

        sb.AppendPathElement(shape.PathCommands, originX, originY, paint + transform);
    }

    private static void AppendText(StringBuilder sb, OfficeDrawingText text) {
        double x = text.X;
        if (text.Alignment == OfficeTextAlignment.Center) {
            x += text.Width / 2D;
        } else if (text.Alignment == OfficeTextAlignment.Right) {
            x += text.Width;
        }

        double fontSize = text.Font.Size > 0 ? text.Font.Size : 10D;
        double y = text.Y + fontSize;
        double lineHeight = text.LineHeight ?? fontSize * 1.2D;
        sb.AppendSvgTextElement(
            text.Text,
            x,
            y,
            lineHeight,
            text.Color ?? OfficeColor.Black,
            text.Font.FamilyName ?? "Arial",
            fontSize,
            text.Alignment,
            text.Font.IsBold,
            text.Font.IsItalic);
    }

    private static void AppendClipPathDefinition(StringBuilder sb, string id, OfficeClipPath clipPath) {
        sb.Append("<defs><clipPath id=\"")
            .Append(Escape(id))
            .Append("\">");

        switch (clipPath.Kind) {
            case OfficeClipPathKind.Rectangle:
                sb.AppendRectElement(0D, 0D, clipPath.Width, clipPath.Height);
                break;
            case OfficeClipPathKind.RoundedRectangle:
                sb.AppendRectElement(0D, 0D, clipPath.Width, clipPath.Height, clipPath.CornerRadius, clipPath.CornerRadius);
                break;
            case OfficeClipPathKind.Path:
                AppendClipPathPath(sb, clipPath);
                break;
        }

        sb.Append("</clipPath></defs>");
    }

    private static void AppendClipPathPath(StringBuilder sb, OfficeClipPath clipPath) {
        sb.AppendPathElement(clipPath.Commands);
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
        sb.AppendStrokeDashStyleAttribute(shape.StrokeDashStyle, shape.StrokeWidth);

        if (shape.StrokeLineCap.HasValue) {
            sb.AppendStrokeLineCapAttribute(shape.StrokeLineCap.Value);
        }

        if (shape.StrokeLineJoin.HasValue) {
            sb.AppendStrokeLineJoinAttribute(shape.StrokeLineJoin.Value);
        }
    }

    private static List<OfficePoint> OffsetPoints(IReadOnlyList<OfficePoint> points, double offsetX, double offsetY) {
        var translated = new List<OfficePoint>(points.Count);
        for (int i = 0; i < points.Count; i++) {
            translated.Add(new OfficePoint(points[i].X + offsetX, points[i].Y + offsetY));
        }

        return translated;
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
        var builder = new StringBuilder();
        builder.AppendMatrixTransformAttribute(value, placementX, placementY);
        return builder.ToString();
    }

    private static string ToCssColor(OfficeColor color) => OfficeSvgFormatting.ToCssColor(color);

    private static double ToOpacity(OfficeColor color) => OfficeSvgFormatting.ToOpacity(color);

    private static string Format(double value) => OfficeSvgFormatting.FormatNumber(value);

    private static string Escape(string? value) => OfficeSvgFormatting.Escape(value);
}
