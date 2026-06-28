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
        AppendElements(sb, drawing.Elements, ref gradientId, ref clipPathId);

        sb.Append("</svg>");
        return sb.ToString();
    }

    /// <summary>
    /// Converts a drawing to UTF-8 SVG bytes.
    /// </summary>
    /// <param name="drawing">Drawing to export.</param>
    /// <returns>UTF-8 encoded SVG bytes.</returns>
    public static byte[] ToSvgBytes(OfficeDrawing drawing) => Encoding.UTF8.GetBytes(ToSvg(drawing));

    private static void AppendElements(StringBuilder sb, IReadOnlyList<OfficeDrawingElement> elements, ref int gradientId, ref int clipPathId) {
        for (int i = 0; i < elements.Count; i++) {
            switch (elements[i]) {
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
                case OfficeDrawingRichText drawingRichText:
                    AppendRichText(sb, drawingRichText);
                    break;
                case OfficeDrawingImage drawingImage:
                    string? imageClipPathId = drawingImage.Projection.HasCrop
                        ? "officeimo-image-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture)
                        : null;
                    AppendImage(sb, drawingImage, imageClipPathId);
                    break;
                case OfficeDrawingGroup drawingGroup:
                    AppendGroup(sb, drawingGroup, ref gradientId, ref clipPathId);
                    break;
            }
        }
    }

    private static void AppendGroup(StringBuilder sb, OfficeDrawingGroup drawingGroup, ref int gradientId, ref int clipPathId) {
        string groupClipPathId = "officeimo-group-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
        AppendClipPathDefinition(sb, groupClipPathId, drawingGroup.ClipPath);
        string transform = BuildGroupTransformAttribute(drawingGroup);
        sb.Append("<g")
            .AppendClipPathReference(groupClipPathId)
            .Append(transform)
            .Append('>');
        AppendElements(sb, drawingGroup.InnerDrawing.Elements, ref gradientId, ref clipPathId);
        sb.Append("</g>");
    }

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

    private static void AppendImage(StringBuilder sb, OfficeDrawingImage drawingImage, string? clipPathId) {
        byte[] bytes = drawingImage.Bytes;
        if (!OfficeSvgImageRenderer.TryCreateDataUri(drawingImage.ContentType, bytes, null, out string dataUri)) {
            return;
        }

        OfficeSvgImageRenderer.AppendImage(
            sb,
            dataUri,
            drawingImage.Projection,
            clipPathId,
            drawingImage.Projection.HasCrop ? drawingImage.Projection.Placement : null,
            "none");
    }

    private static void AppendLine(StringBuilder sb, OfficeDrawingShape drawingShape, string paint, string transform, double originX, double originY) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.Points.Count != 2) {
            return;
        }

        OfficePoint start = new OfficePoint(originX + shape.Points[0].X, originY + shape.Points[0].Y);
        OfficePoint end = new OfficePoint(originX + shape.Points[1].X, originY + shape.Points[1].Y);
        sb.AppendLineElement(
            start.X,
            start.Y,
            end.X,
            end.Y,
            paint + transform);
        AppendLineMarker(sb, shape.StrokeStartMarker, start, new OfficePoint(start.X - end.X, start.Y - end.Y), shape, transform);
        AppendLineMarker(sb, shape.StrokeEndMarker, end, new OfficePoint(end.X - start.X, end.Y - start.Y), shape, transform);
    }

    private static void AppendLineMarker(StringBuilder sb, OfficeLineMarker? marker, OfficePoint tip, OfficePoint lineDirection, OfficeShape shape, string transform) {
        IReadOnlyList<OfficePoint> contour = OfficeLineMarkerGeometry.CreateContour(marker, tip, lineDirection);
        if (contour.Count == 0) {
            return;
        }

        string? paint = BuildLineMarkerPaintAttributes(shape);
        if (paint == null) {
            return;
        }

        sb.AppendPolygonElement(contour, paint + transform);
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
        AppendPathMarkers(sb, shape, originX, originY, transform);
    }

    private static void AppendPathMarkers(StringBuilder sb, OfficeShape shape, double originX, double originY, string transform) {
        if (shape.StrokeStartMarker == null && shape.StrokeEndMarker == null) {
            return;
        }

        OfficeFlattenedPathContour? firstOpen = null;
        OfficeFlattenedPathContour? lastOpen = null;
        IReadOnlyList<OfficeFlattenedPathContour> contours = OfficePathFlattener.Flatten(shape.PathCommands, originX, originY, 1D);
        for (int i = 0; i < contours.Count; i++) {
            if (!contours[i].Closed && contours[i].Points.Count >= 2) {
                firstOpen ??= contours[i];
                lastOpen = contours[i];
            }
        }

        if (firstOpen != null) {
            OfficePoint start = firstOpen.Points[0];
            OfficePoint next = firstOpen.Points[1];
            AppendLineMarker(sb, shape.StrokeStartMarker, start, new OfficePoint(start.X - next.X, start.Y - next.Y), shape, transform);
        }

        if (lastOpen != null) {
            IReadOnlyList<OfficePoint> points = lastOpen.Points;
            OfficePoint end = points[points.Count - 1];
            OfficePoint previous = points[points.Count - 2];
            AppendLineMarker(sb, shape.StrokeEndMarker, end, new OfficePoint(end.X - previous.X, end.Y - previous.Y), shape, transform);
        }
    }

    private static void AppendText(StringBuilder sb, OfficeDrawingText text) {
        bool useFrameTransform = text.FlipHorizontal || text.FlipVertical;
        if (useFrameTransform) {
            AppendTextFrameGroupStart(sb, text);
        }

        if (text.WrapText || text.ShrinkToFit || text.VerticalAlignment != OfficeTextVerticalAlignment.Top || text.HasPadding) {
            AppendTextBlock(sb, text, useFrameTransform);
            if (useFrameTransform) {
                sb.Append("</g>");
            }

            return;
        }

        double contentX = text.X + text.Padding.Left;
        double contentY = text.Y + text.Padding.Top;
        double contentWidth = text.Width - text.Padding.Horizontal;
        double x = contentX;
        if (text.Alignment == OfficeTextAlignment.Center) {
            x += contentWidth / 2D;
        } else if (text.Alignment == OfficeTextAlignment.Right) {
            x += contentWidth;
        }

        double fontSize = text.Font.Size > 0 ? text.Font.Size : 10D;
        double y = contentY + fontSize;
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
            text.Font.IsItalic,
            (text.Font.Style & OfficeFontStyle.Underline) == OfficeFontStyle.Underline,
            useFrameTransform ? 0D : text.RotationDegrees,
            useFrameTransform ? 0D : text.RotationCenterX,
            useFrameTransform ? 0D : text.RotationCenterY,
            (text.Font.Style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough);

        if (useFrameTransform) {
            sb.Append("</g>");
        }
    }

    private static void AppendTextBlock(StringBuilder sb, OfficeDrawingText text, bool useFrameTransform = false) {
        double fontSize = text.Font.Size > 0 ? text.Font.Size : 10D;
        double lineHeightFactor = text.LineHeight.HasValue && text.LineHeight.Value > 0D
            ? Math.Max(1D, text.LineHeight.Value / fontSize)
            : 1.2D;
        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(text.Font);
        OfficeTextMeasurementStyle style = measurer.CreateStyle(new OfficeFontInfo(text.Font.FamilyName, fontSize, text.Font.Style));
        double minimumFontSize = Math.Min(6D, fontSize);
        Func<string?, double, double> measure = (value, size) => {
                OfficeTextMeasurementStyle measuredStyle = measurer.CreateStyle(new OfficeFontInfo(text.Font.FamilyName, size, text.Font.Style));
                return measurer.MeasureWidth(value, measuredStyle);
            };
        double contentX = text.X + text.Padding.Left;
        double contentY = text.Y + text.Padding.Top;
        double contentWidth = text.Width - text.Padding.Horizontal;
        double contentHeight = text.Height - text.Padding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            return;
        }

        OfficeTextBlockLayout layout = text.StackedText
            ? OfficeTextLayoutEngine.LayoutStackedTextBlock(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                text.ShrinkToFit)
            : text.ShrinkToFit && text.WrapText
            ? OfficeTextLayoutEngine.FitWrappedText(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                text.ParagraphIndent)
            : OfficeTextLayoutEngine.LayoutTextBlock(
                text.Text,
                fontSize,
                contentWidth,
                contentHeight,
                lineHeightFactor,
                minimumFontSize,
                measure,
                wrap: text.WrapText,
                shrinkToFit: text.ShrinkToFit,
                paragraphIndent: text.ParagraphIndent);
        sb.AppendSvgTextBlock(
            layout,
            contentX,
            contentY,
            contentWidth,
            contentHeight,
            text.Color ?? OfficeColor.Black,
            string.IsNullOrWhiteSpace(style.FontInfo.FamilyName) ? text.Font.FamilyName : style.FontInfo.FamilyName,
            text.Alignment,
            text.VerticalAlignment,
            text.Font.IsBold,
            text.Font.IsItalic,
            (text.Font.Style & OfficeFontStyle.Underline) == OfficeFontStyle.Underline,
            useFrameTransform ? 0D : text.RotationDegrees,
            useFrameTransform ? 0D : text.RotationCenterX,
            useFrameTransform ? 0D : text.RotationCenterY,
            strikethrough: (text.Font.Style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough);
    }

    private static void AppendRichText(StringBuilder sb, OfficeDrawingRichText text) {
        bool useFrameTransform = text.FlipHorizontal || text.FlipVertical;
        if (useFrameTransform) {
            AppendRichTextFrameGroupStart(sb, text);
        }

        double contentX = text.X + text.Padding.Left;
        double contentY = text.Y + text.Padding.Top;
        double contentWidth = text.Width - text.Padding.Horizontal;
        double contentHeight = text.Height - text.Padding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            if (useFrameTransform) {
                sb.Append("</g>");
            }

            return;
        }

        OfficeRichTextBlockLayout layout = CreateRichTextLayout(text, contentWidth, contentHeight);
        sb.AppendSvgRichTextBlock(
            layout,
            contentX,
            contentY,
            contentWidth,
            contentHeight,
            text.Alignment,
            text.VerticalAlignment,
            useFrameTransform ? 0D : text.RotationDegrees,
            useFrameTransform ? 0D : text.RotationCenterX,
            useFrameTransform ? 0D : text.RotationCenterY);
        if (useFrameTransform) {
            sb.Append("</g>");
        }
    }

    private static OfficeRichTextBlockLayout CreateRichTextLayout(OfficeDrawingRichText text, double contentWidth, double contentHeight) {
        double maxFontSize = 10D;
        for (int i = 0; i < text.Runs.Count; i++) {
            maxFontSize = Math.Max(maxFontSize, text.Runs[i].FontSize);
        }

        double lineHeightFactor = text.LineHeight.HasValue && text.LineHeight.Value > 0D
            ? Math.Max(1D, text.LineHeight.Value / maxFontSize)
            : 1.2D;
        double minimumFontSize = Math.Min(6D, maxFontSize);
        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create();
        Func<string?, double, string?, double> measure = (value, size, family) => {
                OfficeTextMeasurementStyle measuredStyle = measurer.CreateStyle(new OfficeFontInfo(family, size));
                return measurer.MeasureWidth(value, measuredStyle);
            };
        return OfficeTextLayoutEngine.LayoutRichTextBlock(
            text.Runs,
            contentWidth,
            contentHeight,
            lineHeightFactor,
            measure,
            text.WrapText,
            text.ShrinkToFit,
            minimumFontSize,
            text.ParagraphIndent);
    }

    private static void AppendTextFrameGroupStart(StringBuilder sb, OfficeDrawingText text) {
        string? transform = OfficeSvgFormatting.FormatImageFrameTransform(text.CreateFrameTransform());
        if (string.IsNullOrWhiteSpace(transform)) {
            sb.Append("<g>");
            return;
        }

        sb.Append("<g")
            .AppendAttribute("transform", transform)
            .Append('>');
    }

    private static void AppendRichTextFrameGroupStart(StringBuilder sb, OfficeDrawingRichText text) {
        string? transform = OfficeSvgFormatting.FormatImageFrameTransform(text.CreateFrameTransform());
        if (string.IsNullOrWhiteSpace(transform)) {
            sb.Append("<g>");
            return;
        }

        sb.Append("<g")
            .AppendAttribute("transform", transform)
            .Append('>');
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

    private static string? BuildLineMarkerPaintAttributes(OfficeShape shape) {
        OfficeColor? color = shape.StrokeColor ?? shape.FillColor;
        if (!color.HasValue || color.Value.A == 0) {
            return null;
        }

        var sb = new StringBuilder();
        sb.Append(" fill=\"").Append(ToCssColor(color.Value)).Append('"');
        double opacity = shape.StrokeOpacity ?? ToOpacity(color.Value);
        if (opacity < 1D) {
            sb.Append(" fill-opacity=\"").Append(Format(opacity)).Append('"');
        }

        sb.Append(" stroke=\"none\"");
        return sb.ToString();
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

    private static string BuildGroupTransformAttribute(OfficeDrawingGroup drawingGroup) {
        if (drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform) {
            OfficeTransform transform = OfficeTransform.Translate(drawingGroup.X, drawingGroup.Y)
                .Then(drawingGroup.FrameTransform.Value.CreateDestinationTransform());
            return BuildMatrixTransformAttribute(transform, 0D, 0D);
        }

        return " transform=\"translate(" + Format(drawingGroup.X) + " " + Format(drawingGroup.Y) + ")\"";
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
