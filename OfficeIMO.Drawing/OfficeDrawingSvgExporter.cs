using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Exports dependency-free OfficeIMO drawings to SVG for consumers that need a portable visual fallback.
/// </summary>
public static partial class OfficeDrawingSvgExporter {
    /// <summary>
    /// Converts a drawing to an SVG document.
    /// </summary>
    /// <param name="drawing">Drawing to export.</param>
    /// <returns>SVG markup representing the drawing.</returns>
    public static string ToSvg(OfficeDrawing drawing) {
        return ToSvg(drawing, 1D);
    }

    /// <summary>
    /// Converts a drawing to an SVG document with a scaled output surface.
    /// </summary>
    /// <param name="drawing">Drawing to export.</param>
    /// <param name="scale">Scale applied to the exported SVG width and height.</param>
    /// <returns>SVG markup representing the drawing.</returns>
    public static string ToSvg(OfficeDrawing drawing, double scale) {
        return ToSvg(drawing, scale, OfficeSvgSizeUnit.Point);
    }

    /// <summary>
    /// Converts a drawing to UTF-8 SVG bytes.
    /// </summary>
    /// <param name="drawing">Drawing to export.</param>
    /// <returns>UTF-8 encoded SVG bytes.</returns>
    public static byte[] ToSvgBytes(OfficeDrawing drawing) => Encoding.UTF8.GetBytes(ToSvg(drawing));

    /// <summary>
    /// Converts a drawing to UTF-8 SVG bytes with a scaled output surface.
    /// </summary>
    /// <param name="drawing">Drawing to export.</param>
    /// <param name="scale">Scale applied to the exported SVG width and height.</param>
    /// <returns>UTF-8 encoded SVG bytes.</returns>
    public static byte[] ToSvgBytes(OfficeDrawing drawing, double scale) => Encoding.UTF8.GetBytes(ToSvg(drawing, scale));

    private static void AppendEmbeddedFonts(StringBuilder sb, OfficeFontFaceCollection fonts) {
        if (fonts.Faces.Count == 0) {
            return;
        }

        sb.Append("<defs><style type=\"text/css\">");
        foreach (OfficeFontFace face in fonts.Faces) {
            sb.Append("@font-face{font-family:\"")
                .Append(EscapeCssString(face.FamilyName))
                .Append("\";src:url(data:font/ttf;base64,")
                .Append(Convert.ToBase64String(face.DataSnapshot))
                .Append(") format(\"truetype\");font-weight:")
                .Append((face.Style & OfficeFontStyle.Bold) == OfficeFontStyle.Bold ? "700" : "400")
                .Append(";font-style:")
                .Append((face.Style & OfficeFontStyle.Italic) == OfficeFontStyle.Italic ? "italic" : "normal")
                .Append(";}");
        }

        sb.Append("</style></defs>");
    }

    private static string EscapeCssString(string value) {
        var escaped = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (character == '\\' || character == '"' || character == '<' || character == '>' || character == '&' || char.IsControl(character)) {
                escaped.Append('\\')
                    .Append(((int)character).ToString("X", CultureInfo.InvariantCulture))
                    .Append(' ');
            } else {
                escaped.Append(character);
            }
        }

        return escaped.ToString();
    }

    private static void AppendElements(
        StringBuilder sb,
        IReadOnlyList<OfficeDrawingElement> elements,
        IOfficeRasterImageCodec? imageCodec,
        ref int gradientId,
        ref int clipPathId) {
        for (int i = 0; i < elements.Count; i++) {
            switch (elements[i]) {
                case OfficeDrawingShape drawingShape:
                    string? fillGradientId = null;
                    if (drawingShape.Shape.FillRadialGradient != null) {
                        fillGradientId = "officeimo-gradient-" + (++gradientId).ToString(CultureInfo.InvariantCulture);
                        sb.AppendRadialGradientDefinition(fillGradientId, drawingShape.Shape.FillRadialGradient);
                    } else if (drawingShape.Shape.FillGradient != null) {
                        fillGradientId = "officeimo-gradient-" + (++gradientId).ToString(CultureInfo.InvariantCulture);
                        sb.AppendLinearGradientDefinition(fillGradientId, drawingShape.Shape.FillGradient);
                    }

                    string? strokeGradientId = null;
                    if (drawingShape.Shape.StrokeRadialGradient != null) {
                        strokeGradientId = "officeimo-gradient-" + (++gradientId).ToString(CultureInfo.InvariantCulture);
                        sb.AppendRadialGradientDefinition(strokeGradientId, drawingShape.Shape.StrokeRadialGradient);
                    } else if (drawingShape.Shape.StrokeGradient != null) {
                        strokeGradientId = "officeimo-gradient-" + (++gradientId).ToString(CultureInfo.InvariantCulture);
                        sb.AppendLinearGradientDefinition(strokeGradientId, drawingShape.Shape.StrokeGradient);
                    }

                    string? shapeClipPathId = null;
                    if (drawingShape.Shape.ClipPath != null) {
                        shapeClipPathId = "officeimo-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
                        AppendClipPathDefinition(sb, shapeClipPathId, drawingShape.Shape.ClipPath);
                    }

                    AppendShape(sb, drawingShape, fillGradientId, strokeGradientId, shapeClipPathId);
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
                    AppendImage(sb, drawingImage, imageClipPathId, imageCodec);
                    break;
                case OfficeDrawingImagePattern imagePattern:
                    AppendImagePattern(sb, imagePattern, imageCodec, ref clipPathId);
                    break;
                case OfficeDrawingTilingPattern tilingPattern:
                    AppendTilingPattern(sb, tilingPattern, imageCodec, ref gradientId, ref clipPathId);
                    break;
                case OfficeDrawingGroup drawingGroup:
                    AppendGroup(sb, drawingGroup, imageCodec, ref gradientId, ref clipPathId);
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    AppendEffectGroup(sb, effectGroup, imageCodec, ref gradientId, ref clipPathId);
                    break;
            }
        }
    }

    private static void AppendGroup(StringBuilder sb, OfficeDrawingGroup drawingGroup, IOfficeRasterImageCodec? imageCodec, ref int gradientId, ref int clipPathId) {
        string groupClipPathId = "officeimo-group-clip-" + (++clipPathId).ToString(CultureInfo.InvariantCulture);
        AppendClipPathDefinition(sb, groupClipPathId, drawingGroup.ClipPath);
        string transform = BuildGroupTransformAttribute(drawingGroup);
        sb.Append("<g")
            .AppendClipPathReference(groupClipPathId)
            .Append(transform)
            .Append('>');
        bool hasContentOffset = Math.Abs(drawingGroup.ContentOffsetX) > 0.0000001D || Math.Abs(drawingGroup.ContentOffsetY) > 0.0000001D;
        if (hasContentOffset) {
            sb.Append("<g transform=\"translate(")
                .Append(Format(drawingGroup.ContentOffsetX))
                .Append(' ')
                .Append(Format(drawingGroup.ContentOffsetY))
                .Append(")\">");
        }
        AppendElements(sb, drawingGroup.InnerDrawing.Elements, imageCodec, ref gradientId, ref clipPathId);
        if (hasContentOffset) sb.Append("</g>");
        sb.Append("</g>");
    }

    private static void AppendShape(StringBuilder sb, OfficeDrawingShape drawingShape, string? fillGradientId, string? strokeGradientId, string? clipPathId) {
        IReadOnlyList<OfficeDrawingShape> glowShapes = CreateGlowShapes(drawingShape);
        for (int i = 0; i < glowShapes.Count; i++) {
            AppendShapeGeometry(sb, glowShapes[i], fillGradientId: null, strokeGradientId: null, clipPathId);
        }

        IReadOnlyList<OfficeDrawingShape> shadowShapes = CreateShadowShapes(drawingShape);
        for (int i = 0; i < shadowShapes.Count; i++) {
            AppendShapeGeometry(sb, shadowShapes[i], fillGradientId: null, strokeGradientId: null, clipPathId);
        }

        AppendShapeGeometry(sb, drawingShape, fillGradientId, strokeGradientId, clipPathId);
    }

    private static void AppendShapeGeometry(StringBuilder sb, OfficeDrawingShape drawingShape, string? fillGradientId, string? strokeGradientId, string? clipPathId) {
        OfficeShape shape = drawingShape.Shape;
        string paint = BuildPaintAttributes(shape, fillGradientId, strokeGradientId);
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
                AppendLine(sb, drawingShape, paint, transform, originX, originY, strokeGradientId);
                break;
            case OfficeShapeKind.Polygon:
                AppendPolygon(sb, drawingShape, paint, transform, originX, originY);
                break;
            case OfficeShapeKind.Path:
                AppendPath(sb, drawingShape, paint, transform, originX, originY, strokeGradientId);
                break;
        }

        if (clipPathId != null) {
            sb.Append("</g>");
        }
    }

    private static IReadOnlyList<OfficeDrawingShape> CreateGlowShapes(OfficeDrawingShape drawingShape) {
        OfficeShape shape = drawingShape.Shape;
        OfficeGlow? glow = shape.Glow;
        if (glow == null || glow.Radius <= 0D || glow.Opacity <= 0D || glow.Color.A == 0) {
            return Array.Empty<OfficeDrawingShape>();
        }

        const int layers = 4;
        var glowShapes = new List<OfficeDrawingShape>(layers);
        double baseStrokeWidth = Math.Max(0D, shape.StrokeWidth);
        for (int i = layers; i >= 1; i--) {
            double factor = i / (double)layers;
            OfficeShape glowShape = shape.Clone();
            glowShape.Shadow = null;
            glowShape.Glow = null;
            glowShape.FillColor = null;
            glowShape.FillGradient = null;
            glowShape.FillRadialGradient = null;
            glowShape.StrokeColor = glow.Color;
            glowShape.StrokeWidth = Math.Max(1D, baseStrokeWidth + glow.Radius * 2D * factor);
            glowShape.StrokeDashStyle = OfficeStrokeDashStyle.Solid;
            glowShape.StrokeStartMarker = null;
            glowShape.StrokeEndMarker = null;
            glowShape.StrokeOpacity = ComputeGlowLayerOpacity(glow.Opacity, layers - i + 1);
            glowShapes.Add(new OfficeDrawingShape(glowShape, drawingShape.X, drawingShape.Y));
        }

        return glowShapes;
    }

    private static double ComputeGlowLayerOpacity(double opacity, int layerDepth) {
        double clamped = opacity < 0D ? 0D : opacity > 1D ? 1D : opacity;
        return 1D - Math.Pow(1D - clamped, layerDepth + 1);
    }

    private static IReadOnlyList<OfficeDrawingShape> CreateShadowShapes(OfficeDrawingShape drawingShape) {
        OfficeShape shape = drawingShape.Shape;
        OfficeShadow? shadow = shape.Shadow;
        if (shadow == null || shadow.Opacity <= 0D || shadow.Color.A == 0) {
            return Array.Empty<OfficeDrawingShape>();
        }

        bool hasStroke = shape.Kind == OfficeShapeKind.Line ||
            (shape.StrokeWidth > 0D &&
                (shape.StrokeRadialGradient != null ||
                 shape.StrokeGradient != null ||
                 (shape.StrokeColor.HasValue && shape.StrokeColor.Value.A > 0)));
        bool hasFill = shape.Kind != OfficeShapeKind.Line &&
            (shape.FillRadialGradient != null || shape.FillGradient != null || (shape.FillColor.HasValue && shape.FillColor.Value.A > 0));
        OfficeDrawingShape coreShadow = CreateShadowShape(drawingShape, shadow, hasStroke, hasFill, Math.Max(0D, shape.StrokeWidth), shadow.Opacity);
        if (shadow.BlurRadius <= 0D) {
            return new[] { coreShadow };
        }

        const int layers = 4;
        var shadowShapes = new List<OfficeDrawingShape>(layers + 1);
        double baseStrokeWidth = Math.Max(0D, shape.StrokeWidth);
        for (int i = layers; i >= 1; i--) {
            double factor = i / (double)layers;
            double opacity = shadow.Opacity * (0.04D + (layers - i + 1) * 0.05D);
            shadowShapes.Add(CreateShadowShape(
                drawingShape,
                shadow,
                hasStroke: true,
                hasFill: hasFill,
                strokeWidth: Math.Max(1D, baseStrokeWidth + shadow.BlurRadius * 2D * factor),
                opacity: opacity));
        }

        shadowShapes.Add(coreShadow);
        return shadowShapes;
    }

    private static OfficeDrawingShape CreateShadowShape(OfficeDrawingShape drawingShape, OfficeShadow shadow, bool hasStroke, bool hasFill, double strokeWidth, double opacity) {
        OfficeShape shape = drawingShape.Shape;
        var shadowShape = shape.Clone();
        shadowShape.Shadow = null;
        shadowShape.Glow = null;
        shadowShape.FillGradient = null;
        shadowShape.FillRadialGradient = null;
        shadowShape.FillColor = hasFill || !hasStroke ? shadow.Color : null;
        shadowShape.FillOpacity = opacity;
        shadowShape.StrokeColor = hasStroke ? shadow.Color : null;
        shadowShape.StrokeGradient = null;
        shadowShape.StrokeRadialGradient = null;
        shadowShape.StrokeWidth = strokeWidth;
        shadowShape.StrokeDashStyle = OfficeStrokeDashStyle.Solid;
        shadowShape.StrokeStartMarker = null;
        shadowShape.StrokeEndMarker = null;
        shadowShape.StrokeOpacity = opacity;

        return CreateOffsetEffectShape(shadowShape, drawingShape.X + shadow.OffsetX, drawingShape.Y + shadow.OffsetY);
    }

    private static OfficeDrawingShape CreateOffsetEffectShape(OfficeShape shape, double x, double y) {
        double clampedX = Math.Max(0D, x);
        double clampedY = Math.Max(0D, y);
        double offsetX = x - clampedX;
        double offsetY = y - clampedY;
        if (offsetX != 0D || offsetY != 0D) {
            shape = shape.Clone();
            OfficeTransform offsetTransform = OfficeTransform.Translate(offsetX, offsetY);
            shape.Transform = shape.Transform.HasValue ? offsetTransform.Then(shape.Transform.Value) : offsetTransform;
        }

        return new OfficeDrawingShape(shape, clampedX, clampedY);
    }

    private static void AppendImage(StringBuilder sb, OfficeDrawingImage drawingImage, string? clipPathId, IOfficeRasterImageCodec? imageCodec) {
        byte[] bytes = drawingImage.EncodedBytes;
        if (!OfficeSvgImageRenderer.TryCreateDataUri(drawingImage.ContentType, bytes, null, imageCodec, out string dataUri)) {
            return;
        }

        if (drawingImage.Opacity < 1D) {
            sb.Append("<g")
                .AppendAttribute("opacity", OfficeSvgFormatting.FormatNumber(drawingImage.Opacity))
                .Append('>');
        }

        OfficeSvgImageRenderer.AppendImage(
            sb,
            dataUri,
            drawingImage.Projection,
            clipPathId,
            drawingImage.Projection.HasCrop ? drawingImage.Projection.Placement : null,
            "none");
        if (drawingImage.Opacity < 1D) {
            sb.Append("</g>");
        }
    }

    private static void AppendLine(StringBuilder sb, OfficeDrawingShape drawingShape, string paint, string transform, double originX, double originY, string? strokeGradientId) {
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
        AppendLineMarker(sb, shape.StrokeStartMarker, start, new OfficePoint(start.X - end.X, start.Y - end.Y), shape, transform, strokeGradientId);
        AppendLineMarker(sb, shape.StrokeEndMarker, end, new OfficePoint(end.X - start.X, end.Y - start.Y), shape, transform, strokeGradientId);
    }

    private static void AppendLineMarker(StringBuilder sb, OfficeLineMarker? marker, OfficePoint tip, OfficePoint lineDirection, OfficeShape shape, string transform, string? strokeGradientId) {
        IReadOnlyList<OfficePoint> contour = OfficeLineMarkerGeometry.CreateContour(marker, tip, lineDirection);
        if (contour.Count == 0) {
            return;
        }

        string? paint = BuildLineMarkerPaintAttributes(shape, strokeGradientId);
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

    private static void AppendPath(StringBuilder sb, OfficeDrawingShape drawingShape, string paint, string transform, double originX, double originY, string? strokeGradientId) {
        OfficeShape shape = drawingShape.Shape;
        if (shape.PathCommands.Count == 0) {
            return;
        }

        sb.AppendPathElement(shape.PathCommands, originX, originY, paint + BuildFillRuleAttribute(shape.FillRule) + transform);
        AppendPathMarkers(sb, shape, originX, originY, transform, strokeGradientId);
    }

    private static string BuildFillRuleAttribute(OfficeFillRule fillRule) =>
        fillRule == OfficeFillRule.EvenOdd ? " fill-rule=\"evenodd\"" : string.Empty;

    private static string BuildClipRuleAttribute(OfficeFillRule fillRule) =>
        fillRule == OfficeFillRule.EvenOdd ? " clip-rule=\"evenodd\"" : string.Empty;

    private static void AppendPathMarkers(StringBuilder sb, OfficeShape shape, double originX, double originY, string transform, string? strokeGradientId) {
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
            AppendLineMarker(sb, shape.StrokeStartMarker, start, new OfficePoint(start.X - next.X, start.Y - next.Y), shape, transform, strokeGradientId);
        }

        if (lastOpen != null) {
            IReadOnlyList<OfficePoint> points = lastOpen.Points;
            OfficePoint end = points[points.Count - 1];
            OfficePoint previous = points[points.Count - 2];
            AppendLineMarker(sb, shape.StrokeEndMarker, end, new OfficePoint(end.X - previous.X, end.Y - previous.Y), shape, transform, strokeGradientId);
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
        if (text.TextAdvanceWidth.HasValue) {
            sb.AppendSvgPositionedTextElement(
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
                (text.Font.Style & OfficeFontStyle.Strikethrough) == OfficeFontStyle.Strikethrough,
                text.TextAdvanceWidth.Value);
        } else {
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
        }

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
        sb.AppendPathElement(clipPath.Commands, attributes: BuildClipRuleAttribute(clipPath.FillRule));
    }

    private static string BuildPaintAttributes(OfficeShape shape, string? fillGradientId, string? strokeGradientId) {
        var sb = new StringBuilder();
        if (fillGradientId != null) {
            sb.Append(" fill=\"url(#").Append(Escape(fillGradientId)).Append(")\"");
            double fillOpacity = shape.FillOpacity ?? 1D;
            if (fillOpacity < 1D) {
                sb.Append(" fill-opacity=\"").Append(Format(fillOpacity)).Append('"');
            }
        } else if (shape.FillColor.HasValue && shape.FillColor.Value.A > 0) {
            sb.Append(" fill=\"").Append(ToCssColor(shape.FillColor.Value)).Append('"');
            double fillOpacity = (shape.FillOpacity ?? 1D) * ToOpacity(shape.FillColor.Value);
            if (fillOpacity < 1D) {
                sb.Append(" fill-opacity=\"").Append(Format(fillOpacity)).Append('"');
            }
        } else {
            sb.Append(" fill=\"none\"");
        }

        if (strokeGradientId != null && shape.StrokeWidth > 0) {
            sb.Append(" stroke=\"url(#").Append(Escape(strokeGradientId)).Append(")\"")
                .Append(" stroke-width=\"").Append(Format(shape.StrokeWidth)).Append('"');
            double strokeOpacity = shape.StrokeOpacity ?? 1D;
            if (strokeOpacity < 1D) {
                sb.Append(" stroke-opacity=\"").Append(Format(strokeOpacity)).Append('"');
            }

            AppendStrokeStyle(sb, shape);
        } else if (shape.StrokeColor.HasValue && shape.StrokeWidth > 0 && shape.StrokeColor.Value.A > 0) {
            sb.Append(" stroke=\"").Append(ToCssColor(shape.StrokeColor.Value)).Append('"')
                .Append(" stroke-width=\"").Append(Format(shape.StrokeWidth)).Append('"');
            double strokeOpacity = (shape.StrokeOpacity ?? 1D) * ToOpacity(shape.StrokeColor.Value);
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

    private static string? BuildLineMarkerPaintAttributes(OfficeShape shape, string? strokeGradientId) {
        if (strokeGradientId != null) {
            var gradientPaint = new StringBuilder();
            gradientPaint.Append(" fill=\"url(#").Append(Escape(strokeGradientId)).Append(")\"");
            double gradientOpacity = shape.StrokeOpacity ?? 1D;
            if (gradientOpacity < 1D) {
                gradientPaint.Append(" fill-opacity=\"").Append(Format(gradientOpacity)).Append('"');
            }

            gradientPaint.Append(" stroke=\"none\"");
            return gradientPaint.ToString();
        }

        OfficeColor? color = shape.StrokeColor ?? shape.FillColor;
        if (!color.HasValue || color.Value.A == 0) {
            return null;
        }

        var sb = new StringBuilder();
        sb.Append(" fill=\"").Append(ToCssColor(color.Value)).Append('"');
        double opacity = (shape.StrokeOpacity ?? 1D) * ToOpacity(color.Value);
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
