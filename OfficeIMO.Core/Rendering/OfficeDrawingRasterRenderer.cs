using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free raster renderer for <see cref="OfficeDrawing"/> scenes.
/// </summary>
public static partial class OfficeDrawingRasterRenderer {
    /// <summary>
    /// Renders a drawing to an RGBA raster image.
    /// </summary>
    public static OfficeRasterImage Render(OfficeDrawing drawing, double scale = 1D, OfficeColor? background = null) {
        return Render(drawing, new OfficeDrawingRasterRenderOptions { Scale = scale, Background = background });
    }

    /// <summary>Renders a drawing with an optional external image codec.</summary>
    public static OfficeRasterImage Render(OfficeDrawing drawing, OfficeDrawingRasterRenderOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        return RenderCore(drawing, options, options.Scale, options.Scale);
    }

    private static OfficeRasterImage RenderCore(OfficeDrawing drawing, OfficeDrawingRasterRenderOptions options,
        double scaleX, double scaleY, OfficeRasterExportPlan? allocationPlan = null, OfficeFontFaceCollection? fonts = null) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        if (options == null) throw new ArgumentNullException(nameof(options));
        options.CancellationToken.ThrowIfCancellationRequested();
        double scale = options.Scale;

        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
            throw new ArgumentOutOfRangeException(nameof(scale), "Scale must be a finite positive number.");
        }
        if (options.MaximumRasterPixels <= 0L) {
            throw new ArgumentOutOfRangeException(nameof(options.MaximumRasterPixels), "Maximum raster pixels must be positive.");
        }

        bool uniform = scaleX == scale && scaleY == scale;
        OfficeRasterExportPlan plan = allocationPlan ?? OfficeRasterExportPlanner.Resolve(
            uniform ? drawing.Width : drawing.Width * scaleX,
            uniform ? drawing.Height : drawing.Height * scaleY,
            OfficeImageExportFormat.Png,
            new OfficeImageExportOptions {
                Scale = uniform ? scale : 1D,
                MaximumRasterPixels = options.MaximumRasterPixels,
                RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw
            });

        int width = plan.Limit.PixelWidth;
        int height = plan.Limit.PixelHeight;
        OfficeRasterImage image = new OfficeRasterImage(width, height, options.Background);
        OfficeRasterCanvas canvas = new OfficeRasterCanvas(
            image,
            font: null,
            fonts: fonts ?? drawing.Fonts,
            textShapingProvider: options.TextShapingProvider ?? drawing.TextShapingProvider,
            textShapingLanguage: options.TextShapingLanguage ?? drawing.TextShapingLanguage,
            diagnosticSink: options.DiagnosticSink,
            diagnosticSource: options.DiagnosticSource,
            cancellationToken: options.CancellationToken);
        canvas.FontMetricScale = scale;
        canvas.SetCoordinateScale(scaleX / scale, scaleY / scale);
        if (options.TransformedTextBudget != null) canvas.ShareTransformedTextBudget(options.TransformedTextBudget);
        IOfficeRasterImageCodec? imageCodec = options.ThrowOnImageDecodeFailure
            ? new RequiredImageCodec(options.ImageCodec, options.MaximumRasterPixels, options.CancellationToken)
            : options.ImageCodec;
        RenderElements(canvas, drawing.Elements, scale, imageCodec, options.MaximumRasterPixels, options.CancellationToken);

        return image;
    }

    private static void RenderElements(
        OfficeRasterCanvas canvas,
        IEnumerable<OfficeDrawingElement> elements,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        long maximumRasterPixels,
        System.Threading.CancellationToken cancellationToken) {
        foreach (OfficeDrawingElement element in elements) {
            cancellationToken.ThrowIfCancellationRequested();
            if (element is OfficeDrawingShape shape) {
                RenderShape(canvas, shape, scale);
            } else if (element is OfficeDrawingText text) {
                bool preserve = canvas.PreservePaintedGlyphOrder;
                canvas.PreservePaintedGlyphOrder = text.PreservesPaintedGlyphs;
                try {
                    RenderText(canvas, text, scale, maximumRasterPixels);
                } finally {
                    canvas.PreservePaintedGlyphOrder = preserve;
                }
            } else if (element is OfficeDrawingRichText richText) {
                RenderRichText(canvas, richText, scale);
            } else if (element is OfficeDrawingImage drawingImage) {
                RenderImage(canvas, drawingImage, scale, imageCodec, maximumRasterPixels, cancellationToken);
            } else if (element is OfficeDrawingImagePattern imagePattern) {
                RenderImagePattern(canvas, imagePattern, scale, imageCodec, maximumRasterPixels, cancellationToken);
            } else if (element is OfficeDrawingTilingPattern tilingPattern) {
                RenderTilingPattern(canvas, tilingPattern, scale, imageCodec, maximumRasterPixels, cancellationToken);
            } else if (element is OfficeDrawingGroup drawingGroup) {
                RenderGroup(canvas, drawingGroup, scale, imageCodec, maximumRasterPixels, cancellationToken);
            } else if (element is OfficeDrawingEffectGroup effectGroup) {
                RenderEffectGroup(canvas, effectGroup, scale, imageCodec, maximumRasterPixels, cancellationToken);
            }
        }
    }

    /// <summary>
    /// Renders a drawing to PNG bytes.
    /// </summary>
    public static byte[] ToPng(OfficeDrawing drawing, double scale = 1D, OfficeColor? background = null) =>
        OfficePngWriter.Encode(Render(drawing, scale, background));

    /// <summary>Renders a drawing to PNG bytes with an optional external image codec.</summary>
    public static byte[] ToPng(OfficeDrawing drawing, OfficeDrawingRasterRenderOptions options) => OfficePngWriter.Encode(Render(drawing, options));

    private static void RenderGroup(
        OfficeRasterCanvas canvas,
        OfficeDrawingGroup drawingGroup,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        long maximumRasterPixels,
        System.Threading.CancellationToken cancellationToken) {
        if (drawingGroup.ClipPath.Kind == OfficeClipPathKind.Empty) return;
        canvas = canvas.WithDrawingTextProfile(drawingGroup.InnerDrawing);
        using (PushGroupClip(canvas, drawingGroup, scale)) {
            var translated = new OfficeDrawing(
                Math.Max(1D, canvas.Width / scale),
                Math.Max(1D, canvas.Height / scale));
            double contentX = drawingGroup.X + drawingGroup.ContentOffsetX;
            double contentY = drawingGroup.Y + drawingGroup.ContentOffsetY;
            if (drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform) {
                translated.AddDrawingForClippedRendering(drawingGroup.InnerDrawing, contentX, contentY, drawingGroup.FrameTransform.Value);
            } else {
                translated.AddDrawingForClippedRendering(drawingGroup.InnerDrawing, contentX, contentY, null);
            }

            RenderElements(canvas, translated.Elements, scale, imageCodec, maximumRasterPixels, cancellationToken);
        }
    }

    private static IDisposable PushGroupClip(OfficeRasterCanvas canvas, OfficeDrawingGroup drawingGroup, double scale) {
        IReadOnlyList<IReadOnlyList<OfficePoint>> contours = CreateGroupClipContours(drawingGroup, scale);
        if (contours.Count > 0) {
            return contours.Count == 1 && drawingGroup.ClipPath.Kind != OfficeClipPathKind.Path
                ? PushSingleContourClip(canvas, drawingGroup.ClipPath.Kind, contours[0])
                : PushClipPolygons(canvas, contours, drawingGroup.ClipPath.FillRule);
        }

        if (drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform) {
            OfficeTransform transform = drawingGroup.FrameTransform.Value.CreateDestinationTransform();
            return canvas.PushClipPolygon(new[] {
                ScalePoint(transform.TransformPoint(new OfficePoint(drawingGroup.X, drawingGroup.Y)), scale),
                ScalePoint(transform.TransformPoint(new OfficePoint(drawingGroup.X + drawingGroup.ClipPath.Width, drawingGroup.Y)), scale),
                ScalePoint(transform.TransformPoint(new OfficePoint(drawingGroup.X + drawingGroup.ClipPath.Width, drawingGroup.Y + drawingGroup.ClipPath.Height)), scale),
                ScalePoint(transform.TransformPoint(new OfficePoint(drawingGroup.X, drawingGroup.Y + drawingGroup.ClipPath.Height)), scale)
            });
        }

        return canvas.PushClipRectangle(
            drawingGroup.X * scale,
            drawingGroup.Y * scale,
            drawingGroup.ClipPath.Width * scale,
            drawingGroup.ClipPath.Height * scale);
    }

    private static OfficePoint ScalePoint(OfficePoint point, double scale) =>
        new OfficePoint(point.X * scale, point.Y * scale);

    private static void RenderShape(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        if (drawingShape.Shape.ClipPath?.Kind == OfficeClipPathKind.Empty) return;
        IReadOnlyList<OfficeDrawingShape> glowShapes = CreateGlowShapes(drawingShape);
        for (int i = 0; i < glowShapes.Count; i++) {
            RenderShape(canvas, glowShapes[i], scale);
        }

        IReadOnlyList<OfficeDrawingShape> shadowShapes = CreateShadowShapes(drawingShape);
        for (int i = 0; i < shadowShapes.Count; i++) {
            RenderShape(canvas, shadowShapes[i], scale);
        }

        IDisposable? clipScope = PushShapeClip(canvas, drawingShape, scale);
        try {
            RenderShapeGeometry(canvas, drawingShape, scale);
        } finally {
            clipScope?.Dispose();
        }
    }

    private static void RenderShapeGeometry(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        OfficeShape shape = drawingShape.Shape;
        if (HasNonIdentityTransform(shape.Transform)) {
            RenderTransformedShape(canvas, drawingShape, scale);
            return;
        }

        double x = drawingShape.X * scale;
        double y = drawingShape.Y * scale;
        double width = shape.Width * scale;
        double height = shape.Height * scale;
        OfficeColor? fill = ApplyOpacity(shape.FillColor, shape.FillOpacity);
        OfficeColor? stroke = ApplyOpacity(shape.StrokeColor, shape.StrokeOpacity);
        OfficeRadialGradient? radialGradient = shape.FillRadialGradient == null ? null : ApplyOpacity(shape.FillRadialGradient, shape.FillOpacity);
        OfficeLinearGradient? linearGradient = shape.FillGradient == null ? null : ApplyOpacity(shape.FillGradient, shape.FillOpacity);
        OfficeRadialGradient? strokeRadialGradient = shape.StrokeRadialGradient == null ? null : ApplyOpacity(shape.StrokeRadialGradient, shape.StrokeOpacity);
        OfficeLinearGradient? strokeLinearGradient = shape.StrokeGradient == null ? null : ApplyOpacity(shape.StrokeGradient, shape.StrokeOpacity);
        double strokeWidth = shape.StrokeWidth * scale;

        switch (shape.Kind) {
            case OfficeShapeKind.Rectangle:
                if (radialGradient != null) {
                    canvas.FillRadialGradientRectangle(x, y, width, height, radialGradient);
                } else if (linearGradient != null) {
                    canvas.FillLinearGradientRectangle(x, y, width, height, linearGradient);
                } else if (fill.HasValue) {
                    canvas.FillRectangle(x, y, width, height, fill.Value);
                }

                if ((stroke.HasValue || strokeLinearGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
                    if (shape.StrokeDashStyle == OfficeStrokeDashStyle.Solid) {
                        DrawGradientOrSolidPolyline(canvas, CreateRectangleContour(x, y, width, height), stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape, close: true);
                    } else {
                        DrawGradientOrSolidPolyline(canvas, CreateRectangleContour(x, y, width, height), stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape, close: true);
                    }
                }

                break;
            case OfficeShapeKind.RoundedRectangle:
                IReadOnlyList<OfficePoint> rounded = OffsetPoints(CreateRoundedRectangleContour(width, height, shape.CornerRadius * scale, 1D), x, y, 1D);
                if (radialGradient != null) canvas.FillRadialGradientPolygon(rounded, radialGradient);
                else if (linearGradient != null) canvas.FillLinearGradientPolygon(rounded, linearGradient);
                else if (fill.HasValue) canvas.FillPolygon(rounded, fill.Value);
                DrawGradientOrSolidPolyline(canvas, rounded, stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape, close: true);
                break;
            case OfficeShapeKind.Ellipse:
                IReadOnlyList<OfficePoint> ellipseFill = OffsetPoints(CreateEllipseContour(width, height, 1D), x, y, 1D);
                if (radialGradient != null) canvas.FillRadialGradientPolygon(ellipseFill, radialGradient);
                else if (linearGradient != null) canvas.FillLinearGradientPolygon(ellipseFill, linearGradient);
                else if (fill.HasValue) canvas.FillPolygon(ellipseFill, fill.Value);
                if ((stroke.HasValue || strokeLinearGradient != null || strokeRadialGradient != null) && strokeWidth > 0D) {
                    if (shape.StrokeDashStyle == OfficeStrokeDashStyle.Solid) {
                        DrawGradientOrSolidPolyline(canvas, OffsetPoints(CreateEllipseContour(width, height, 1D), x, y, 1D), stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape, close: true);
                    } else {
                        DrawGradientOrSolidPolyline(canvas, OffsetPoints(CreateEllipseContour(width, height, 1D), x, y, 1D), stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth, shape, close: true);
                    }
                }

                break;
            case OfficeShapeKind.Line:
                if (strokeWidth > 0D && (stroke.HasValue || strokeLinearGradient != null || strokeRadialGradient != null))
                    RenderLine(canvas, shape, x, y, scale, stroke ?? OfficeColor.Transparent, strokeLinearGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Polygon:
                RenderPolygon(canvas, shape, x, y, scale, fill, linearGradient, radialGradient, stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth);
                break;
            case OfficeShapeKind.Path:
                RenderPath(canvas, shape, x, y, scale, fill, linearGradient, radialGradient, stroke, strokeLinearGradient, strokeRadialGradient, strokeWidth);
                break;
        }
    }

    internal static void RenderRichText(OfficeRasterCanvas canvas, OfficeDrawingRichText text, double scale) {
        OfficeTextPadding scaledPadding = text.Padding.Scale(scale);
        double contentX = (text.X * scale) + scaledPadding.Left;
        double contentY = (text.Y * scale) + scaledPadding.Top;
        double contentWidth = (text.Width * scale) - scaledPadding.Horizontal;
        double contentHeight = (text.Height * scale) - scaledPadding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            return;
        }

        OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(
            text, contentWidth, contentHeight, canvas, scale);
        OfficeTextBlockRenderer.DrawRasterRichTextBlock(
            canvas,
            layout,
            contentX,
            contentY,
            contentWidth,
            contentHeight,
            text.Alignment,
            text.VerticalAlignment,
            text.RotationDegrees,
            text.RotationCenterX * scale,
            text.RotationCenterY * scale,
            flipHorizontal: text.FlipHorizontal,
            flipVertical: text.FlipVertical);
    }

    private static IDisposable? PushShapeClip(OfficeRasterCanvas canvas, OfficeDrawingShape drawingShape, double scale) {
        OfficeClipPath? clipPath = drawingShape.Shape.ClipPath;
        if (clipPath == null) {
            return null;
        }

        IReadOnlyList<IReadOnlyList<OfficePoint>> contours = CreateClipContours(drawingShape, clipPath, scale);
        if (contours.Count == 0) {
            return null;
        }

        return contours.Count == 1 && clipPath.Kind != OfficeClipPathKind.Path
            ? PushSingleContourClip(canvas, clipPath.Kind, contours[0])
            : PushClipPolygons(canvas, contours, clipPath.FillRule);
    }

    private static IReadOnlyList<IReadOnlyList<OfficePoint>> CreateClipContours(OfficeDrawingShape drawingShape, OfficeClipPath clipPath, double scale) {
        return OfficeClipPathGeometry.CreateContours(clipPath, contour => TransformClipContour(drawingShape, contour, scale));
    }

    private static IReadOnlyList<IReadOnlyList<OfficePoint>> CreateGroupClipContours(OfficeDrawingGroup drawingGroup, double scale) {
        OfficeTransform? transform = drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform
            ? drawingGroup.FrameTransform.Value.CreateDestinationTransform()
            : null;
        return OfficeClipPathGeometry.CreateContours(
            drawingGroup.ClipPath,
            contour => TransformGroupClipContour(drawingGroup, contour, scale, transform));
    }

    private static void FillPathContours(OfficeRasterCanvas canvas, IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeColor color, OfficeFillRule fillRule) {
        if (fillRule == OfficeFillRule.NonZero) {
            canvas.FillPolygonsNonZero(contours, color);
        } else {
            canvas.FillPolygonsEvenOdd(contours, color);
        }
    }

    private static IDisposable PushClipPolygons(OfficeRasterCanvas canvas, IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule fillRule) =>
        fillRule == OfficeFillRule.NonZero
            ? canvas.PushClipPolygonsNonZero(contours)
            : canvas.PushClipPolygonsEvenOdd(contours);

    private static IReadOnlyList<OfficePoint> TransformClipContour(OfficeDrawingShape drawingShape, IReadOnlyList<OfficePoint> contour, double scale) =>
        HasNonIdentityTransform(drawingShape.Shape.Transform)
            ? TransformShapePoints(drawingShape, contour, scale)
            : OffsetPoints(contour, drawingShape.X * scale, drawingShape.Y * scale, scale);

    private static IReadOnlyList<OfficePoint> TransformGroupClipContour(OfficeDrawingGroup drawingGroup, IReadOnlyList<OfficePoint> contour, double scale, OfficeTransform? transform) {
        List<OfficePoint> points = new List<OfficePoint>(contour.Count);
        for (int i = 0; i < contour.Count; i++) {
            OfficePoint point = new OfficePoint(drawingGroup.X + contour[i].X, drawingGroup.Y + contour[i].Y);
            if (transform.HasValue) {
                point = transform.Value.TransformPoint(point);
            }

            points.Add(ScalePoint(point, scale));
        }

        return points;
    }

    private static List<OfficePoint> OffsetPoints(IReadOnlyList<OfficePoint> source, double x, double y, double scale) {
        List<OfficePoint> points = new List<OfficePoint>(source.Count);
        for (int i = 0; i < source.Count; i++) {
            points.Add(new OfficePoint(x + (source[i].X * scale), y + (source[i].Y * scale)));
        }

        return points;
    }

    private static IReadOnlyList<OfficePoint> CreateShapeContour(OfficeShape shape, double pixelsPerUnit = 1D) {
        switch (shape.Kind) {
            case OfficeShapeKind.Ellipse:
                return CreateEllipseContour(shape.Width, shape.Height, pixelsPerUnit);
            case OfficeShapeKind.RoundedRectangle:
                return CreateRoundedRectangleContour(shape.Width, shape.Height, shape.CornerRadius, pixelsPerUnit);
            case OfficeShapeKind.Rectangle:
            default:
                return new[] {
                    new OfficePoint(0D, 0D),
                    new OfficePoint(shape.Width, 0D),
                    new OfficePoint(shape.Width, shape.Height),
                    new OfficePoint(0D, shape.Height)
                };
        }
    }

    private static IReadOnlyList<OfficePoint> CreateEllipseContour(double width, double height, double pixelsPerUnit) {
        return OfficeCurveFlattening.Ellipse(width / 2D, height / 2D, width / 2D, height / 2D, pixelsPerUnit);
    }

    private static IReadOnlyList<OfficePoint> CreateRoundedRectangleContour(double width, double height, double radius, double pixelsPerUnit) =>
        OfficeCurveFlattening.RoundedRectangle(0D, 0D, width, height, radius, radius, pixelsPerUnit);

    private static double GetShapePixelScale(OfficeDrawingShape shape, double scale) {
        OfficeTransform transform = shape.Shape.Transform ?? OfficeTransform.Identity;
        return scale * Math.Sqrt(transform.M11 * transform.M11 + transform.M12 * transform.M12 + transform.M21 * transform.M21 + transform.M22 * transform.M22);
    }
    private static List<OfficePoint> TransformShapePoints(OfficeDrawingShape drawingShape, IReadOnlyList<OfficePoint> points, double scale) {
        List<OfficePoint> transformed = new List<OfficePoint>(points.Count);
        for (int i = 0; i < points.Count; i++) {
            transformed.Add(TransformShapePoint(drawingShape, points[i], scale));
        }

        return transformed;
    }

    private static OfficePoint TransformShapePoint(OfficeDrawingShape drawingShape, OfficePoint point, double scale) {
        OfficePoint local = drawingShape.Shape.Transform.HasValue
            ? drawingShape.Shape.Transform.Value.TransformPoint(point)
            : point;
        return new OfficePoint((drawingShape.X + local.X) * scale, (drawingShape.Y + local.Y) * scale);
    }

    private static bool HasNonIdentityTransform(OfficeTransform? transform) =>
        transform.HasValue && transform.Value != OfficeTransform.Identity;

    private static OfficeColor? ApplyOpacity(OfficeColor? color, double? opacity) {
        if (!color.HasValue) return null;
        if (!opacity.HasValue) return color;
        double clamped = opacity.Value < 0D ? 0D : opacity.Value > 1D ? 1D : opacity.Value;
        return OfficeColor.FromRgba(color.Value.R, color.Value.G, color.Value.B, (byte)Math.Round(color.Value.A * clamped));
    }

    private static OfficeColor? ApplyOpacity(OfficeColor color, double? opacity) =>
        ApplyOpacity((OfficeColor?)color, opacity);

    private static OfficeLinearGradient ApplyOpacity(OfficeLinearGradient gradient, double? opacity) {
        if (!opacity.HasValue) {
            return gradient;
        }

        var stops = new List<OfficeGradientStop>(gradient.Stops.Count);
        for (int i = 0; i < gradient.Stops.Count; i++) {
            OfficeGradientStop stop = gradient.Stops[i];
            stops.Add(new OfficeGradientStop(
                stop.Offset,
                ApplyOpacity(stop.Color, opacity) ?? stop.Color));
        }

        return OfficeLinearGradient.CreateImported(
            gradient.StartX,
            gradient.StartY,
            gradient.EndX,
            gradient.EndY,
            stops).WithColorInterpolation(gradient.ColorInterpolation).WithSeparateAlphaInterpolation(gradient.InterpolateAlphaSeparately);
    }

    private static OfficeRadialGradient ApplyOpacity(OfficeRadialGradient gradient, double? opacity) {
        if (!opacity.HasValue) {
            return gradient;
        }

        var stops = new List<OfficeGradientStop>(gradient.Stops.Count);
        for (int i = 0; i < gradient.Stops.Count; i++) {
            OfficeGradientStop stop = gradient.Stops[i];
            stops.Add(new OfficeGradientStop(
                stop.Offset,
                ApplyOpacity(stop.Color, opacity) ?? stop.Color));
        }

        return gradient.WithStops(stops);
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
            glowShape.StrokeGradient = null;
            glowShape.StrokeRadialGradient = null;
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
        double baseStrokeWidth = Math.Max(0D, shape.StrokeWidth);
        IReadOnlyList<OfficeShadowLayer> layers = OfficeShadowLayerPlanner.Create(
            shadow.Opacity,
            shadow.BlurRadius,
            baseStrokeWidth,
            hasFill,
            hasStroke,
            OfficeShadowLayerPlanner.CanExpand(shape),
            Math.Min(shape.Width, shape.Height));
        var shadowShapes = new List<OfficeDrawingShape>(layers.Count);
        for (int index = 0; index < layers.Count; index++) {
            OfficeShadowLayer layer = layers[index];
            shadowShapes.Add(CreateShadowShape(drawingShape, shadow, layer));
        }
        return shadowShapes;
    }

    private static OfficeDrawingShape CreateShadowShape(OfficeDrawingShape drawingShape, OfficeShadow shadow, OfficeShadowLayer layer) {
        OfficeShape shape = drawingShape.Shape;
        OfficeShape shadowShape = Math.Abs(layer.Expansion) > 0.000000001D
            ? OfficeShadowLayerPlanner.CreateExpandedShape(shape, layer.Expansion)
            : shape.Clone();
        shadowShape.Shadow = null;
        shadowShape.Glow = null;
        shadowShape.FillGradient = null;
        shadowShape.FillRadialGradient = null;
        shadowShape.FillColor = layer.HasFill || !layer.HasStroke ? shadow.Color : null;
        shadowShape.FillOpacity = layer.Opacity;
        shadowShape.StrokeColor = layer.HasStroke ? shadow.Color : null;
        shadowShape.StrokeGradient = null;
        shadowShape.StrokeRadialGradient = null;
        shadowShape.StrokeWidth = layer.StrokeWidth;
        shadowShape.StrokeDashStyle = OfficeStrokeDashStyle.Solid;
        shadowShape.StrokeStartMarker = null;
        shadowShape.StrokeEndMarker = null;
        shadowShape.StrokeOpacity = layer.Opacity;

        return CreateOffsetEffectShape(
            shadowShape,
            drawingShape.X + shadow.OffsetX - layer.Expansion,
            drawingShape.Y + shadow.OffsetY - layer.Expansion);
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
}
