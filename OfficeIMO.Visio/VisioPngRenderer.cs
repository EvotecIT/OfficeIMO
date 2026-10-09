using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static partial class VisioPngRenderer {

        internal static OfficeRasterImage RenderRaster(VisioPage page, VisioPngSaveOptions options, VisioRenderLayerVisibility layerVisibility) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (options.PixelsPerInch <= 0D || double.IsNaN(options.PixelsPerInch) || double.IsInfinity(options.PixelsPerInch)) {
                throw new ArgumentOutOfRangeException(nameof(options), "PixelsPerInch must be a finite positive number.");
            }

            if (options.Supersampling < 1 || options.Supersampling > 4) {
                throw new ArgumentOutOfRangeException(nameof(options), "Supersampling must be between 1 and 4.");
            }

            VisioRenderProjection projection = VisioRenderProjection.Create(page, options.PixelsPerInch, options.Supersampling);
            int width = Math.Max(1, (int)Math.Ceiling(projection.WidthInches * options.PixelsPerInch));
            int height = Math.Max(1, (int)Math.Ceiling(projection.HeightInches * options.PixelsPerInch));
            RasterCanvas canvas = new(
                width,
                height,
                options.Supersampling,
                options.BackgroundColor,
                ResolveTextFont(options),
                options.Fonts,
                options.TextShapingProvider,
                options.TextShapingLanguage,
                options.ImageDiagnostics,
                options.ImageDiagnosticSource,
                options.CancellationToken);
            foreach (VisioPage contentPage in VisioBackgroundComposition.Resolve(page, options.CancellationToken, options.ImageDiagnostics, options.ImageDiagnosticSource)) {
                canvas.Projection = VisioRenderProjection.CreateForContent(contentPage, page, options.PixelsPerInch, options.Supersampling);
                var contentVisibility = ReferenceEquals(contentPage, page) ? layerVisibility : new VisioRenderLayerVisibility(contentPage, options.LayerMode);
                var textStyles = new VisioNativeTextStyleResolver(contentPage.OwnerDocument, options.CancellationToken, options.ImageDiagnostics, options.ImageDiagnosticSource);
                foreach (VisioShape shape in contentPage.Shapes) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    DrawShape(canvas, contentPage, shape, options, textStyles, contentVisibility);
                }
                VisioRenderLabelLayout? labelLayout = options.ResolveConnectorLabelOverlaps
                    ? VisioRenderLabelLayout.Create(contentPage, contentVisibility) : null;
                foreach (VisioConnector connector in contentPage.Connectors) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    if (!contentVisibility.IsVisible(connector)) continue;
                    DrawConnector(canvas, contentPage, connector, options, labelLayout, textStyles);
                }
            }

            return OfficeRasterImage.FromRgba32(width, height, canvas.Resolve());
        }

        private static OfficeTrueTypeFont? ResolveTextFont(VisioPngSaveOptions options) {
            if (!string.IsNullOrWhiteSpace(options.FontFilePath)) {
                OfficeTrueTypeFont? configured = OfficeTrueTypeFont.TryLoad(options.FontFilePath, options.FontCollectionIndex, options.FontFaceName);
                if (configured != null) {
                    return configured;
                }
            }

            return null;
        }

        private static void DrawShape(RasterCanvas canvas, VisioPage page, VisioShape shape, VisioPngSaveOptions options, VisioNativeTextStyleResolver textStyles, VisioRenderLayerVisibility layerVisibility) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (layerVisibility.IsVisible(shape)) DrawShapeContent(canvas, page, shape, options, textStyles);
            foreach (VisioShape child in shape.Children) {
                DrawShape(canvas, page, child, options, textStyles, layerVisibility);
            }
        }

        private static void DrawShapeContent(RasterCanvas canvas, VisioPage page, VisioShape shape, VisioPngSaveOptions options, VisioNativeTextStyleResolver textStyles) {
            VisioNativeShapeTransform transform;
            try {
                transform = VisioNativeShapeTransform.Create(shape, options.ImageDiagnostics, options.ImageDiagnosticSource);
            } catch (Exception exception) when (exception is ArgumentException || exception is InvalidDataException) {
                VisioNativeShapeTransform.ReportInvalid(shape, options.ImageDiagnostics, options.ImageDiagnosticSource, exception);
                return;
            }
            string kind = VisioShapeGeometry.ResolveRenderKind(shape);
            bool foreign = VisioForeignImage.IsForeign(shape);
            if (foreign) {
                if (VisioForeignImage.TryGetProjection(shape, page, canvas.Projection.GeometryDensity, options.ImageDiagnostics, options.ImageDiagnosticSource, out OfficeImageProjection projection))
                    canvas.DrawImage(VisioForeignImage.Decode(shape, options.ImageCodec, options.ImageDiagnostics, options.ImageDiagnosticSource, options.CancellationToken), projection.Translate(0, canvas.Projection.ContentOffsetY));
            } else if (VisioShapeGeometry.TryGetRenderClosedPaths(shape, out List<VisioShapeGeometryPath> preservedPaths)) {
                List<RenderedPreservedPath> renderedPaths = new();
                foreach (VisioShapeGeometryPath preservedPath in preservedPaths) {
                    List<(double X, double Y)> points = new();
                    for (int i = 0; i < preservedPath.Points.Count; i++) {
                        OfficePoint point = transform.PagePoint(preservedPath.Points[i].X, preservedPath.Points[i].Y);
                        points.Add(ToRaster(page, point.X, point.Y, canvas.Projection));
                    }

                    renderedPaths.Add(new RenderedPreservedPath(preservedPath, points));
                }

                Color fill = shape.FillPattern == 0 ? Color.Transparent : shape.FillColor;
                Color stroke = HasVisibleLine(shape) ? shape.LineColor : Color.Transparent;
                double strokeWidth = Math.Max(shape.LineWeight * canvas.Projection.PhysicalDensity, canvas.Supersampling);
                for (int i = 0; i < renderedPaths.Count;) {
                    RenderedPreservedPath renderedPath = renderedPaths[i];
                    int fillGroup = renderedPath.Path.FillGroup;
                    List<List<(double X, double Y)>> contours = new();
                    int end = i + 1;
                    while (end < renderedPaths.Count &&
                           renderedPaths[end].Path.FillGroup == fillGroup) {
                        end++;
                    }

                    for (int pathIndex = i; pathIndex < end; pathIndex++) {
                        if (renderedPaths[pathIndex].Path.CanFill) contours.Add(renderedPaths[pathIndex].Points);
                    }
                    // Fill closes each eligible contour implicitly; stroke closure stays native.
                    if (fill.A > 0 && contours.Count > 0) canvas.FillPolygonsEvenOdd(contours, fill);
                    for (int pathIndex = i; pathIndex < end; pathIndex++) {
                        StrokeRenderedPreservedPath(canvas, renderedPaths[pathIndex], stroke, strokeWidth, OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern));
                    }

                    i = end;
                }
            } else if (kind == "ellipse" || kind == "circle") {
                (double centerX, double centerY) = GetPagePoint(shape, shape.Width / 2D, shape.Height / 2D);
                (double cx, double cy) = ToRaster(page, centerX, centerY, canvas.Projection);
                canvas.DrawEllipse(
                    cx,
                    cy,
                    Math.Abs(shape.Width * canvas.Projection.GeometryDensity / 2D),
                    Math.Abs(shape.Height * canvas.Projection.GeometryDensity / 2D),
                    shape.FillPattern == 0 ? Color.Transparent : shape.FillColor,
                    HasVisibleLine(shape) ? shape.LineColor : Color.Transparent,
                    Math.Max(shape.LineWeight * canvas.Projection.PhysicalDensity, canvas.Supersampling),
                    OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern),
                    ToRasterRotation(Math.Atan2(transform.Matrix.M12, transform.Matrix.M11)),
                    cx,
                    cy);
            } else if (kind == "database") {
                DrawDatabaseShape(canvas, page, shape);
                if (transform.HasReflection) VisioNativeShapeTransform.ReportArtwork(shape, options.ImageDiagnostics, options.ImageDiagnosticSource);
            } else {
                List<(double X, double Y)> local = VisioShapeGeometry.GetBuiltinClosedPath(shape, kind);
                List<(double X, double Y)> points = new();
                for (int i = 0; i < local.Count; i++) {
                    OfficePoint point = transform.PagePoint(local[i].X, local[i].Y);
                    points.Add(ToRaster(page, point.X, point.Y, canvas.Projection));
                }

                canvas.FillPolygon(points, shape.FillPattern == 0 ? Color.Transparent : shape.FillColor);
                canvas.StrokePolygon(points, HasVisibleLine(shape) ? shape.LineColor : Color.Transparent, Math.Max(shape.LineWeight * canvas.Projection.PhysicalDensity, canvas.Supersampling), OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern));
            }

            if (!foreign && options.RenderStencilArtwork) {
                bool preview = DrawPackagePreviewArtwork(canvas, page, shape, options);
                if (!preview) {
                    DrawStencilArtwork(canvas, page, shape);
                }
                if (transform.HasReflection && (preview || !string.IsNullOrEmpty(VisioStencilArtwork.GetKey(shape))))
                    VisioNativeShapeTransform.ReportArtwork(shape, options.ImageDiagnostics, options.ImageDiagnosticSource);
            }

            if (options.RenderText && !string.IsNullOrEmpty(shape.Text)) {
                VisioTextStyle? style = shape.TextStyle;
                VisioTextFramePlacement frame = VisioTextFramePlacement.Resolve(shape, canvas.Projection.DrawingToPhysical, transform);
                (double x, double y) = ToRaster(page, frame.PageX, frame.PageY, canvas.Projection);
                DrawText(
                    canvas,
                    shape.Text!,
                    x,
                    y,
                    style,
                    10D,
                    Math.Max(canvas.Supersampling * 12D, frame.ContentWidth * canvas.Projection.GeometryDensity),
                    Math.Max(canvas.Supersampling * 8D, frame.ContentHeight * canvas.Projection.GeometryDensity),
                    ToRasterRotation(frame.Angle),
                    false,
                    VisioRichTextProjection.Create(page, shape, canvas.Projection.PhysicalDensity, options.CancellationToken, textStyles));
            }

        }


        private static void StrokeRenderedPreservedPath(
            RasterCanvas canvas,
            RenderedPreservedPath renderedPath,
            Color stroke,
            double strokeWidth,
            OfficeStrokeDashStyle dashStyle) {
            Color pathStroke = renderedPath.Path.NoLine ? Color.Transparent : stroke;
            if (renderedPath.Path.IsClosed) {
                canvas.StrokePolygon(renderedPath.Points, pathStroke, strokeWidth, dashStyle);
            } else {
                canvas.StrokePolyline(renderedPath.Points, pathStroke, strokeWidth, dashStyle);
            }
        }

        private readonly struct RenderedPreservedPath {
            internal RenderedPreservedPath(VisioShapeGeometryPath path, List<(double X, double Y)> points) {
                Path = path;
                Points = points;
            }

            internal VisioShapeGeometryPath Path { get; }

            internal List<(double X, double Y)> Points { get; }
        }

        private static bool HasVisibleLine(VisioShape shape) =>
            shape.LinePattern != 0 && shape.LineWeight > 0D && shape.LineColor.A > 0;

        private static void DrawDatabaseShape(RasterCanvas canvas, VisioPage page, VisioShape shape) {
            double capHeight = Math.Min(shape.Height * 0.18D, shape.Width * 0.16D);
            double midX = shape.Width / 2D;
            (double topX, double topY) = ToRasterPoint(page, shape, midX, shape.Height - capHeight, canvas.Projection);
            (double bottomX, double bottomY) = ToRasterPoint(page, shape, midX, capHeight, canvas.Projection);
            double radiusX = Math.Max(0.5D, shape.Width * canvas.Projection.GeometryDensity / 2D);
            double radiusY = Math.Max(0.5D, capHeight * canvas.Projection.GeometryDensity);
            Color fill = shape.FillPattern == 0 ? Color.Transparent : shape.FillColor;
            Color stroke = HasVisibleLine(shape) ? shape.LineColor : Color.Transparent;
            double strokeWidth = Math.Max(shape.LineWeight * canvas.Projection.PhysicalDensity, canvas.Supersampling);
            OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern);

            List<(double X, double Y)> body = new() {
                ToRasterPoint(page, shape, 0D, capHeight, canvas.Projection),
                ToRasterPoint(page, shape, 0D, shape.Height - capHeight, canvas.Projection),
                ToRasterPoint(page, shape, shape.Width, shape.Height - capHeight, canvas.Projection),
                ToRasterPoint(page, shape, shape.Width, capHeight, canvas.Projection)
            };

            canvas.FillPolygon(body, fill);
            double rasterRotation = ToRasterRotation(shape.Angle);
            canvas.DrawEllipse(bottomX, bottomY, radiusX, radiusY, fill, Color.Transparent, strokeWidth, dashStyle, rasterRotation, bottomX, bottomY);
            canvas.DrawEllipse(topX, topY, radiusX, radiusY, fill, Color.Transparent, strokeWidth, dashStyle, rasterRotation, topX, topY);
            if (stroke.A == 0) {
                return;
            }

            canvas.StrokePolyline(
                new[] {
                    ToRasterPoint(page, shape, 0D, capHeight, canvas.Projection),
                    ToRasterPoint(page, shape, 0D, shape.Height - capHeight, canvas.Projection)
                },
                stroke,
                strokeWidth,
                dashStyle);
            canvas.StrokePolyline(
                new[] {
                    ToRasterPoint(page, shape, shape.Width, capHeight, canvas.Projection),
                    ToRasterPoint(page, shape, shape.Width, shape.Height - capHeight, canvas.Projection)
                },
                stroke,
                strokeWidth,
                dashStyle);
            canvas.DrawEllipse(bottomX, bottomY, radiusX, radiusY, Color.Transparent, stroke, strokeWidth, dashStyle, rasterRotation, bottomX, bottomY);
            canvas.DrawEllipse(topX, topY, radiusX, radiusY, Color.Transparent, stroke, strokeWidth, dashStyle, rasterRotation, topX, topY);
        }

        private static void DrawConnector(RasterCanvas canvas, VisioPage page, VisioConnector connector, VisioPngSaveOptions options, VisioRenderLabelLayout? labelLayout, VisioNativeTextStyleResolver textStyles) {
            List<(double X, double Y)> pagePoints = GetConnectorPoints(connector);
            List<(double X, double Y)> points = new();
            for (int i = 0; i < pagePoints.Count; i++) {
                points.Add(ToRaster(page, pagePoints[i].X, pagePoints[i].Y, canvas.Projection));
            }

            bool visibleLine = VisioConnectorGeometry.HasVisibleLine(connector);
            double weight = Math.Max(connector.LineWeight * canvas.Projection.PhysicalDensity, canvas.Supersampling);
            canvas.StrokePolyline(points, visibleLine ? connector.LineColor : Color.Transparent, weight, OfficeStrokeDashStyleMapper.FromVisioLinePattern(connector.LinePattern));

            if (visibleLine && connector.BeginArrow.HasValue && connector.BeginArrow.Value != EndArrow.None && OfficeGeometry.TryGetArrowheadSegment(points, fromStart: true, out (double X, double Y) beginTip, out (double X, double Y) beginFrom)) {
                DrawArrow(canvas, beginTip, beginFrom, connector.LineColor, weight);
            }

            if (visibleLine && connector.EndArrow.HasValue && connector.EndArrow.Value != EndArrow.None && OfficeGeometry.TryGetArrowheadSegment(points, fromStart: false, out (double X, double Y) endTip, out (double X, double Y) endFrom)) {
                DrawArrow(canvas, endTip, endFrom, connector.LineColor, weight);
            }

            if (options.RenderConnectorLabels && !string.IsNullOrEmpty(connector.Label)) {
                VisioRenderConnectorLabelPlacement label = labelLayout?.Resolve(connector, pagePoints) ?? VisioRenderLabelLayout.ResolveUnadjusted(connector, pagePoints, canvas.Projection);
                (double labelCenterX, double labelCenterY) = VisioConnectorGeometry.GetLabelCenter(connector, label.X, label.Y, label.Width, label.Height);
                (double x, double y) = ToRaster(page, labelCenterX, labelCenterY, canvas.Projection);
                double maxWidth = label.Width * canvas.Projection.GeometryDensity;
                double maxHeight = label.Height * canvas.Projection.GeometryDensity;
                DrawText(canvas, connector.Label!, x, y, connector.TextStyle, 9D, maxWidth, maxHeight,
                    ToRasterRotation(VisioConnectorLabelFrame.ResolveAngle(connector)), true,
                    VisioRichTextProjection.Create(page, connector, canvas.Projection.PhysicalDensity, options.CancellationToken, textStyles));
            }
        }

        private static void DrawArrow(RasterCanvas canvas, (double X, double Y) tip, (double X, double Y) from, Color color, double weight) {
            if (!OfficeGeometry.TryCreateArrowheadPoints(
                    new OfficePoint(tip.X, tip.Y),
                    new OfficePoint(from.X, from.Y),
                    weight,
                    out OfficePoint[] arrow,
                    minimumLength: canvas.Supersampling * 8D)) {
                return;
            }

            canvas.FillPolygon(ToTuples(arrow), color);
        }

        private static List<(double X, double Y)> ToTuples(IReadOnlyList<OfficePoint> points) {
            List<(double X, double Y)> converted = new(points.Count);
            for (int i = 0; i < points.Count; i++) {
                converted.Add((points[i].X, points[i].Y));
            }

            return converted;
        }

    }
}
