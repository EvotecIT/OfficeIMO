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

        public static byte[] Render(VisioPage page, VisioPngSaveOptions options) =>
            OfficeRasterImageEncoder.Encode(
                RenderRaster(page, options),
                OfficeImageExportFormat.Png,
                new OfficeRasterEncodingOptions {
                    DpiX = options.PixelsPerInch,
                    DpiY = options.PixelsPerInch
                });

        internal static OfficeRasterImage RenderRaster(VisioPage page, VisioPngSaveOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (options.PixelsPerInch <= 0D || double.IsNaN(options.PixelsPerInch) || double.IsInfinity(options.PixelsPerInch)) {
                throw new ArgumentOutOfRangeException(nameof(options), "PixelsPerInch must be a finite positive number.");
            }

            if (options.Supersampling < 1 || options.Supersampling > 4) {
                throw new ArgumentOutOfRangeException(nameof(options), "Supersampling must be between 1 and 4.");
            }

            int width = Math.Max(1, (int)Math.Ceiling(Math.Max(page.Width, 0.01D) * options.PixelsPerInch));
            int height = Math.Max(1, (int)Math.Ceiling(Math.Max(page.Height, 0.01D) * options.PixelsPerInch));
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
            canvas.Scale = options.PixelsPerInch * options.Supersampling;

            foreach (VisioShape shape in page.Shapes) {
                options.CancellationToken.ThrowIfCancellationRequested();
                DrawShape(canvas, page, shape, options);
            }

            VisioRenderLabelLayout? labelLayout = options.ResolveConnectorLabelOverlaps
                ? VisioRenderLabelLayout.Create(page)
                : null;
            foreach (VisioConnector connector in page.Connectors) {
                options.CancellationToken.ThrowIfCancellationRequested();
                DrawConnector(canvas, page, connector, options, labelLayout);
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

        private static void DrawShape(RasterCanvas canvas, VisioPage page, VisioShape shape, VisioPngSaveOptions options) {
            string kind = VisioShapeGeometry.ResolveRenderKind(shape);
            if (VisioShapeGeometry.TryGetRenderClosedPaths(shape, out List<VisioShapeGeometryPath> preservedPaths)) {
                List<RenderedPreservedPath> renderedPaths = new();
                foreach (VisioShapeGeometryPath preservedPath in preservedPaths) {
                    List<(double X, double Y)> points = new();
                    for (int i = 0; i < preservedPath.Points.Count; i++) {
                        (double px, double py) = GetPagePoint(shape, preservedPath.Points[i].X, preservedPath.Points[i].Y);
                        points.Add(ToRaster(page, px, py, canvas.Scale));
                    }

                    renderedPaths.Add(new RenderedPreservedPath(preservedPath, points));
                }

                Color fill = shape.FillPattern == 0 ? Color.Transparent : shape.FillColor;
                Color stroke = HasVisibleLine(shape) ? shape.LineColor : Color.Transparent;
                double strokeWidth = Math.Max(shape.LineWeight * canvas.Scale, canvas.Supersampling);
                for (int i = 0; i < renderedPaths.Count;) {
                    RenderedPreservedPath renderedPath = renderedPaths[i];
                    if (!renderedPath.Path.IsClosed || renderedPath.Path.NoFill || fill.A == 0) {
                        StrokeRenderedPreservedPath(canvas, renderedPath, stroke, strokeWidth, OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern));
                        i++;
                        continue;
                    }

                    int fillGroup = renderedPath.Path.FillGroup;
                    List<List<(double X, double Y)>> contours = new() { renderedPath.Points };
                    int end = i + 1;
                    while (end < renderedPaths.Count &&
                           renderedPaths[end].Path.IsClosed &&
                           !renderedPaths[end].Path.NoFill &&
                           renderedPaths[end].Path.FillGroup == fillGroup) {
                        contours.Add(renderedPaths[end].Points);
                        end++;
                    }

                    canvas.FillPolygonsEvenOdd(contours, fill);
                    for (int pathIndex = i; pathIndex < end; pathIndex++) {
                        StrokeRenderedPreservedPath(canvas, renderedPaths[pathIndex], stroke, strokeWidth, OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern));
                    }

                    i = end;
                }
            } else if (kind == "ellipse" || kind == "circle") {
                (double centerX, double centerY) = GetPagePoint(shape, shape.Width / 2D, shape.Height / 2D);
                (double cx, double cy) = ToRaster(page, centerX, centerY, canvas.Scale);
                canvas.DrawEllipse(
                    cx,
                    cy,
                    Math.Abs(shape.Width * canvas.Scale / 2D),
                    Math.Abs(shape.Height * canvas.Scale / 2D),
                    shape.FillPattern == 0 ? Color.Transparent : shape.FillColor,
                    HasVisibleLine(shape) ? shape.LineColor : Color.Transparent,
                    Math.Max(shape.LineWeight * canvas.Scale, canvas.Supersampling),
                    OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern),
                    ToRasterRotation(shape.Angle),
                    cx,
                    cy);
            } else if (kind == "database") {
                DrawDatabaseShape(canvas, page, shape);
            } else {
                List<(double X, double Y)> local = VisioShapeGeometry.GetBuiltinClosedPath(shape, kind);
                List<(double X, double Y)> points = new();
                for (int i = 0; i < local.Count; i++) {
                    (double px, double py) = GetPagePoint(shape, local[i].X, local[i].Y);
                    points.Add(ToRaster(page, px, py, canvas.Scale));
                }

                canvas.FillPolygon(points, shape.FillPattern == 0 ? Color.Transparent : shape.FillColor);
                canvas.StrokePolygon(points, HasVisibleLine(shape) ? shape.LineColor : Color.Transparent, Math.Max(shape.LineWeight * canvas.Scale, canvas.Supersampling), OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern));
            }

            if (options.RenderStencilArtwork) {
                if (!DrawPackagePreviewArtwork(canvas, page, shape, options)) {
                    DrawStencilArtwork(canvas, page, shape);
                }
            }

            if (options.RenderText && !string.IsNullOrEmpty(shape.Text)) {
                VisioTextStyle? style = shape.TextStyle;
                double textWidth = Math.Max(0.05D, style?.TextWidth ?? shape.Width);
                double textHeight = Math.Max(0.05D, style?.TextHeight ?? shape.Height);
                (double localX, double localY) = ResolveTextBoxCenter(
                    style?.TextPinX ?? shape.Width / 2D,
                    style?.TextPinY ?? shape.Height / 2D,
                    textWidth,
                    textHeight,
                    style);
                (double textX, double textY) = GetPagePoint(shape, localX, localY);
                (double x, double y) = ToRaster(page, textX, textY, canvas.Scale);
                double horizontalMargins = (style?.LeftMargin ?? 0.05D) + (style?.RightMargin ?? 0.05D);
                double verticalMargins = (style?.TopMargin ?? 0.03D) + (style?.BottomMargin ?? 0.03D);
                DrawText(
                    canvas,
                    shape.Text!,
                    x,
                    y,
                    style,
                    10D,
                    Math.Max(canvas.Supersampling * 12D, (textWidth - horizontalMargins) * canvas.Scale),
                    Math.Max(canvas.Supersampling * 8D, (textHeight - verticalMargins) * canvas.Scale),
                    ToRasterRotation(shape.Angle + (style?.TextAngle ?? 0D)),
                    false);
            }

            foreach (VisioShape child in shape.Children) {
                DrawShape(canvas, page, child, options);
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

        private static (double X, double Y) ResolveTextBoxCenter(double pinX, double pinY, double width, double height, VisioTextStyle? style) {
            double locPinX = style?.TextLocPinX ?? width / 2D;
            double locPinY = style?.TextLocPinY ?? height / 2D;
            return (pinX + (width / 2D) - locPinX, pinY + (height / 2D) - locPinY);
        }

        private static bool HasVisibleLine(VisioShape shape) =>
            shape.LinePattern != 0 && shape.LineWeight > 0D && shape.LineColor.A > 0;

        private static void DrawDatabaseShape(RasterCanvas canvas, VisioPage page, VisioShape shape) {
            double capHeight = Math.Min(shape.Height * 0.18D, shape.Width * 0.16D);
            double midX = shape.Width / 2D;
            (double topX, double topY) = ToRasterPoint(page, shape, midX, shape.Height - capHeight, canvas.Scale);
            (double bottomX, double bottomY) = ToRasterPoint(page, shape, midX, capHeight, canvas.Scale);
            double radiusX = Math.Max(0.5D, shape.Width * canvas.Scale / 2D);
            double radiusY = Math.Max(0.5D, capHeight * canvas.Scale);
            Color fill = shape.FillPattern == 0 ? Color.Transparent : shape.FillColor;
            Color stroke = HasVisibleLine(shape) ? shape.LineColor : Color.Transparent;
            double strokeWidth = Math.Max(shape.LineWeight * canvas.Scale, canvas.Supersampling);
            OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern);

            List<(double X, double Y)> body = new() {
                ToRasterPoint(page, shape, 0D, capHeight, canvas.Scale),
                ToRasterPoint(page, shape, 0D, shape.Height - capHeight, canvas.Scale),
                ToRasterPoint(page, shape, shape.Width, shape.Height - capHeight, canvas.Scale),
                ToRasterPoint(page, shape, shape.Width, capHeight, canvas.Scale)
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
                    ToRasterPoint(page, shape, 0D, capHeight, canvas.Scale),
                    ToRasterPoint(page, shape, 0D, shape.Height - capHeight, canvas.Scale)
                },
                stroke,
                strokeWidth,
                dashStyle);
            canvas.StrokePolyline(
                new[] {
                    ToRasterPoint(page, shape, shape.Width, capHeight, canvas.Scale),
                    ToRasterPoint(page, shape, shape.Width, shape.Height - capHeight, canvas.Scale)
                },
                stroke,
                strokeWidth,
                dashStyle);
            canvas.DrawEllipse(bottomX, bottomY, radiusX, radiusY, Color.Transparent, stroke, strokeWidth, dashStyle, rasterRotation, bottomX, bottomY);
            canvas.DrawEllipse(topX, topY, radiusX, radiusY, Color.Transparent, stroke, strokeWidth, dashStyle, rasterRotation, topX, topY);
        }

        private static void DrawConnector(RasterCanvas canvas, VisioPage page, VisioConnector connector, VisioPngSaveOptions options, VisioRenderLabelLayout? labelLayout) {
            List<(double X, double Y)> pagePoints = GetConnectorPoints(connector);
            List<(double X, double Y)> points = new();
            for (int i = 0; i < pagePoints.Count; i++) {
                points.Add(ToRaster(page, pagePoints[i].X, pagePoints[i].Y, canvas.Scale));
            }

            bool visibleLine = connector.LinePattern != 0 && connector.LineWeight > 0D && connector.LineColor.A > 0;
            double weight = Math.Max(connector.LineWeight * canvas.Scale, canvas.Supersampling);
            canvas.StrokePolyline(points, visibleLine ? connector.LineColor : Color.Transparent, weight, OfficeStrokeDashStyleMapper.FromVisioLinePattern(connector.LinePattern));

            if (visibleLine && connector.BeginArrow.HasValue && connector.BeginArrow.Value != EndArrow.None && OfficeGeometry.TryGetArrowheadSegment(points, fromStart: true, out (double X, double Y) beginTip, out (double X, double Y) beginFrom)) {
                DrawArrow(canvas, beginTip, beginFrom, connector.LineColor, weight);
            }

            if (visibleLine && connector.EndArrow.HasValue && connector.EndArrow.Value != EndArrow.None && OfficeGeometry.TryGetArrowheadSegment(points, fromStart: false, out (double X, double Y) endTip, out (double X, double Y) endFrom)) {
                DrawArrow(canvas, endTip, endFrom, connector.LineColor, weight);
            }

            if (options.RenderConnectorLabels && !string.IsNullOrEmpty(connector.Label)) {
                VisioRenderConnectorLabelPlacement label = labelLayout?.Resolve(connector, pagePoints) ?? ResolveConnectorLabel(connector, pagePoints);
                (double labelCenterX, double labelCenterY) = ResolveTextBoxCenter(label.X, label.Y, label.Width, label.Height, connector.TextStyle);
                (double x, double y) = ToRaster(page, labelCenterX, labelCenterY, canvas.Scale);
                double maxWidth = label.Width * canvas.Scale;
                double maxHeight = label.Height * canvas.Scale;
                DrawText(canvas, connector.Label!, x, y, connector.TextStyle, 9D, maxWidth, maxHeight, 0D, true);
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
