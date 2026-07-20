using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private const int DefaultSize = 256;
        private const int MaximumSize = 1024;

        internal static bool TryRasterize(byte[]? data, out OfficeRasterImage? image) =>
            TryRasterize(data, null, out image);

        internal static bool TryRasterize(
            byte[]? data,
            Func<string, byte[]?>? imageResolver,
            out OfficeRasterImage? image) =>
            TryRasterize(
                data,
                imageResolver,
                outlineFont: null,
                fonts: null,
                textShapingProvider: null,
                textShapingLanguage: null,
                diagnosticSink: null,
                diagnosticSource: null,
                cancellationToken: default,
                out image);

        internal static bool TryRasterize(
            byte[]? data,
            Func<string, byte[]?>? imageResolver,
            OfficeTrueTypeFont? outlineFont,
            OfficeFontFaceCollection? fonts,
            IOfficeTextShapingProvider? textShapingProvider,
            string? textShapingLanguage,
            ICollection<OfficeImageExportDiagnostic>? diagnosticSink,
            string? diagnosticSource,
            System.Threading.CancellationToken cancellationToken,
            out OfficeRasterImage? image) {
            image = null;
            if (data == null || data.Length == 0) {
                return false;
            }

            XDocument document;
            try {
                using var stream = new MemoryStream(data, writable: false);
                document = XDocument.Load(stream, LoadOptions.None);
            } catch {
                return false;
            }

            XElement? root = document.Root;
            if (root == null || !string.Equals(root.Name.LocalName, "svg", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            ResolveViewport(root, out double viewLeft, out double viewTop, out double viewWidth, out double viewHeight, out int width, out int height);
            if (viewWidth <= 0D || viewHeight <= 0D || width <= 0 || height <= 0) {
                return false;
            }

            OfficeRasterImage raster = new(width, height, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new(
                raster,
                outlineFont,
                fonts,
                textShapingProvider,
                textShapingLanguage,
                diagnosticSink,
                diagnosticSource,
                cancellationToken);
            SvgRenderContext context = SvgRenderContext.Create(root, new SvgPaintBounds(viewLeft, viewTop, viewWidth, viewHeight), imageResolver);
            double rootOpacity = SvgPaint.ReadOwnOpacity(root, context);
            if (rootOpacity <= 0D) {
                return false;
            }

            bool useRootOpacityLayer = rootOpacity < 1D;
            SvgPaint inherited = SvgPaint.Resolve(root, SvgPaint.Default, context, applyOwnOpacity: !useRootOpacityLayer);
            SvgTransform transform = CreateViewBoxTransform(viewLeft, viewTop, viewWidth, viewHeight, 0D, 0D, width, height, root.Attribute("preserveAspectRatio")?.Value);
            using IDisposable rootTextStyle = context.PushTextStyle(SvgTextStyle.Resolve(root, SvgTextStyle.Default, context));
            using IDisposable rootFillRule = context.PushFillRule(ResolveFillRule(root, context));
            OfficeRasterCanvas targetCanvas = canvas;
            OfficeRasterImage? rootLayer = null;
            if (useRootOpacityLayer) {
                rootLayer = new OfficeRasterImage(width, height, OfficeColor.Transparent);
                targetCanvas = CreateLayerCanvas(canvas, rootLayer);
            }

            bool rendered = RenderChildren(targetCanvas, root, inherited, transform, context);
            if (!rendered) {
                return false;
            }

            if (useRootOpacityLayer && rootLayer != null) {
                canvas.DrawImage(ApplyImageOpacity(rootLayer, rootOpacity), 0D, 0D, width, height);
            }

            image = raster;
            return true;
        }

        private static bool RenderChildren(OfficeRasterCanvas canvas, XElement element, SvgPaint inherited, SvgTransform transform, SvgRenderContext context) {
            bool rendered = false;
            foreach (XElement child in element.Elements()) {
                canvas.CancellationToken.ThrowIfCancellationRequested();
                if (RenderElement(canvas, child, inherited, transform, context)) {
                    rendered = true;
                }
            }

            return rendered;
        }

        private static bool RenderElement(OfficeRasterCanvas canvas, XElement element, SvgPaint inherited, SvgTransform transform, SvgRenderContext context) {
            canvas.CancellationToken.ThrowIfCancellationRequested();
            string name = element.Name.LocalName;
            if (string.Equals(name, "defs", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "style", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "title", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "desc", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (IsElementDisplayNone(element, context)) {
                return false;
            }

            SvgTransform localTransform = transform.Multiply(ReadTransform(element.Attribute("transform")?.Value));
            using IDisposable visibilityScope = context.PushVisibility(ReadVisibilityOverride(element, context));
            using IDisposable paintBoundsScope = context.PushPaintBounds(TryGetElementPaintBounds(element, name, context, out SvgPaintBounds bounds) ? bounds : null);
            using IDisposable textStyleScope = context.PushTextStyle(SvgTextStyle.Resolve(element, context.CurrentTextStyle, context));
            using IDisposable fillRuleScope = context.PushFillRule(ResolveFillRule(element, context));
            bool appliesElementOpacity = CanApplyElementOpacity(name);
            double elementOpacity = appliesElementOpacity ? SvgPaint.ReadOwnOpacity(element, context) : 1D;
            if (appliesElementOpacity && elementOpacity <= 0D) {
                return false;
            }

            bool useElementOpacityLayer = appliesElementOpacity && elementOpacity < 1D;
            SvgPaint paint = SvgPaint.Resolve(element, inherited, context, applyOwnOpacity: !useElementOpacityLayer);
            if (!context.IsVisible && !CanHiddenElementHaveVisibleDescendants(name)) {
                return false;
            }

            if (useElementOpacityLayer) {
                OfficeRasterImage layer = new(canvas.Width, canvas.Height, OfficeColor.Transparent);
                OfficeRasterCanvas layerCanvas = CreateLayerCanvas(canvas, layer);
                bool rendered = RenderElementCore(layerCanvas, element, name, paint, localTransform, context);
                if (!rendered) {
                    return false;
                }

                using IDisposable? groupClipScope = PushClipPath(canvas, element, localTransform, context);
                canvas.DrawImage(ApplyImageOpacity(layer, elementOpacity), 0D, 0D, canvas.Width, canvas.Height);
                return true;
            }

            using IDisposable? clipScope = PushClipPath(canvas, element, localTransform, context);
            return RenderElementCore(canvas, element, name, paint, localTransform, context);
        }

        private static OfficeRasterCanvas CreateLayerCanvas(
            OfficeRasterCanvas parent,
            OfficeRasterImage image) =>
            new(
                image,
                parent.OutlineFont,
                parent.Fonts,
                parent.TextShapingProvider,
                parent.TextShapingLanguage,
                parent.DiagnosticSink,
                parent.DiagnosticSource,
                parent.CancellationToken);

        private static bool CanApplyElementOpacity(string name) =>
            string.Equals(name, "g", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "svg", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "use", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "image", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "text", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "rect", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "circle", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "ellipse", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "line", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "polyline", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "polygon", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "path", StringComparison.OrdinalIgnoreCase);

        private static bool CanHiddenElementHaveVisibleDescendants(string name) =>
            string.Equals(name, "g", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "svg", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "use", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "text", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "tspan", StringComparison.OrdinalIgnoreCase);

        private static bool TryGetElementPaintBounds(XElement element, string name, SvgRenderContext context, out SvgPaintBounds bounds) {
            bounds = default;
            if (string.Equals(name, "rect", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "image", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "use", StringComparison.OrdinalIgnoreCase)) {
                double x = ReadLength(element, "x", 0D, context, SvgLengthAxis.X);
                double y = ReadLength(element, "y", 0D, context, SvgLengthAxis.Y);
                double width = ReadLength(element, "width", 0D, context, SvgLengthAxis.X);
                double height = ReadLength(element, "height", 0D, context, SvgLengthAxis.Y);
                if (width > 0D && height > 0D) {
                    bounds = new SvgPaintBounds(x, y, width, height);
                    return true;
                }
            }

            if (string.Equals(name, "circle", StringComparison.OrdinalIgnoreCase)) {
                double radius = ReadLength(element, "r", 0D, context, SvgLengthAxis.Diagonal);
                if (radius > 0D) {
                    double cx = ReadLength(element, "cx", 0D, context, SvgLengthAxis.X);
                    double cy = ReadLength(element, "cy", 0D, context, SvgLengthAxis.Y);
                    bounds = new SvgPaintBounds(cx - radius, cy - radius, radius * 2D, radius * 2D);
                    return true;
                }
            }

            if (string.Equals(name, "ellipse", StringComparison.OrdinalIgnoreCase)) {
                double rx = ReadLength(element, "rx", 0D, context, SvgLengthAxis.X);
                double ry = ReadLength(element, "ry", 0D, context, SvgLengthAxis.Y);
                if (rx > 0D && ry > 0D) {
                    double cx = ReadLength(element, "cx", 0D, context, SvgLengthAxis.X);
                    double cy = ReadLength(element, "cy", 0D, context, SvgLengthAxis.Y);
                    bounds = new SvgPaintBounds(cx - rx, cy - ry, rx * 2D, ry * 2D);
                    return true;
                }
            }

            if (string.Equals(name, "line", StringComparison.OrdinalIgnoreCase)) {
                double x1 = ReadLength(element, "x1", 0D, context, SvgLengthAxis.X);
                double y1 = ReadLength(element, "y1", 0D, context, SvgLengthAxis.Y);
                double x2 = ReadLength(element, "x2", 0D, context, SvgLengthAxis.X);
                double y2 = ReadLength(element, "y2", 0D, context, SvgLengthAxis.Y);
                bounds = CreatePaintBounds(new[] { (x1, y1), (x2, y2) });
                return true;
            }

            if (string.Equals(name, "polyline", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "polygon", StringComparison.OrdinalIgnoreCase)) {
                if (TryParsePoints(element.Attribute("points")?.Value, out List<(double X, double Y)> points) && points.Count > 0) {
                    bounds = CreatePaintBounds(points);
                    return true;
                }
            }

            if (string.Equals(name, "path", StringComparison.OrdinalIgnoreCase)) {
                if (TryParsePath(element.Attribute("d")?.Value, out List<SvgPathContour> contours)) {
                    var points = new List<(double X, double Y)>();
                    for (int i = 0; i < contours.Count; i++) {
                        points.AddRange(contours[i].Points);
                    }

                    if (points.Count > 0) {
                        bounds = CreatePaintBounds(points);
                        return true;
                    }
                }
            }

            return false;
        }

        private static SvgPaintBounds CreatePaintBounds(IReadOnlyList<(double X, double Y)> points) {
            double left = points[0].X;
            double right = points[0].X;
            double top = points[0].Y;
            double bottom = points[0].Y;
            for (int i = 1; i < points.Count; i++) {
                left = Math.Min(left, points[i].X);
                right = Math.Max(right, points[i].X);
                top = Math.Min(top, points[i].Y);
                bottom = Math.Max(bottom, points[i].Y);
            }

            return new SvgPaintBounds(left, top, right - left, bottom - top);
        }

        private static bool RenderElementCore(OfficeRasterCanvas canvas, XElement element, string name, SvgPaint paint, SvgTransform localTransform, SvgRenderContext context) {
            if (string.Equals(name, "g", StringComparison.OrdinalIgnoreCase)) {
                return RenderChildren(canvas, element, paint, localTransform, context);
            }

            if (string.Equals(name, "svg", StringComparison.OrdinalIgnoreCase)) {
                return RenderNestedSvg(canvas, element, paint, localTransform, context);
            }

            if (string.Equals(name, "use", StringComparison.OrdinalIgnoreCase)) {
                return RenderUse(canvas, element, paint, localTransform, context);
            }

            if (string.Equals(name, "image", StringComparison.OrdinalIgnoreCase)) {
                return RenderImage(canvas, element, paint, localTransform, context);
            }

            if (string.Equals(name, "text", StringComparison.OrdinalIgnoreCase)) {
                return RenderText(canvas, element, paint, localTransform, context);
            }

            if (string.Equals(name, "rect", StringComparison.OrdinalIgnoreCase)) {
                return RenderRectangle(canvas, element, paint, localTransform, context);
            }

            if (string.Equals(name, "circle", StringComparison.OrdinalIgnoreCase)) {
                double radius = ReadLength(element, "r", 0D, context, SvgLengthAxis.Diagonal);
                return RenderEllipse(canvas, ReadLength(element, "cx", 0D, context, SvgLengthAxis.X), ReadLength(element, "cy", 0D, context, SvgLengthAxis.Y), radius, radius, paint, localTransform);
            }

            if (string.Equals(name, "ellipse", StringComparison.OrdinalIgnoreCase)) {
                return RenderEllipse(canvas, ReadLength(element, "cx", 0D, context, SvgLengthAxis.X), ReadLength(element, "cy", 0D, context, SvgLengthAxis.Y), ReadLength(element, "rx", 0D, context, SvgLengthAxis.X), ReadLength(element, "ry", 0D, context, SvgLengthAxis.Y), paint, localTransform);
            }

            if (string.Equals(name, "line", StringComparison.OrdinalIgnoreCase)) {
                return RenderPolyline(canvas, new[] {
                    (ReadLength(element, "x1", 0D, context, SvgLengthAxis.X), ReadLength(element, "y1", 0D, context, SvgLengthAxis.Y)),
                    (ReadLength(element, "x2", 0D, context, SvgLengthAxis.X), ReadLength(element, "y2", 0D, context, SvgLengthAxis.Y))
                }, closed: false, paint, localTransform);
            }

            if (string.Equals(name, "polyline", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "polygon", StringComparison.OrdinalIgnoreCase)) {
                if (!TryParsePoints(element.Attribute("points")?.Value, out List<(double X, double Y)> points)) {
                    return false;
                }

                return RenderPolyline(canvas, points, string.Equals(name, "polygon", StringComparison.OrdinalIgnoreCase), paint, localTransform);
            }

            if (string.Equals(name, "path", StringComparison.OrdinalIgnoreCase)) {
                if (!TryParsePath(element.Attribute("d")?.Value, out List<SvgPathContour> contours)) {
                    return false;
                }

                return RenderPath(canvas, element, contours, paint, localTransform, context);
            }

            return RenderChildren(canvas, element, paint, localTransform, context);
        }

        private static bool RenderUse(OfficeRasterCanvas canvas, XElement element, SvgPaint inherited, SvgTransform transform, SvgRenderContext context) {
            string? href = ReadHref(element);
            if (string.IsNullOrWhiteSpace(href) || href![0] != '#') {
                return false;
            }

            string id = href.Substring(1);
            if (id.Length == 0 || !context.TryGetDefinition(id, out XElement? definition) || definition == null || !context.TryEnterUse(id)) {
                return false;
            }

            try {
                string definitionName = definition.Name.LocalName;
                if (string.Equals(definitionName, "symbol", StringComparison.OrdinalIgnoreCase)) {
                    return RenderSymbolUse(canvas, element, definition, inherited, transform, context);
                }

                SvgTransform useTransform = transform.Multiply(SvgTransform.Create(1D, 0D, 0D, 1D, ReadLength(element, "x", 0D, context, SvgLengthAxis.X), ReadLength(element, "y", 0D, context, SvgLengthAxis.Y)));
                if (string.Equals(definitionName, "svg", StringComparison.OrdinalIgnoreCase)) {
                    return RenderSymbolUse(canvas, element, definition, inherited, transform, context);
                }

                return RenderElement(canvas, definition, inherited, useTransform, context);
            } finally {
                context.ExitUse(id);
            }
        }

        private static bool RenderSymbolUse(OfficeRasterCanvas canvas, XElement useElement, XElement symbol, SvgPaint inherited, SvgTransform transform, SvgRenderContext context) {
            double x = ReadLength(useElement, "x", 0D, context, SvgLengthAxis.X);
            double y = ReadLength(useElement, "y", 0D, context, SvgLengthAxis.Y);
            double viewLeft = 0D;
            double viewTop = 0D;
            double viewWidth = ReadLength(useElement, "width", context.ViewportBounds.Width, context, SvgLengthAxis.X);
            double viewHeight = ReadLength(useElement, "height", context.ViewportBounds.Height, context, SvgLengthAxis.Y);
            bool hasViewBox = TryParseNumbers(symbol.Attribute("viewBox")?.Value, out List<double> viewBox) &&
                viewBox.Count >= 4 &&
                viewBox[2] > 0D &&
                viewBox[3] > 0D;
            if (hasViewBox) {
                viewLeft = viewBox[0];
                viewTop = viewBox[1];
                viewWidth = viewBox[2];
                viewHeight = viewBox[3];
            }

            double width = ReadLength(useElement, "width", viewWidth, context, SvgLengthAxis.X);
            double height = ReadLength(useElement, "height", viewHeight, context, SvgLengthAxis.Y);
            if (width <= 0D || height <= 0D || viewWidth <= 0D || viewHeight <= 0D) {
                return false;
            }

            SvgTransform contentTransform = CreateViewBoxTransform(
                viewLeft,
                viewTop,
                viewWidth,
                viewHeight,
                x,
                y,
                width,
                height,
                useElement.Attribute("preserveAspectRatio")?.Value ?? symbol.Attribute("preserveAspectRatio")?.Value);
            contentTransform = transform.Multiply(contentTransform);
            IReadOnlyList<OfficePoint> clip = ProjectPoints(new[] {
                (x, y),
                (x + width, y),
                (x + width, y + height),
                (x, y + height)
            }, transform);
            using IDisposable clipScope = canvas.PushClipPolygon(clip);
            using IDisposable viewportScope = context.PushViewportBounds(new SvgPaintBounds(viewLeft, viewTop, viewWidth, viewHeight));
            using IDisposable textStyleScope = context.PushTextStyle(SvgTextStyle.Resolve(symbol, context.CurrentTextStyle, context));
            using IDisposable fillRuleScope = context.PushFillRule(ResolveFillRule(symbol, context));
            SvgPaint symbolPaint = SvgPaint.Resolve(symbol, inherited, context);
            return RenderChildren(canvas, symbol, symbolPaint, contentTransform, context);
        }

        private static bool RenderNestedSvg(OfficeRasterCanvas canvas, XElement element, SvgPaint inherited, SvgTransform transform, SvgRenderContext context) {
            double x = ReadLength(element, "x", 0D, context, SvgLengthAxis.X);
            double y = ReadLength(element, "y", 0D, context, SvgLengthAxis.Y);
            double width = ReadLength(element, "width", context.ViewportBounds.Width, context, SvgLengthAxis.X);
            double height = ReadLength(element, "height", context.ViewportBounds.Height, context, SvgLengthAxis.Y);
            if (width <= 0D || height <= 0D) {
                return false;
            }

            SvgTransform contentTransform = transform.Multiply(SvgTransform.Create(1D, 0D, 0D, 1D, x, y));
            SvgPaintBounds viewportBounds = new(0D, 0D, width, height);
            if (TryParseNumbers(element.Attribute("viewBox")?.Value, out List<double> viewBox) &&
                viewBox.Count >= 4 &&
                viewBox[2] > 0D &&
                viewBox[3] > 0D) {
                contentTransform = CreateViewBoxTransform(viewBox[0], viewBox[1], viewBox[2], viewBox[3], x, y, width, height, element.Attribute("preserveAspectRatio")?.Value);
                contentTransform = transform.Multiply(contentTransform);
                viewportBounds = new SvgPaintBounds(viewBox[0], viewBox[1], viewBox[2], viewBox[3]);
            }

            IReadOnlyList<OfficePoint> clip = ProjectPoints(new[] {
                (x, y),
                (x + width, y),
                (x + width, y + height),
                (x, y + height)
            }, transform);
            using IDisposable clipScope = canvas.PushClipPolygon(clip);
            using IDisposable viewportScope = context.PushViewportBounds(viewportBounds);
            return RenderChildren(canvas, element, inherited, contentTransform, context);
        }

        private static bool RenderRectangle(OfficeRasterCanvas canvas, XElement element, SvgPaint paint, SvgTransform transform, SvgRenderContext context) {
            double x = ReadLength(element, "x", 0D, context, SvgLengthAxis.X);
            double y = ReadLength(element, "y", 0D, context, SvgLengthAxis.Y);
            double width = ReadLength(element, "width", 0D, context, SvgLengthAxis.X);
            double height = ReadLength(element, "height", 0D, context, SvgLengthAxis.Y);
            if (width <= 0D || height <= 0D) {
                return false;
            }

            bool hasRx = TryParseLength(element.Attribute("rx")?.Value, GetLengthReference(context, SvgLengthAxis.X), out double rx);
            bool hasRy = TryParseLength(element.Attribute("ry")?.Value, GetLengthReference(context, SvgLengthAxis.Y), out double ry);
            if (hasRx && !hasRy) {
                ry = rx;
            } else if (!hasRx && hasRy) {
                rx = ry;
            }

            rx = Math.Min(Math.Abs(rx), width / 2D);
            ry = Math.Min(Math.Abs(ry), height / 2D);
            List<(double X, double Y)> points = rx > 0D && ry > 0D
                ? CreateRoundedRectanglePoints(x, y, width, height, rx, ry)
                : new List<(double X, double Y)> {
                    (x, y),
                    (x + width, y),
                    (x + width, y + height),
                    (x, y + height)
                };

            return RenderPolyline(canvas, points, closed: true, paint, transform);
        }

        private static bool RenderEllipse(OfficeRasterCanvas canvas, double cx, double cy, double rx, double ry, SvgPaint paint, SvgTransform transform) {
            if (rx <= 0D || ry <= 0D) {
                return false;
            }

            IReadOnlyList<OfficePoint> ellipsePoints = CreateEllipsePoints(cx, cy, rx, ry, transform);
            if (paint.HasFill) {
                if (paint.FillRadialGradient != null || paint.FillGradient != null) {
                    if (paint.FillRadialGradient != null) {
                        canvas.FillRadialGradientPolygon(ellipsePoints, paint.FillRadialGradient);
                    } else {
                        canvas.FillLinearGradientPolygon(ellipsePoints, paint.FillGradient!);
                    }
                } else {
                    canvas.FillPolygon(ellipsePoints, paint.Fill);
                }
            }

            if (paint.HasStroke && paint.StrokeWidth > 0D) {
                double strokeScale = GetStrokeScale(paint, transform);
                double strokeWidth = Math.Max(1D, paint.StrokeWidth * strokeScale);
                IReadOnlyList<double>? dashPattern = ScaleDashPattern(paint.DashPattern, strokeScale);
                StrokeClosedContour(canvas, ellipsePoints, paint, strokeWidth, dashPattern);
            }

            return paint.HasFill || paint.HasStroke;
        }

        private static bool RenderPath(OfficeRasterCanvas canvas, XElement element, IReadOnlyList<SvgPathContour> contours, SvgPaint paint, SvgTransform transform, SvgRenderContext context) {
            if (contours.Count == 0) {
                return false;
            }

            bool rendered = false;
            List<IReadOnlyList<OfficePoint>> closedContours = new();
            List<(IReadOnlyList<OfficePoint> Points, bool Closed)> projectedContours = new(contours.Count);
            for (int i = 0; i < contours.Count; i++) {
                List<OfficePoint> projected = ProjectPoints(contours[i].Points, transform);
                if (projected.Count < 2) {
                    continue;
                }

                bool closed = contours[i].IsClosed && projected.Count >= 3;
                projectedContours.Add((projected, closed));
                if (projected.Count >= 3) {
                    closedContours.Add(projected);
                }
            }

            if (paint.HasFill && closedContours.Count > 0) {
                bool useEvenOddFill = context.CurrentFillRule == OfficeFillRule.EvenOdd;
                if (paint.FillRadialGradient != null) {
                    FillGradientContours(canvas, closedContours, null, paint.FillRadialGradient, useEvenOddFill);
                } else if (paint.FillGradient != null) {
                    FillGradientContours(canvas, closedContours, paint.FillGradient, null, useEvenOddFill);
                } else if (useEvenOddFill) {
                    if (closedContours.Count > 1) {
                        canvas.FillPolygonsEvenOdd(closedContours, paint.Fill);
                    } else {
                        canvas.FillPolygon(closedContours[0], paint.Fill);
                    }
                } else {
                    canvas.FillPolygonsNonZero(closedContours, paint.Fill);
                }

                rendered = true;
            }

            if (paint.HasStroke && paint.StrokeWidth > 0D) {
                double strokeScale = GetStrokeScale(paint, transform);
                double strokeWidth = Math.Max(1D, paint.StrokeWidth * strokeScale);
                IReadOnlyList<double>? dashPattern = ScaleDashPattern(paint.DashPattern, strokeScale);
                for (int i = 0; i < projectedContours.Count; i++) {
                    if (projectedContours[i].Closed) {
                        StrokeClosedContour(canvas, projectedContours[i].Points, paint, strokeWidth, dashPattern);
                    } else {
                        StrokeOpenContour(canvas, projectedContours[i].Points, paint, strokeWidth, dashPattern);
                    }
                }

                rendered = true;
            }

            return rendered;
        }

        private static void FillGradientContours(
            OfficeRasterCanvas canvas,
            IReadOnlyList<IReadOnlyList<OfficePoint>> contours,
            OfficeLinearGradient? linearGradient,
            OfficeRadialGradient? radialGradient,
            bool useEvenOddFill) {
            if (contours.Count == 1 && useEvenOddFill) {
                FillGradientContour(canvas, contours[0], linearGradient, radialGradient);
                return;
            }

            using IDisposable clipScope = useEvenOddFill
                ? canvas.PushClipPolygonsEvenOdd(contours)
                : canvas.PushClipPolygonsNonZero(contours);
            if (!TryGetContourBounds(contours, out double left, out double top, out double right, out double bottom)) {
                return;
            }

            var bounds = new[] {
                new OfficePoint(left, top),
                new OfficePoint(right, top),
                new OfficePoint(right, bottom),
                new OfficePoint(left, bottom)
            };
            FillGradientContour(canvas, bounds, linearGradient, radialGradient);
        }

        private static void FillGradientContour(OfficeRasterCanvas canvas, IReadOnlyList<OfficePoint> contour, OfficeLinearGradient? linearGradient, OfficeRadialGradient? radialGradient) {
            if (radialGradient != null) {
                canvas.FillRadialGradientPolygon(contour, radialGradient);
            } else if (linearGradient != null) {
                canvas.FillLinearGradientPolygon(contour, linearGradient);
            }
        }

        private static bool TryGetContourBounds(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, out double left, out double top, out double right, out double bottom) {
            left = 0D;
            top = 0D;
            right = 0D;
            bottom = 0D;
            bool hasPoint = false;
            for (int contourIndex = 0; contourIndex < contours.Count; contourIndex++) {
                IReadOnlyList<OfficePoint> contour = contours[contourIndex];
                for (int pointIndex = 0; pointIndex < contour.Count; pointIndex++) {
                    OfficePoint point = contour[pointIndex];
                    if (!hasPoint) {
                        left = right = point.X;
                        top = bottom = point.Y;
                        hasPoint = true;
                        continue;
                    }

                    left = Math.Min(left, point.X);
                    top = Math.Min(top, point.Y);
                    right = Math.Max(right, point.X);
                    bottom = Math.Max(bottom, point.Y);
                }
            }

            return hasPoint && right > left && bottom > top;
        }

        private static bool RenderPolyline(OfficeRasterCanvas canvas, IReadOnlyList<(double X, double Y)> points, bool closed, SvgPaint paint, SvgTransform transform) {
            if (points.Count < 2) {
                return false;
            }

            List<OfficePoint> projected = ProjectPoints(points, transform);

            bool filled = projected.Count >= 3 && paint.HasFill;
            if (filled) {
                if (paint.FillRadialGradient != null) {
                    canvas.FillRadialGradientPolygon(projected, paint.FillRadialGradient);
                } else if (paint.FillGradient != null) {
                    canvas.FillLinearGradientPolygon(projected, paint.FillGradient);
                } else {
                    canvas.FillPolygon(projected, paint.Fill);
                }
            }

            if (paint.HasStroke && paint.StrokeWidth > 0D) {
                double strokeScale = GetStrokeScale(paint, transform);
                double strokeWidth = Math.Max(1D, paint.StrokeWidth * strokeScale);
                IReadOnlyList<double>? dashPattern = ScaleDashPattern(paint.DashPattern, strokeScale);
                if (closed && projected.Count >= 3) {
                    StrokeClosedContour(canvas, projected, paint, strokeWidth, dashPattern);
                } else {
                    StrokeOpenContour(canvas, projected, paint, strokeWidth, dashPattern);
                }
            }

            return filled || paint.HasStroke;
        }

        private static void StrokeOpenContour(OfficeRasterCanvas canvas, IReadOnlyList<OfficePoint> points, SvgPaint paint, double strokeWidth, IReadOnlyList<double>? dashPattern) {
            if (points.Count < 2) {
                return;
            }

            if (dashPattern != null) {
                canvas.DrawPatternedPolyline(points, GetStrokeFallbackColor(paint), strokeWidth, dashPattern);
                return;
            }

            IReadOnlyList<OfficePoint> strokePoints = paint.StrokeLineCap == SvgStrokeLineCap.Square
                ? ExtendOpenStrokeForSquareCap(points, strokeWidth)
                : points;
            DrawJoinedPolyline(canvas, strokePoints, paint, strokeWidth, closed: false);

            if (paint.StrokeLineCap == SvgStrokeLineCap.Round) {
                OfficeColor capColor = GetStrokeFallbackColor(paint);
                DrawRoundStrokeCap(canvas, points[0], capColor, strokeWidth);
                DrawRoundStrokeCap(canvas, points[points.Count - 1], capColor, strokeWidth);
            }
        }

        private static IReadOnlyList<OfficePoint> ExtendOpenStrokeForSquareCap(IReadOnlyList<OfficePoint> points, double strokeWidth) {
            double distance = strokeWidth / 2D;
            List<OfficePoint> extended = new(points.Count);
            extended.Add(ExtendCapPoint(points[0], points[1], distance));
            for (int i = 1; i < points.Count - 1; i++) {
                extended.Add(points[i]);
            }

            extended.Add(ExtendCapPoint(points[points.Count - 1], points[points.Count - 2], distance));
            return extended;
        }

        private static OfficePoint ExtendCapPoint(OfficePoint endpoint, OfficePoint neighbor, double distance) {
            double dx = endpoint.X - neighbor.X;
            double dy = endpoint.Y - neighbor.Y;
            double length = Math.Sqrt((dx * dx) + (dy * dy));
            if (length <= double.Epsilon) {
                return endpoint;
            }

            double scale = distance / length;
            return new OfficePoint(endpoint.X + (dx * scale), endpoint.Y + (dy * scale));
        }

        private static void DrawJoinedPolyline(OfficeRasterCanvas canvas, IReadOnlyList<OfficePoint> points, SvgPaint paint, double strokeWidth, bool closed) {
            DrawFlatPolyline(canvas, points, paint, strokeWidth, closed);
            DrawStrokeLineJoins(canvas, points, GetStrokeFallbackColor(paint), strokeWidth, paint.StrokeLineJoin, closed);
        }

        private static void DrawFlatPolyline(OfficeRasterCanvas canvas, IReadOnlyList<OfficePoint> points, SvgPaint paint, double strokeWidth, bool closed) {
            GetPointBounds(points, out double left, out double top, out double width, out double height);
            for (int i = 1; i < points.Count; i++) {
                DrawFlatStrokeSegment(canvas, points[i - 1], points[i], paint, strokeWidth, left, top, width, height);
            }

            if (closed && points.Count > 2) {
                DrawFlatStrokeSegment(canvas, points[points.Count - 1], points[0], paint, strokeWidth, left, top, width, height);
            }
        }

        private static void DrawFlatStrokeSegment(OfficeRasterCanvas canvas, OfficePoint start, OfficePoint end, SvgPaint paint, double strokeWidth, double pathLeft, double pathTop, double pathWidth, double pathHeight) {
            double dx = end.X - start.X;
            double dy = end.Y - start.Y;
            double length = Math.Sqrt((dx * dx) + (dy * dy));
            if (length <= double.Epsilon) {
                return;
            }

            double half = strokeWidth / 2D;
            double offsetX = (-dy / length) * half;
            double offsetY = (dx / length) * half;
            OfficePoint[] polygon = {
                    new OfficePoint(start.X + offsetX, start.Y + offsetY),
                    new OfficePoint(end.X + offsetX, end.Y + offsetY),
                    new OfficePoint(end.X - offsetX, end.Y - offsetY),
                    new OfficePoint(start.X - offsetX, start.Y - offsetY)
                };
            if (paint.StrokeRadialGradient != null) {
                canvas.FillRadialGradientPolygon(polygon, paint.StrokeRadialGradient);
            } else if (paint.StrokeGradient != null) {
                canvas.FillLinearGradientPolygon(polygon, RebaseLinearGradientToSegmentBounds(paint.StrokeGradient, polygon, pathLeft, pathTop, pathWidth, pathHeight));
            } else {
                canvas.FillPolygon(polygon, paint.Stroke);
            }
        }

        private static OfficeLinearGradient RebaseLinearGradientToSegmentBounds(OfficeLinearGradient gradient, IReadOnlyList<OfficePoint> segmentPoints, double pathLeft, double pathTop, double pathWidth, double pathHeight) {
            GetPointBounds(segmentPoints, out double segmentLeft, out double segmentTop, out double segmentWidth, out double segmentHeight);
            pathWidth = Math.Max(pathWidth, 0.0001D);
            pathHeight = Math.Max(pathHeight, 0.0001D);
            segmentWidth = Math.Max(segmentWidth, 0.0001D);
            segmentHeight = Math.Max(segmentHeight, 0.0001D);
            double startX = ((pathLeft + (gradient.StartX * pathWidth)) - segmentLeft) / segmentWidth;
            double startY = ((pathTop + (gradient.StartY * pathHeight)) - segmentTop) / segmentHeight;
            double endX = ((pathLeft + (gradient.EndX * pathWidth)) - segmentLeft) / segmentWidth;
            double endY = ((pathTop + (gradient.EndY * pathHeight)) - segmentTop) / segmentHeight;
            return new OfficeLinearGradient(startX, startY, endX, endY, gradient.Stops);
        }

        private static void GetPointBounds(IReadOnlyList<OfficePoint> points, out double left, out double top, out double width, out double height) {
            left = points.Count > 0 ? points[0].X : 0D;
            double right = left;
            top = points.Count > 0 ? points[0].Y : 0D;
            double bottom = top;
            for (int i = 1; i < points.Count; i++) {
                left = Math.Min(left, points[i].X);
                right = Math.Max(right, points[i].X);
                top = Math.Min(top, points[i].Y);
                bottom = Math.Max(bottom, points[i].Y);
            }

            width = right - left;
            height = bottom - top;
        }

        private static OfficeColor GetStrokeFallbackColor(SvgPaint paint) {
            if (paint.Stroke.A > 0) {
                return paint.Stroke;
            }

            OfficeGradientStop? stop = paint.StrokeGradient?.Stops.Count > 0
                ? paint.StrokeGradient.Stops[paint.StrokeGradient.Stops.Count - 1]
                : paint.StrokeRadialGradient?.Stops.Count > 0
                    ? paint.StrokeRadialGradient.Stops[paint.StrokeRadialGradient.Stops.Count - 1]
                    : null;
            return stop?.Color ?? OfficeColor.Transparent;
        }

        private static void DrawStrokeLineJoins(OfficeRasterCanvas canvas, IReadOnlyList<OfficePoint> points, OfficeColor color, double strokeWidth, SvgStrokeLineJoin lineJoin, bool closed) {
            if (points.Count < 3) {
                return;
            }

            int start = closed ? 0 : 1;
            int end = closed ? points.Count : points.Count - 1;
            for (int i = start; i < end; i++) {
                OfficePoint previous = points[(i - 1 + points.Count) % points.Count];
                OfficePoint current = points[i];
                OfficePoint next = points[(i + 1) % points.Count];
                if (lineJoin == SvgStrokeLineJoin.Round) {
                    DrawRoundStrokeCap(canvas, current, color, strokeWidth);
                } else if (lineJoin == SvgStrokeLineJoin.Miter) {
                    if (!DrawMiterStrokeJoin(canvas, previous, current, next, color, strokeWidth)) {
                        DrawBevelStrokeJoin(canvas, previous, current, next, color, strokeWidth);
                    }
                } else {
                    DrawBevelStrokeJoin(canvas, previous, current, next, color, strokeWidth);
                }
            }
        }

        private static void DrawRoundStrokeCap(OfficeRasterCanvas canvas, OfficePoint point, OfficeColor color, double strokeWidth) =>
            canvas.DrawEllipse(point.X, point.Y, strokeWidth / 2D, strokeWidth / 2D, color, OfficeColor.Transparent, 0D);

        private static bool DrawMiterStrokeJoin(OfficeRasterCanvas canvas, OfficePoint previous, OfficePoint current, OfficePoint next, OfficeColor color, double strokeWidth) {
            if (!TryCreateOuterJoin(previous, current, next, strokeWidth, out OfficePoint outerA, out OfficePoint outerB, out OfficePoint incomingUnit, out OfficePoint outgoingUnit)) {
                return false;
            }

            if (!TryIntersectLines(outerA, incomingUnit, outerB, outgoingUnit, out OfficePoint miter)) {
                return false;
            }

            double distance = Distance(current, miter);
            if (distance > strokeWidth * 4D) {
                return false;
            }

            canvas.FillPolygon(new[] { current, outerA, miter, outerB }, color);
            return true;
        }

        private static void DrawBevelStrokeJoin(OfficeRasterCanvas canvas, OfficePoint previous, OfficePoint current, OfficePoint next, OfficeColor color, double strokeWidth) {
            if (!TryCreateOuterJoin(previous, current, next, strokeWidth, out OfficePoint outerA, out OfficePoint outerB, out _, out _)) {
                return;
            }

            canvas.FillPolygon(new[] { current, outerA, outerB }, color);
        }

        private static bool TryCreateOuterJoin(
            OfficePoint previous,
            OfficePoint current,
            OfficePoint next,
            double strokeWidth,
            out OfficePoint outerA,
            out OfficePoint outerB,
            out OfficePoint incomingUnit,
            out OfficePoint outgoingUnit) {
            outerA = current;
            outerB = current;
            incomingUnit = current;
            outgoingUnit = current;
            double incomingLength = Distance(previous, current);
            double outgoingLength = Distance(current, next);
            if (incomingLength <= double.Epsilon || outgoingLength <= double.Epsilon) {
                return false;
            }

            incomingUnit = new OfficePoint((current.X - previous.X) / incomingLength, (current.Y - previous.Y) / incomingLength);
            outgoingUnit = new OfficePoint((next.X - current.X) / outgoingLength, (next.Y - current.Y) / outgoingLength);
            double turn = Cross(incomingUnit, outgoingUnit);
            if (Math.Abs(turn) <= 0.0001D) {
                return false;
            }

            double side = turn > 0D ? -1D : 1D;
            double half = strokeWidth / 2D;
            OfficePoint incomingNormal = new(-incomingUnit.Y * side, incomingUnit.X * side);
            OfficePoint outgoingNormal = new(-outgoingUnit.Y * side, outgoingUnit.X * side);
            outerA = new OfficePoint(current.X + (incomingNormal.X * half), current.Y + (incomingNormal.Y * half));
            outerB = new OfficePoint(current.X + (outgoingNormal.X * half), current.Y + (outgoingNormal.Y * half));
            return true;
        }

        private static bool TryIntersectLines(OfficePoint pointA, OfficePoint directionA, OfficePoint pointB, OfficePoint directionB, out OfficePoint intersection) {
            intersection = pointA;
            double denominator = Cross(directionA, directionB);
            if (Math.Abs(denominator) <= 0.0001D) {
                return false;
            }

            OfficePoint delta = new(pointB.X - pointA.X, pointB.Y - pointA.Y);
            double t = Cross(delta, directionB) / denominator;
            intersection = new OfficePoint(pointA.X + (directionA.X * t), pointA.Y + (directionA.Y * t));
            return true;
        }

        private static double Cross(OfficePoint a, OfficePoint b) =>
            (a.X * b.Y) - (a.Y * b.X);

        private static double Distance(OfficePoint a, OfficePoint b) {
            double dx = b.X - a.X;
            double dy = b.Y - a.Y;
            return Math.Sqrt((dx * dx) + (dy * dy));
        }

        private static void StrokeClosedContour(OfficeRasterCanvas canvas, IReadOnlyList<OfficePoint> points, SvgPaint paint, double strokeWidth, IReadOnlyList<double>? dashPattern) {
            if (dashPattern == null) {
                DrawJoinedPolyline(canvas, points, paint, strokeWidth, closed: true);
                return;
            }

            List<OfficePoint> closed = new(points.Count + 1);
            for (int i = 0; i < points.Count; i++) {
                closed.Add(points[i]);
            }

            closed.Add(points[0]);
            canvas.DrawPatternedPolyline(closed, GetStrokeFallbackColor(paint), strokeWidth, dashPattern);
        }

        private static IReadOnlyList<double>? ScaleDashPattern(IReadOnlyList<double>? pattern, double scale) {
            if (pattern == null || pattern.Count == 0) {
                return null;
            }

            List<double> scaled = new(pattern.Count);
            for (int i = 0; i < pattern.Count; i++) {
                double value = pattern[i] * Math.Max(0.0001D, scale);
                if (value > 0D && !double.IsNaN(value) && !double.IsInfinity(value)) {
                    scaled.Add(value);
                }
            }

            return scaled.Count == 0 ? null : scaled;
        }

        private static double GetStrokeScale(SvgPaint paint, SvgTransform transform) =>
            paint.NonScalingStroke ? 1D : transform.StrokeScale;

        private static IReadOnlyList<OfficePoint> CreateEllipsePoints(double cx, double cy, double rx, double ry, SvgTransform transform) {
            const int segments = 72;
            List<OfficePoint> points = new(segments);
            for (int i = 0; i < segments; i++) {
                double angle = (Math.PI * 2D * i) / segments;
                points.Add(transform.Apply(cx + (Math.Cos(angle) * rx), cy + (Math.Sin(angle) * ry)));
            }

            return points;
        }

        private static List<(double X, double Y)> CreateRoundedRectanglePoints(double x, double y, double width, double height, double rx, double ry) {
            const int quarterSegments = 8;
            List<(double X, double Y)> points = new(quarterSegments * 4);
            AddArcPoints(points, x + width - rx, y + ry, rx, ry, -Math.PI / 2D, 0D, quarterSegments);
            AddArcPoints(points, x + width - rx, y + height - ry, rx, ry, 0D, Math.PI / 2D, quarterSegments);
            AddArcPoints(points, x + rx, y + height - ry, rx, ry, Math.PI / 2D, Math.PI, quarterSegments);
            AddArcPoints(points, x + rx, y + ry, rx, ry, Math.PI, Math.PI * 3D / 2D, quarterSegments);
            return points;
        }

        private static void AddArcPoints(List<(double X, double Y)> points, double cx, double cy, double rx, double ry, double start, double end, int segments) {
            for (int i = 0; i <= segments; i++) {
                if (points.Count > 0 && i == 0) {
                    continue;
                }

                double t = i / (double)segments;
                double angle = start + ((end - start) * t);
                points.Add((cx + Math.Cos(angle) * rx, cy + Math.Sin(angle) * ry));
            }
        }

        private static List<OfficePoint> ProjectPoints(IReadOnlyList<(double X, double Y)> points, SvgTransform transform) {
            List<OfficePoint> projected = new(points.Count);
            for (int i = 0; i < points.Count; i++) {
                projected.Add(transform.Apply(points[i].X, points[i].Y));
            }

            return projected;
        }

        private static OfficeFillRule ResolveFillRule(XElement element, SvgRenderContext context) {
            Dictionary<string, string> style = context.StyleSheet.CreateStyle(element);
            string? fillRule = style.TryGetValue("fill-rule", out string? styleValue)
                ? styleValue
                : element.Attribute("fill-rule")?.Value;

            if (string.Equals(fillRule, "evenodd", StringComparison.OrdinalIgnoreCase)) {
                return OfficeFillRule.EvenOdd;
            }

            if (string.Equals(fillRule, "nonzero", StringComparison.OrdinalIgnoreCase)) {
                return OfficeFillRule.NonZero;
            }

            return context.CurrentFillRule;
        }

        private static string? ReadHref(XElement element) {
            foreach (XAttribute attribute in element.Attributes()) {
                if (string.Equals(attribute.Name.LocalName, "href", StringComparison.OrdinalIgnoreCase)) {
                    return attribute.Value;
                }
            }

            return null;
        }

        private static bool TryReadViewBoxTransform(XElement definition, XElement useElement, out SvgTransform transform) {
            transform = SvgTransform.Identity;
            if (!TryParseNumbers(definition.Attribute("viewBox")?.Value, out List<double> viewBox) ||
                viewBox.Count < 4 ||
                viewBox[2] <= 0D ||
                viewBox[3] <= 0D) {
                return false;
            }

            double width = ReadLength(useElement, "width", viewBox[2]);
            double height = ReadLength(useElement, "height", viewBox[3]);
            if (width <= 0D || height <= 0D) {
                return false;
            }

            transform = CreateViewBoxTransform(viewBox[0], viewBox[1], viewBox[2], viewBox[3], 0D, 0D, width, height, useElement.Attribute("preserveAspectRatio")?.Value ?? definition.Attribute("preserveAspectRatio")?.Value);
            return true;
        }

        private static SvgTransform CreateViewBoxTransform(double viewLeft, double viewTop, double viewWidth, double viewHeight, double viewportX, double viewportY, double viewportWidth, double viewportHeight, string? preserveAspectRatio) {
            string align = "xMidYMid";
            string meetOrSlice = "meet";
            if (!string.IsNullOrWhiteSpace(preserveAspectRatio)) {
                string[] parts = preserveAspectRatio!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                int partOffset = parts.Length > 0 && string.Equals(parts[0], "defer", StringComparison.OrdinalIgnoreCase) ? 1 : 0;

                if (parts.Length > partOffset) {
                    align = parts[partOffset];
                }

                if (parts.Length > partOffset + 1) {
                    meetOrSlice = parts[partOffset + 1];
                }
            }

            double scaleX = viewportWidth / viewWidth;
            double scaleY = viewportHeight / viewHeight;
            if (string.Equals(align, "none", StringComparison.OrdinalIgnoreCase)) {
                return SvgTransform.Create(scaleX, 0D, 0D, scaleY, viewportX - (viewLeft * scaleX), viewportY - (viewTop * scaleY));
            }

            double scale = string.Equals(meetOrSlice, "slice", StringComparison.OrdinalIgnoreCase)
                ? Math.Max(scaleX, scaleY)
                : Math.Min(scaleX, scaleY);
            double renderedWidth = viewWidth * scale;
            double renderedHeight = viewHeight * scale;
            double offsetX = align.IndexOf("xMax", StringComparison.OrdinalIgnoreCase) >= 0
                ? viewportWidth - renderedWidth
                : align.IndexOf("xMid", StringComparison.OrdinalIgnoreCase) >= 0
                    ? (viewportWidth - renderedWidth) / 2D
                    : 0D;
            double offsetY = align.IndexOf("YMax", StringComparison.OrdinalIgnoreCase) >= 0
                ? viewportHeight - renderedHeight
                : align.IndexOf("YMid", StringComparison.OrdinalIgnoreCase) >= 0
                    ? (viewportHeight - renderedHeight) / 2D
                    : 0D;

            return SvgTransform.Create(scale, 0D, 0D, scale, viewportX + offsetX - (viewLeft * scale), viewportY + offsetY - (viewTop * scale));
        }

        private static void ResolveViewport(XElement root, out double viewLeft, out double viewTop, out double viewWidth, out double viewHeight, out int width, out int height) {
            viewLeft = 0D;
            viewTop = 0D;
            viewWidth = ReadLength(root, "width", DefaultSize);
            viewHeight = ReadLength(root, "height", DefaultSize);
            if (TryParseNumbers(root.Attribute("viewBox")?.Value, out List<double> viewBox) && viewBox.Count >= 4 && viewBox[2] > 0D && viewBox[3] > 0D) {
                viewLeft = viewBox[0];
                viewTop = viewBox[1];
                viewWidth = viewBox[2];
                viewHeight = viewBox[3];
            }

            double rawWidth = ReadLength(root, "width", viewWidth);
            double rawHeight = ReadLength(root, "height", viewHeight);
            if (rawWidth <= 0D) {
                rawWidth = viewWidth;
            }

            if (rawHeight <= 0D) {
                rawHeight = viewHeight;
            }

            width = ClampSize((int)Math.Round(rawWidth));
            height = ClampSize((int)Math.Round(rawHeight));
        }

        private static int ClampSize(int value) => Math.Max(1, Math.Min(MaximumSize, value));

        private static double ReadLength(XElement element, string name, double fallback) =>
            TryParseLength(element.Attribute(name)?.Value, out double value) ? value : fallback;

        private static double ReadLength(XElement element, string name, double fallback, SvgRenderContext context, SvgLengthAxis axis) =>
            ReadLength(element, name, fallback, GetLengthReference(context, axis));

        private static double ReadLength(XElement element, string name, double fallback, double? percentageReference) =>
            TryParseLength(element.Attribute(name)?.Value, percentageReference, out double value) ? value : fallback;

        private static double? GetLengthReference(SvgRenderContext context, SvgLengthAxis axis) {
            SvgPaintBounds viewport = context.ViewportBounds;
            double width = viewport.Width;
            double height = viewport.Height;
            if (width <= 0D || height <= 0D) {
                return null;
            }

            return axis switch {
                SvgLengthAxis.X => width,
                SvgLengthAxis.Y => height,
                _ => Math.Sqrt((width * width) + (height * height)) / Math.Sqrt(2D)
            };
        }

        private static SvgTransform ReadTransform(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return SvgTransform.Identity;
            }

            SvgTransform transform = SvgTransform.Identity;
            int offset = 0;
            while (offset < value!.Length) {
                while (offset < value.Length && char.IsWhiteSpace(value[offset])) {
                    offset++;
                }

                int nameStart = offset;
                while (offset < value.Length && char.IsLetter(value[offset])) {
                    offset++;
                }

                string name = value.Substring(nameStart, offset - nameStart);
                if (string.IsNullOrEmpty(name) || offset >= value.Length || value[offset] != '(') {
                    break;
                }

                int close = value.IndexOf(')', offset + 1);
                if (close < 0) {
                    break;
                }

                if (TryParseNumbers(value.Substring(offset + 1, close - offset - 1), out List<double> numbers)) {
                    if (string.Equals(name, "translate", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(SvgTransform.Create(1D, 0D, 0D, 1D, numbers[0], numbers.Count > 1 ? numbers[1] : 0D));
                    } else if (string.Equals(name, "scale", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(SvgTransform.Create(numbers[0], 0D, 0D, numbers.Count > 1 ? numbers[1] : numbers[0], 0D, 0D));
                    } else if (string.Equals(name, "rotate", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(CreateRotationTransform(numbers));
                    } else if (string.Equals(name, "skewX", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(SvgTransform.Create(1D, 0D, Math.Tan(OfficeGeometry.DegreesToRadians(numbers[0])), 1D, 0D, 0D));
                    } else if (string.Equals(name, "skewY", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(SvgTransform.Create(1D, Math.Tan(OfficeGeometry.DegreesToRadians(numbers[0])), 0D, 1D, 0D, 0D));
                    } else if (string.Equals(name, "matrix", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 6) {
                        transform = transform.Multiply(SvgTransform.Create(numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5]));
                    }
                }

                offset = close + 1;
                while (offset < value.Length && (char.IsWhiteSpace(value[offset]) || value[offset] == ',')) {
                    offset++;
                }
            }

            return transform;
        }

        private static SvgTransform CreateRotationTransform(IReadOnlyList<double> numbers) {
            double radians = OfficeGeometry.DegreesToRadians(numbers[0]);
            double cos = Math.Cos(radians);
            double sin = Math.Sin(radians);
            SvgTransform rotate = SvgTransform.Create(cos, sin, -sin, cos, 0D, 0D);
            if (numbers.Count < 3) {
                return rotate;
            }

            double cx = numbers[1];
            double cy = numbers[2];
            return SvgTransform.Create(1D, 0D, 0D, 1D, cx, cy)
                .Multiply(rotate)
                .Multiply(SvgTransform.Create(1D, 0D, 0D, 1D, -cx, -cy));
        }

        private static bool TryParseLength(string? value, out double result) {
            return TryParseLength(value, null, out result);
        }

        private static bool TryParseLength(string? value, double? percentageReference, out double result) {
            result = 0D;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string trimmed = value!.Trim();
            if (trimmed.EndsWith("%", StringComparison.Ordinal)) {
                if (!percentageReference.HasValue || percentageReference.Value <= 0D) {
                    return false;
                }

                string rawPercentage = trimmed.Substring(0, trimmed.Length - 1);
                if (!double.TryParse(rawPercentage, NumberStyles.Float, CultureInfo.InvariantCulture, out double percentage)) {
                    return false;
                }

                result = percentageReference.Value * percentage / 100D;
                return true;
            }

            int end = 0;
            while (end < trimmed.Length && (char.IsDigit(trimmed[end]) || trimmed[end] == '-' || trimmed[end] == '+' || trimmed[end] == '.' || trimmed[end] == 'e' || trimmed[end] == 'E')) {
                end++;
            }

            if (end == 0 || !double.TryParse(trimmed.Substring(0, end), NumberStyles.Float, CultureInfo.InvariantCulture, out result)) {
                return false;
            }

            string unit = trimmed.Substring(end).Trim();
            switch (unit.ToLowerInvariant()) {
                case "":
                case "px":
                    return true;
                case "in":
                    result *= 96D;
                    return true;
                case "cm":
                    result *= 96D / 2.54D;
                    return true;
                case "mm":
                    result *= 96D / 25.4D;
                    return true;
                case "q":
                    result *= 96D / 101.6D;
                    return true;
                case "pt":
                    result *= 96D / 72D;
                    return true;
                case "pc":
                    result *= 16D;
                    return true;
                default:
                    result = 0D;
                    return false;
            }
        }

        private enum SvgLengthAxis {
            X,
            Y,
            Diagonal
        }

        private static bool TryParsePoints(string? value, out List<(double X, double Y)> points) {
            points = new List<(double X, double Y)>();
            if (!TryParseNumbers(value, out List<double> numbers) || numbers.Count < 4) {
                return false;
            }

            for (int i = 0; i + 1 < numbers.Count; i += 2) {
                points.Add((numbers[i], numbers[i + 1]));
            }

            return points.Count >= 2;
        }

        private static bool TryParseNumbers(string? value, out List<double> numbers) {
            numbers = new List<double>();
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            int index = 0;
            while (index < value!.Length) {
                while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ',')) {
                    index++;
                }

                int start = index;
                if (index < value.Length && (value[index] == '-' || value[index] == '+')) {
                    index++;
                }

                while (index < value.Length && (char.IsDigit(value[index]) || value[index] == '.')) {
                    index++;
                }

                if (index < value.Length && (value[index] == 'e' || value[index] == 'E')) {
                    index++;
                    if (index < value.Length && (value[index] == '-' || value[index] == '+')) {
                        index++;
                    }

                    while (index < value.Length && char.IsDigit(value[index])) {
                        index++;
                    }
                }

                if (index == start) {
                    break;
                }

                if (!double.TryParse(value.Substring(start, index - start), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                    return false;
                }

                numbers.Add(number);
            }

            return numbers.Count > 0;
        }

        private static bool TryParseRgbColor(string? value, out OfficeColor color) {
            color = OfficeColor.Black;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string trimmed = value!.Trim();
            if (!trimmed.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase) ||
                !trimmed.EndsWith(")", StringComparison.Ordinal)) {
                return false;
            }

            string inner = trimmed.Substring(4, trimmed.Length - 5);
            string[] components = inner.Split(new[] { ',', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (components.Length < 3) {
                return false;
            }

            if (!TryParseRgbComponent(components[0], out byte red) ||
                !TryParseRgbComponent(components[1], out byte green) ||
                !TryParseRgbComponent(components[2], out byte blue)) {
                return false;
            }

            color = OfficeColor.FromRgb(red, green, blue);
            return true;
        }

        private static bool TryParseRgbComponent(string raw, out byte component) {
            component = 0;
            string value = raw.Trim();
            bool percent = value.EndsWith("%", StringComparison.Ordinal);
            if (percent) {
                value = value.Substring(0, value.Length - 1);
            }

            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) {
                return false;
            }

            double scaled = percent ? parsed * 255D / 100D : parsed;
            component = (byte)Math.Max(0D, Math.Min(255D, Math.Round(scaled)));
            return true;
        }

        private static bool TryReadStyleValue(string? raw, string name, out string? value) {
            value = null;
            if (string.IsNullOrWhiteSpace(raw)) {
                return false;
            }

            string[] declarations = raw!.Split(';');
            for (int i = 0; i < declarations.Length; i++) {
                int separator = declarations[i].IndexOf(':');
                if (separator <= 0) {
                    continue;
                }

                if (string.Equals(declarations[i].Substring(0, separator).Trim(), name, StringComparison.OrdinalIgnoreCase)) {
                    value = declarations[i].Substring(separator + 1).Trim();
                    return true;
                }
            }

            return false;
        }

        private readonly struct SvgTransform {
            internal static SvgTransform Identity => Create(1D, 0D, 0D, 1D, 0D, 0D);

            private SvgTransform(double a, double b, double c, double d, double e, double f) {
                A = a;
                B = b;
                C = c;
                D = d;
                E = e;
                F = f;
            }

            internal double ScaleX => Math.Sqrt((A * A) + (B * B));

            internal double ScaleY => Math.Sqrt((C * C) + (D * D));

            internal double StrokeScale => Math.Max(0.0001D, (ScaleX + ScaleY) / 2D);

            internal double RotationDegrees => OfficeGeometry.RadiansToDegrees(Math.Atan2(B, A));

            private double A { get; }

            private double B { get; }

            private double C { get; }

            private double D { get; }

            private double E { get; }

            private double F { get; }

            internal static SvgTransform Create(double a, double b, double c, double d, double e, double f) => new(a, b, c, d, e, f);

            internal SvgTransform Multiply(SvgTransform other) =>
                new(
                    (A * other.A) + (C * other.B),
                    (B * other.A) + (D * other.B),
                    (A * other.C) + (C * other.D),
                    (B * other.C) + (D * other.D),
                    (A * other.E) + (C * other.F) + E,
                    (B * other.E) + (D * other.F) + F);

            internal OfficePoint Apply(double x, double y) => new((A * x) + (C * y) + E, (B * x) + (D * y) + F);

            internal OfficeTransform ToOfficeTransform() => new(A, B, C, D, E, F);
        }
    }
}
