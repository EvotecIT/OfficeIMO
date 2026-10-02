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
            cancellationToken.ThrowIfCancellationRequested();

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
            SvgRenderContext context = SvgRenderContext.Create(
                root,
                new SvgPaintBounds(viewLeft, viewTop, viewWidth, viewHeight),
                imageResolver,
                cancellationToken);
            if (context.StyleSheet.HasUnsupportedConditionalRules) {
                context.ReportUnsupportedFeature();
            }
            if (IsElementDisplayNone(root, context)) {
                return false;
            }
            double rootOpacity = SvgPaint.ReadOwnOpacity(root, context);
            if (rootOpacity <= 0D) {
                return false;
            }
            if (context.StyleSheet.HasActiveVisualEffect(root)) {
                context.ReportUnsupportedFeature();
            }

            bool useRootOpacityLayer = rootOpacity < 1D;
            SvgPaint inherited = SvgPaint.Resolve(root, SvgPaint.Default, context, applyOwnOpacity: !useRootOpacityLayer);
            SvgTransform transform = CreateViewBoxTransform(viewLeft, viewTop, viewWidth, viewHeight, 0D, 0D, width, height, root.Attribute("preserveAspectRatio")?.Value);
            using IDisposable rootTextStyle = context.PushTextStyle(SvgTextStyle.Resolve(root, SvgTextStyle.Default, context));
            using IDisposable rootFillRule = context.PushFillRule(ResolveFillRule(root, context));
            using IDisposable rootVisibilityScope = context.PushVisibility(ReadVisibilityOverride(root, context));
            OfficeRasterCanvas targetCanvas = canvas;
            OfficeRasterImage? rootLayer = null;
            if (useRootOpacityLayer) {
                rootLayer = new OfficeRasterImage(width, height, OfficeColor.Transparent);
                targetCanvas = CreateLayerCanvas(canvas, rootLayer);
            }

            using IDisposable rootPaintBoundsScope = context.PushPaintBounds(context.ViewportBounds);
            using IDisposable? rootClipScope = PushClipPath(targetCanvas, root, transform, context);
            bool rendered = RenderChildren(targetCanvas, root, inherited, transform, context);
            if (context.RenderBudgetExceeded) {
                AddSvgLossDiagnostic(
                    diagnosticSink,
                    diagnosticSource,
                    "The embedded SVG preview exceeded a bounded rendering depth, element, or selector-work budget and was omitted.",
                    OfficeConversionLossKind.Omission);
                return false;
            }
            if (context.UnsupportedFeatureCount > 0) {
                AddSvgLossDiagnostic(
                    diagnosticSink,
                    diagnosticSource,
                    context.UnsupportedFeatureCount.ToString(CultureInfo.InvariantCulture) +
                    " embedded SVG feature(s) were not represented completely by the dependency-free preview renderer.",
                    OfficeConversionLossKind.Approximation);
            }
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

        private static bool RenderElementWithinBudget(OfficeRasterCanvas canvas, XElement element, SvgPaint inherited, SvgTransform transform, SvgRenderContext context) {
            canvas.CancellationToken.ThrowIfCancellationRequested();
            string name = element.Name.LocalName;
            if ((!string.IsNullOrEmpty(element.Name.NamespaceName) &&
                 !string.Equals(element.Name.NamespaceName, "http://www.w3.org/2000/svg", StringComparison.Ordinal))) {
                context.ReportUnsupportedFeature();
            }
            if (string.Equals(name, "defs", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "style", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "title", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "desc", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "metadata", StringComparison.OrdinalIgnoreCase)) {
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
            bool hasActiveVisualEffect = context.StyleSheet.HasActiveVisualEffect(element);
            if (context.IsVisible && hasActiveVisualEffect) {
                context.ReportUnsupportedFeature();
            }

            if (useElementOpacityLayer) {
                OfficeRasterImage layer = new(canvas.Width, canvas.Height, OfficeColor.Transparent);
                OfficeRasterCanvas layerCanvas = CreateLayerCanvas(canvas, layer);
                bool rendered = RenderElementCore(layerCanvas, element, name, paint, localTransform, context);
                if (!rendered) {
                    return false;
                }
                if (!context.IsVisible && hasActiveVisualEffect) {
                    context.ReportUnsupportedFeature();
                }

                using IDisposable? groupClipScope = PushClipPath(canvas, element, localTransform, context);
                canvas.DrawImage(ApplyImageOpacity(layer, elementOpacity), 0D, 0D, canvas.Width, canvas.Height);
                return true;
            }

            using IDisposable? clipScope = PushClipPath(canvas, element, localTransform, context);
            bool renderedWithoutLayer = RenderElementCore(canvas, element, name, paint, localTransform, context);
            if (renderedWithoutLayer && !context.IsVisible && hasActiveVisualEffect) {
                context.ReportUnsupportedFeature();
            }
            return renderedWithoutLayer;
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
            IsGroupingElement(name) ||
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
            IsGroupingElement(name) ||
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
            if (IsGroupingElement(name)) {
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
                if (!TryParsePath(element.Attribute("d")?.Value, out List<SvgPathContour> contours, localTransform.CurveScale)) {
                    return false;
                }

                return RenderPath(canvas, element, contours, paint, localTransform, context);
            }

            context.ReportUnsupportedFeature();
            return RenderChildren(canvas, element, paint, localTransform, context);
        }

        private static bool IsGroupingElement(string name) =>
            string.Equals(name, "g", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(name, "a", StringComparison.OrdinalIgnoreCase);

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

            internal double CurveScale => Math.Sqrt(ScaleX * ScaleX + ScaleY * ScaleY);

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
