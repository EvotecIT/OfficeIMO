using System;
using System.Collections.Generic;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static IDisposable? PushClipPath(OfficeRasterCanvas canvas, XElement element, SvgTransform transform, SvgRenderContext context) {
            Dictionary<string, string> style = context.StyleSheet.CreateStyle(element);
            string? rawClip = style.TryGetValue("clip-path", out string? styleClip) ? styleClip : element.Attribute("clip-path")?.Value;
            if (!TryReadUrlId(rawClip, out string? id) ||
                id == null ||
                !context.TryGetDefinition(id, out XElement? definition) ||
                definition == null ||
                !string.Equals(definition.Name.LocalName, "clipPath", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            bool objectBoundingBox = string.Equals(definition.Attribute("clipPathUnits")?.Value?.Trim(), "objectBoundingBox", StringComparison.OrdinalIgnoreCase);
            SvgTransform clipTransform = transform;
            if (objectBoundingBox && context.CurrentPaintBounds is SvgPaintBounds paintBounds && paintBounds.HasArea) {
                clipTransform = transform.Multiply(SvgTransform.Create(paintBounds.Width, 0D, 0D, paintBounds.Height, paintBounds.Left, paintBounds.Top));
            }

            List<IReadOnlyList<OfficePoint>> contours = new();
            foreach (XElement child in definition.Elements()) {
                AddClipElementContours(child, clipTransform, contours, context, objectBoundingBox);
            }

            if (contours.Count == 0) {
                return null;
            }

            bool useEvenOddClip = UseEvenOddClip(definition, context.StyleSheet);
            return contours.Count == 1
                ? canvas.PushClipPolygon(contours[0])
                : useEvenOddClip
                    ? canvas.PushClipPolygonsEvenOdd(contours)
                    : canvas.PushClipPolygonsNonZero(contours);
        }

        private static void AddClipElementContours(XElement element, SvgTransform parentTransform, List<IReadOnlyList<OfficePoint>> contours, SvgRenderContext context, bool objectBoundingBox) {
            string name = element.Name.LocalName;
            SvgTransform transform = parentTransform.Multiply(ReadTransform(element.Attribute("transform")?.Value));
            if (string.Equals(name, "rect", StringComparison.OrdinalIgnoreCase)) {
                double x = ReadClipLength(element, "x", 0D, context, SvgLengthAxis.X, objectBoundingBox);
                double y = ReadClipLength(element, "y", 0D, context, SvgLengthAxis.Y, objectBoundingBox);
                double width = ReadClipLength(element, "width", 0D, context, SvgLengthAxis.X, objectBoundingBox);
                double height = ReadClipLength(element, "height", 0D, context, SvgLengthAxis.Y, objectBoundingBox);
                if (width > 0D && height > 0D) {
                    contours.Add(ProjectPoints(new[] {
                        (x, y),
                        (x + width, y),
                        (x + width, y + height),
                        (x, y + height)
                    }, transform));
                }

                return;
            }

            if (string.Equals(name, "circle", StringComparison.OrdinalIgnoreCase)) {
                double radius = ReadClipLength(element, "r", 0D, context, SvgLengthAxis.Diagonal, objectBoundingBox);
                if (radius > 0D) {
                    contours.Add(CreateEllipsePoints(ReadClipLength(element, "cx", 0D, context, SvgLengthAxis.X, objectBoundingBox), ReadClipLength(element, "cy", 0D, context, SvgLengthAxis.Y, objectBoundingBox), radius, radius, transform));
                }

                return;
            }

            if (string.Equals(name, "ellipse", StringComparison.OrdinalIgnoreCase)) {
                double rx = ReadClipLength(element, "rx", 0D, context, SvgLengthAxis.X, objectBoundingBox);
                double ry = ReadClipLength(element, "ry", 0D, context, SvgLengthAxis.Y, objectBoundingBox);
                if (rx > 0D && ry > 0D) {
                    contours.Add(CreateEllipsePoints(ReadClipLength(element, "cx", 0D, context, SvgLengthAxis.X, objectBoundingBox), ReadClipLength(element, "cy", 0D, context, SvgLengthAxis.Y, objectBoundingBox), rx, ry, transform));
                }

                return;
            }

            if (string.Equals(name, "polygon", StringComparison.OrdinalIgnoreCase)) {
                if (TryParsePoints(element.Attribute("points")?.Value, out List<(double X, double Y)> points) && points.Count >= 3) {
                    contours.Add(ProjectPoints(points, transform));
                }

                return;
            }

            if (string.Equals(name, "path", StringComparison.OrdinalIgnoreCase)) {
                if (TryParsePath(element.Attribute("d")?.Value, out List<SvgPathContour> pathContours)) {
                    for (int i = 0; i < pathContours.Count; i++) {
                        if (pathContours[i].Points.Count >= 3) {
                            contours.Add(ProjectPoints(pathContours[i].Points, transform));
                        }
                    }
                }

                return;
            }

            foreach (XElement child in element.Elements()) {
                AddClipElementContours(child, transform, contours, context, objectBoundingBox);
            }
        }

        private static double ReadClipLength(XElement element, string name, double fallback, SvgRenderContext context, SvgLengthAxis axis, bool objectBoundingBox) =>
            objectBoundingBox
                ? ReadLength(element, name, fallback, 1D)
                : ReadLength(element, name, fallback, context, axis);

        private static bool UseEvenOddClip(XElement element, SvgStyleSheet styleSheet) {
            Dictionary<string, string> style = styleSheet.CreateStyle(element);
            string? clipRule = style.TryGetValue("clip-rule", out string? clipValue)
                ? clipValue
                : element.Attribute("clip-rule")?.Value;
            if (string.IsNullOrWhiteSpace(clipRule)) {
                clipRule = style.TryGetValue("fill-rule", out string? fillValue)
                    ? fillValue
                    : element.Attribute("fill-rule")?.Value;
            }

            return string.Equals(clipRule, "evenodd", StringComparison.OrdinalIgnoreCase);
        }
    }
}
