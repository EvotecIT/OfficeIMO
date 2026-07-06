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

            List<IReadOnlyList<OfficePoint>> contours = new();
            foreach (XElement child in definition.Elements()) {
                AddClipElementContours(child, transform, contours);
            }

            if (contours.Count == 0) {
                return null;
            }

            return contours.Count == 1 ? canvas.PushClipPolygon(contours[0]) : canvas.PushClipPolygonsNonZero(contours);
        }

        private static void AddClipElementContours(XElement element, SvgTransform parentTransform, List<IReadOnlyList<OfficePoint>> contours) {
            string name = element.Name.LocalName;
            SvgTransform transform = parentTransform.Multiply(ReadTransform(element.Attribute("transform")?.Value));
            if (string.Equals(name, "rect", StringComparison.OrdinalIgnoreCase)) {
                double x = ReadLength(element, "x", 0D);
                double y = ReadLength(element, "y", 0D);
                double width = ReadLength(element, "width", 0D);
                double height = ReadLength(element, "height", 0D);
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
                double radius = ReadLength(element, "r", 0D);
                if (radius > 0D) {
                    contours.Add(CreateEllipsePoints(ReadLength(element, "cx", 0D), ReadLength(element, "cy", 0D), radius, radius, transform));
                }

                return;
            }

            if (string.Equals(name, "ellipse", StringComparison.OrdinalIgnoreCase)) {
                double rx = ReadLength(element, "rx", 0D);
                double ry = ReadLength(element, "ry", 0D);
                if (rx > 0D && ry > 0D) {
                    contours.Add(CreateEllipsePoints(ReadLength(element, "cx", 0D), ReadLength(element, "cy", 0D), rx, ry, transform));
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
                        if (pathContours[i].IsClosed && pathContours[i].Points.Count >= 3) {
                            contours.Add(ProjectPoints(pathContours[i].Points, transform));
                        }
                    }
                }

                return;
            }

            foreach (XElement child in element.Elements()) {
                AddClipElementContours(child, transform, contours);
            }
        }
    }
}
