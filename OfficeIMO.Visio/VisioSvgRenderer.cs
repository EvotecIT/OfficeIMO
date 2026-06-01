using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static class VisioSvgRenderer {
        private const string SvgNamespace = "http://www.w3.org/2000/svg";

        public static string Render(VisioPage page, VisioSvgSaveOptions options) {
            if (options.PixelsPerInch <= 0D || double.IsNaN(options.PixelsPerInch) || double.IsInfinity(options.PixelsPerInch)) {
                throw new ArgumentOutOfRangeException(nameof(options), "PixelsPerInch must be a finite positive number.");
            }

            double scale = options.PixelsPerInch;
            double width = Math.Max(page.Width, 0.01D) * scale;
            double height = Math.Max(page.Height, 0.01D) * scale;

            StringBuilder builder = new();
            XmlWriterSettings settings = new() {
                OmitXmlDeclaration = !options.IncludeXmlDeclaration,
                Indent = true
            };

            using (XmlWriter writer = XmlWriter.Create(new StringWriter(builder, CultureInfo.InvariantCulture), settings)) {
                writer.WriteStartDocument();
                writer.WriteStartElement("svg", SvgNamespace);
                writer.WriteAttributeString("width", Format(width));
                writer.WriteAttributeString("height", Format(height));
                writer.WriteAttributeString("viewBox", "0 0 " + Format(width) + " " + Format(height));
                writer.WriteAttributeString("role", "img");
                writer.WriteAttributeString("aria-label", string.IsNullOrWhiteSpace(page.Name) ? "OfficeIMO Visio page" : page.Name);

                WriteDefinitions(writer);
                if (options.BackgroundColor.HasValue && options.BackgroundColor.Value.A > 0) {
                    writer.WriteStartElement("rect", SvgNamespace);
                    writer.WriteAttributeString("x", "0");
                    writer.WriteAttributeString("y", "0");
                    writer.WriteAttributeString("width", Format(width));
                    writer.WriteAttributeString("height", Format(height));
                    WriteColor(writer, "fill", options.BackgroundColor.Value);
                    writer.WriteEndElement();
                }

                writer.WriteStartElement("g", SvgNamespace);
                writer.WriteAttributeString("data-officeimo-visio-page", page.Name);

                foreach (VisioShape shape in page.Shapes) {
                    WriteShape(writer, page, shape, options, scale);
                }

                foreach (VisioConnector connector in page.Connectors) {
                    WriteConnector(writer, page, connector, options, scale);
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }

            return builder.ToString();
        }

        private static void WriteDefinitions(XmlWriter writer) {
            writer.WriteStartElement("defs", SvgNamespace);
            writer.WriteStartElement("marker", SvgNamespace);
            writer.WriteAttributeString("id", "officeimo-visio-arrow");
            writer.WriteAttributeString("viewBox", "0 0 10 10");
            writer.WriteAttributeString("refX", "9");
            writer.WriteAttributeString("refY", "5");
            writer.WriteAttributeString("markerWidth", "8");
            writer.WriteAttributeString("markerHeight", "8");
            writer.WriteAttributeString("orient", "auto-start-reverse");
            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("d", "M 0 0 L 10 5 L 0 10 z");
            writer.WriteAttributeString("fill", "context-stroke");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteShape(XmlWriter writer, VisioPage page, VisioShape shape, VisioSvgSaveOptions options, double scale) {
            writer.WriteStartElement("g", SvgNamespace);
            writer.WriteAttributeString("data-visio-shape-id", shape.Id);
            if (!string.IsNullOrWhiteSpace(shape.NameU)) {
                writer.WriteAttributeString("data-visio-nameu", shape.NameU);
            }

            WriteShapeGeometry(writer, page, shape, scale);

            if (options.RenderText && !string.IsNullOrEmpty(shape.Text)) {
                WriteShapeText(writer, page, shape, scale);
            }

            foreach (VisioShape child in shape.Children) {
                WriteShape(writer, page, child, options, scale);
            }

            writer.WriteEndElement();
        }

        private static void WriteShapeGeometry(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            string kind = NormalizeKind(shape.MasterNameU ?? shape.NameU ?? shape.Name ?? string.Empty);
            if (kind == "ellipse" || kind == "circle") {
                (double centerX, double centerY) = GetPagePoint(shape, shape.LocPinX, shape.LocPinY);
                (double cx, double cy) = ToSvg(page, centerX, centerY, scale);
                writer.WriteStartElement("ellipse", SvgNamespace);
                writer.WriteAttributeString("cx", Format(cx));
                writer.WriteAttributeString("cy", Format(cy));
                writer.WriteAttributeString("rx", Format(Math.Abs(shape.Width * scale / 2D)));
                writer.WriteAttributeString("ry", Format(Math.Abs(shape.Height * scale / 2D)));
                if (Math.Abs(shape.Angle) > 1e-9) {
                    writer.WriteAttributeString("transform", "rotate(" + Format(RadiansToDegrees(-shape.Angle)) + " " + Format(cx) + " " + Format(cy) + ")");
                }

                WriteShapeStyle(writer, shape, scale);
                writer.WriteEndElement();
                return;
            }

            List<(double X, double Y)> points = GetShapePoints(shape, kind);
            if (points.Count == 0) {
                points = GetShapePoints(shape, "rectangle");
            }

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("d", BuildClosedPath(page, shape, points, scale));
            WriteShapeStyle(writer, shape, scale);
            writer.WriteEndElement();
        }

        private static List<(double X, double Y)> GetShapePoints(VisioShape shape, string kind) {
            double width = shape.Width;
            double height = shape.Height;
            double midX = width / 2D;
            double midY = height / 2D;
            switch (kind) {
                case "diamond":
                case "decision":
                    return new List<(double X, double Y)> { (midX, 0), (width, midY), (midX, height), (0, midY) };
                case "triangle":
                    return new List<(double X, double Y)> { (0, 0), (midX, height), (width, 0) };
                case "pentagon":
                case "offpagereference":
                    return new List<(double X, double Y)> {
                        (midX, height),
                        (width, height * 0.62D),
                        (width * 0.8D, 0),
                        (width * 0.2D, 0),
                        (0, height * 0.62D)
                    };
                case "parallelogram":
                case "data":
                    double offset = Math.Min(width / 4D, Math.Max(width / 10D, height / 3D));
                    return new List<(double X, double Y)> { (offset, 0), (width, 0), (width - offset, height), (0, height) };
                case "hexagon":
                case "preparation":
                    double inset = Math.Min(width / 4D, Math.Max(width / 8D, height / 4D));
                    return new List<(double X, double Y)> { (inset, 0), (width - inset, 0), (width, midY), (width - inset, height), (inset, height), (0, midY) };
                case "trapezoid":
                case "manualoperation":
                    double trapInset = Math.Min(width / 5D, Math.Max(width / 10D, height / 4D));
                    return new List<(double X, double Y)> { (trapInset, height), (width - trapInset, height), (width, 0), (0, 0) };
                default:
                    return new List<(double X, double Y)> { (0, 0), (width, 0), (width, height), (0, height) };
            }
        }

        private static void WriteShapeStyle(XmlWriter writer, VisioShape shape, double scale) {
            if (shape.FillPattern == 0 || shape.FillColor.A == 0) {
                writer.WriteAttributeString("fill", "none");
            } else {
                WriteColor(writer, "fill", shape.FillColor);
            }

            if (shape.LinePattern == 0 || shape.LineWeight <= 0D || shape.LineColor.A == 0) {
                writer.WriteAttributeString("stroke", "none");
            } else {
                WriteColor(writer, "stroke", shape.LineColor);
                writer.WriteAttributeString("stroke-width", Format(Math.Max(shape.LineWeight * scale, 0.75D)));
                if (shape.LinePattern != 1) {
                    writer.WriteAttributeString("stroke-dasharray", Format(6D) + " " + Format(4D));
                }
            }
        }

        private static void WriteShapeText(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            VisioTextStyle? style = shape.TextStyle;
            double localX = style?.TextPinX ?? shape.Width / 2D;
            double localY = style?.TextPinY ?? shape.Height / 2D;
            (double textX, double textY) = GetPagePoint(shape, localX, localY);
            (double x, double y) = ToSvg(page, textX, textY, scale);
            WriteText(
                writer,
                shape.Text!,
                x,
                y,
                style,
                defaultSize: 10D,
                scale: scale,
                rotateRadians: shape.Angle + (style?.TextAngle ?? 0D));
        }

        private static void WriteConnector(XmlWriter writer, VisioPage page, VisioConnector connector, VisioSvgSaveOptions options, double scale) {
            List<(double X, double Y)> points = GetConnectorPoints(connector);
            writer.WriteStartElement("g", SvgNamespace);
            writer.WriteAttributeString("data-visio-connector-id", connector.Id);

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("d", BuildOpenPath(page, points, scale));
            writer.WriteAttributeString("fill", "none");
            if (connector.LinePattern == 0 || connector.LineWeight <= 0D || connector.LineColor.A == 0) {
                writer.WriteAttributeString("stroke", "none");
            } else {
                WriteColor(writer, "stroke", connector.LineColor);
                writer.WriteAttributeString("stroke-width", Format(Math.Max(connector.LineWeight * scale, 0.75D)));
                writer.WriteAttributeString("stroke-linecap", "round");
                writer.WriteAttributeString("stroke-linejoin", "round");
                if (connector.LinePattern != 1) {
                    writer.WriteAttributeString("stroke-dasharray", Format(6D) + " " + Format(4D));
                }
            }

            if (connector.BeginArrow.HasValue && connector.BeginArrow.Value != EndArrow.None) {
                writer.WriteAttributeString("marker-start", "url(#officeimo-visio-arrow)");
            }

            if (connector.EndArrow.HasValue && connector.EndArrow.Value != EndArrow.None) {
                writer.WriteAttributeString("marker-end", "url(#officeimo-visio-arrow)");
            }

            writer.WriteEndElement();

            if (options.RenderConnectorLabels && !string.IsNullOrEmpty(connector.Label)) {
                (double labelX, double labelY) = ResolveConnectorLabelPoint(connector, points);
                (double x, double y) = ToSvg(page, labelX, labelY, scale);
                WriteText(writer, connector.Label!, x, y, connector.TextStyle, defaultSize: 9D, scale, rotateRadians: 0D);
            }

            writer.WriteEndElement();
        }

        private static List<(double X, double Y)> GetConnectorPoints(VisioConnector connector) {
            ComputeConnectorEndpoints(connector, out double startX, out double startY, out double endX, out double endY);
            List<(double X, double Y)> points = new() { (startX, startY) };
            if (connector.Waypoints.Count > 0) {
                foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                    points.Add((waypoint.X, waypoint.Y));
                }
            } else if (connector.Kind == ConnectorKind.RightAngle) {
                points.Add((startX, endY));
            }

            points.Add((endX, endY));
            return points;
        }

        private static void ComputeConnectorEndpoints(VisioConnector connector, out double startX, out double startY, out double endX, out double endY) {
            if (connector.FromConnectionPoint != null) {
                (startX, startY) = GetPagePoint(connector.From, connector.FromConnectionPoint.X, connector.FromConnectionPoint.Y);
            } else {
                (double fromLeft, double fromBottom, double fromRight, double fromTop) = GetPageBounds(connector.From);
                (double toLeft, double toBottom, double toRight, double toTop) = GetPageBounds(connector.To);
                ResolveFallbackEndpoint(fromLeft, fromBottom, fromRight, fromTop, toLeft, toBottom, toRight, toTop, out startX, out startY);
            }

            if (connector.ToConnectionPoint != null) {
                (endX, endY) = GetPagePoint(connector.To, connector.ToConnectionPoint.X, connector.ToConnectionPoint.Y);
            } else {
                (double toLeft, double toBottom, double toRight, double toTop) = GetPageBounds(connector.To);
                (double fromLeft, double fromBottom, double fromRight, double fromTop) = GetPageBounds(connector.From);
                ResolveFallbackEndpoint(toLeft, toBottom, toRight, toTop, fromLeft, fromBottom, fromRight, fromTop, out endX, out endY);
            }
        }

        private static (double X, double Y) ResolveConnectorLabelPoint(VisioConnector connector, IReadOnlyList<(double X, double Y)> points) {
            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
            if (placement?.AbsolutePinX.HasValue == true && placement.AbsolutePinY.HasValue) {
                return (placement.AbsolutePinX.Value, placement.AbsolutePinY.Value);
            }

            double position = VisioConnectorLabelPlacement.ClampPosition(placement?.Position ?? 0.5D);
            (double x, double y) = InterpolatePath(points, position);
            return (x + (placement?.OffsetX ?? 0D), y + (placement?.OffsetY ?? 0D));
        }

        private static (double X, double Y) InterpolatePath(IReadOnlyList<(double X, double Y)> points, double position) {
            if (points.Count == 0) return (0D, 0D);
            if (points.Count == 1) return points[0];

            double total = 0D;
            for (int i = 1; i < points.Count; i++) {
                total += Distance(points[i - 1], points[i]);
            }

            if (total <= 0D) return points[0];
            double target = total * position;
            double traversed = 0D;
            for (int i = 1; i < points.Count; i++) {
                double segment = Distance(points[i - 1], points[i]);
                if (traversed + segment >= target) {
                    double t = segment <= 0D ? 0D : (target - traversed) / segment;
                    return (
                        points[i - 1].X + ((points[i].X - points[i - 1].X) * t),
                        points[i - 1].Y + ((points[i].Y - points[i - 1].Y) * t));
                }

                traversed += segment;
            }

            return points[points.Count - 1];
        }

        private static void WriteText(XmlWriter writer, string text, double x, double y, VisioTextStyle? style, double defaultSize, double scale, double rotateRadians) {
            double fontSize = PointsToSvgPixels(style?.Size ?? defaultSize, scale);
            writer.WriteStartElement("text", SvgNamespace);
            writer.WriteAttributeString("x", Format(x));
            writer.WriteAttributeString("y", Format(y));
            writer.WriteAttributeString("font-family", string.IsNullOrWhiteSpace(style?.FontFamily) ? "Aptos, Calibri, Arial, sans-serif" : style!.FontFamily);
            writer.WriteAttributeString("font-size", Format(fontSize));
            writer.WriteAttributeString("text-anchor", GetTextAnchor(style));
            writer.WriteAttributeString("dominant-baseline", "middle");
            writer.WriteAttributeString("fill", style?.Color.HasValue == true ? "#" + style.Color.Value.ToRgbHex() : "#111827");
            if (style?.Bold == true) writer.WriteAttributeString("font-weight", "700");
            if (style?.Italic == true) writer.WriteAttributeString("font-style", "italic");
            if (style?.Underline == true) writer.WriteAttributeString("text-decoration", "underline");
            if (Math.Abs(rotateRadians) > 1e-9) {
                writer.WriteAttributeString("transform", "rotate(" + Format(RadiansToDegrees(-rotateRadians)) + " " + Format(x) + " " + Format(y) + ")");
            }

            string[] lines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            double lineHeight = fontSize * 1.2D;
            double startOffset = -((lines.Length - 1) * lineHeight) / 2D;
            for (int i = 0; i < lines.Length; i++) {
                writer.WriteStartElement("tspan", SvgNamespace);
                writer.WriteAttributeString("x", Format(x));
                writer.WriteAttributeString("dy", i == 0 ? Format(startOffset) : Format(lineHeight));
                writer.WriteString(lines[i]);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private static string GetTextAnchor(VisioTextStyle? style) {
            switch (style?.HorizontalAlignment) {
                case VisioTextHorizontalAlignment.Left:
                    return "start";
                case VisioTextHorizontalAlignment.Right:
                    return "end";
                default:
                    return "middle";
            }
        }

        private static string BuildClosedPath(VisioPage page, VisioShape shape, IReadOnlyList<(double X, double Y)> localPoints, double scale) {
            StringBuilder builder = new();
            for (int i = 0; i < localPoints.Count; i++) {
                (double absX, double absY) = GetPagePoint(shape, localPoints[i].X, localPoints[i].Y);
                (double x, double y) = ToSvg(page, absX, absY, scale);
                builder.Append(i == 0 ? "M " : " L ");
                builder.Append(Format(x)).Append(' ').Append(Format(y));
            }

            builder.Append(" Z");
            return builder.ToString();
        }

        private static (double X, double Y) GetPagePoint(VisioShape shape, double x, double y) {
            (double absX, double absY) = shape.GetAbsolutePoint(x, y);
            return shape.Parent != null
                ? GetPagePoint(shape.Parent, absX, absY)
                : (absX, absY);
        }

        private static (double Left, double Bottom, double Right, double Top) GetPageBounds(VisioShape shape) {
            (double x1, double y1) = GetPagePoint(shape, 0, 0);
            (double x2, double y2) = GetPagePoint(shape, shape.Width, 0);
            (double x3, double y3) = GetPagePoint(shape, 0, shape.Height);
            (double x4, double y4) = GetPagePoint(shape, shape.Width, shape.Height);
            double left = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            double right = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            double bottom = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            double top = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
            return (left, bottom, right, top);
        }

        private static void ResolveFallbackEndpoint(
            double sourceLeft,
            double sourceBottom,
            double sourceRight,
            double sourceTop,
            double targetLeft,
            double targetBottom,
            double targetRight,
            double targetTop,
            out double x,
            out double y) {
            double sourceCenterX = (sourceLeft + sourceRight) / 2D;
            double sourceCenterY = (sourceBottom + sourceTop) / 2D;
            double targetCenterX = (targetLeft + targetRight) / 2D;
            double targetCenterY = (targetBottom + targetTop) / 2D;
            double dx = targetCenterX - sourceCenterX;
            double dy = targetCenterY - sourceCenterY;

            if (Math.Abs(dy) > Math.Abs(dx)) {
                x = sourceCenterX;
                y = dy >= 0D ? sourceTop : sourceBottom;
                return;
            }

            x = dx >= 0D ? sourceRight : sourceLeft;
            y = sourceCenterY;
        }

        private static double PointsToSvgPixels(double points, double scale) {
            return points * scale / 72D;
        }

        private static string BuildOpenPath(VisioPage page, IReadOnlyList<(double X, double Y)> points, double scale) {
            StringBuilder builder = new();
            for (int i = 0; i < points.Count; i++) {
                (double x, double y) = ToSvg(page, points[i].X, points[i].Y, scale);
                builder.Append(i == 0 ? "M " : " L ");
                builder.Append(Format(x)).Append(' ').Append(Format(y));
            }

            return builder.ToString();
        }

        private static (double X, double Y) ToSvg(VisioPage page, double x, double y, double scale) {
            return (x * scale, (page.Height - y) * scale);
        }

        private static void WriteColor(XmlWriter writer, string attributeName, Color color) {
            writer.WriteAttributeString(attributeName, "#" + color.ToRgbHex());
            if (color.A < 255) {
                writer.WriteAttributeString(attributeName + "-opacity", Format(color.A / 255D));
            }
        }

        private static string NormalizeKind(string value) {
            StringBuilder builder = new();
            foreach (char c in value) {
                if (char.IsLetterOrDigit(c)) {
                    builder.Append(char.ToLowerInvariant(c));
                }
            }

            return builder.ToString();
        }

        private static double Distance((double X, double Y) a, (double X, double Y) b) {
            double dx = b.X - a.X;
            double dy = b.Y - a.Y;
            return Math.Sqrt((dx * dx) + (dy * dy));
        }

        private static double RadiansToDegrees(double radians) => radians * 180D / Math.PI;

        private static string Format(double value) {
            if (Math.Abs(value) < 0.0000001D) value = 0D;
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }
    }
}
