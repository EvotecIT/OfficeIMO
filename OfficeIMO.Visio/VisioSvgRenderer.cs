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

                VisioRenderLabelLayout? labelLayout = options.ResolveConnectorLabelOverlaps
                    ? VisioRenderLabelLayout.Create(page)
                    : null;
                foreach (VisioConnector connector in page.Connectors) {
                    WriteConnector(writer, page, connector, options, scale, labelLayout);
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }

            return builder.ToString();
        }

        private static void WriteShape(XmlWriter writer, VisioPage page, VisioShape shape, VisioSvgSaveOptions options, double scale) {
            writer.WriteStartElement("g", SvgNamespace);
            writer.WriteAttributeString("data-visio-shape-id", shape.Id);
            if (!string.IsNullOrWhiteSpace(shape.NameU)) {
                writer.WriteAttributeString("data-visio-nameu", shape.NameU);
            }

            WriteShapeGeometry(writer, page, shape, scale);

            if (options.RenderStencilArtwork) {
                if (!WritePackagePreviewArtwork(writer, page, shape, scale)) {
                    WriteStencilArtwork(writer, page, shape, scale);
                }
            }

            if (options.RenderText && !string.IsNullOrEmpty(shape.Text)) {
                WriteShapeText(writer, page, shape, scale);
            }

            foreach (VisioShape child in shape.Children) {
                WriteShape(writer, page, child, options, scale);
            }

            writer.WriteEndElement();
        }

        private static void WriteShapeGeometry(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            string kind = VisioShapeGeometry.ResolveRenderKind(shape);
            if (VisioShapeGeometry.TryGetRenderClosedPaths(shape, out List<VisioShapeGeometryPath> preservedPaths)) {
                foreach (VisioShapeGeometryPath preservedPath in preservedPaths) {
                    writer.WriteStartElement("path", SvgNamespace);
                    writer.WriteAttributeString("d", BuildPath(page, shape, preservedPath.Points, scale, preservedPath.IsClosed));
                    writer.WriteAttributeString("data-officeimo-preserved-geometry", "true");
                    WriteShapeStyle(writer, shape, scale, preservedPath.NoFill || !preservedPath.IsClosed, preservedPath.NoLine);
                    writer.WriteEndElement();
                }

                return;
            }

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

            if (kind == "database") {
                WriteDatabaseGeometry(writer, page, shape, scale);
                return;
            }

            List<(double X, double Y)> points = VisioShapeGeometry.GetBuiltinClosedPath(shape, kind);

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("d", BuildPath(page, shape, points, scale, isClosed: true));
            WriteShapeStyle(writer, shape, scale);
            writer.WriteEndElement();
        }

        private static void WriteDatabaseGeometry(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            (double centerXPage, double centerYPage) = GetPagePoint(shape, shape.LocPinX, shape.LocPinY);
            (double centerX, double centerY) = ToSvg(page, centerXPage, centerYPage, scale);
            double width = Math.Max(0.01D, shape.Width * scale);
            double height = Math.Max(0.01D, shape.Height * scale);
            double capHeight = Math.Min(height * 0.18D, width * 0.16D);
            double left = centerX - (width / 2D);
            double right = centerX + (width / 2D);
            double top = centerY - (height / 2D);
            double bottom = centerY + (height / 2D);
            string transform = Math.Abs(shape.Angle) > 1e-9
                ? FormatTextRotation(shape.Angle, centerX, centerY)
                : string.Empty;

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-database-geometry", "true");
            writer.WriteAttributeString(
                "d",
                "M " + Format(left) + " " + Format(top + capHeight) +
                " C " + Format(left) + " " + Format(top - (capHeight * 0.35D)) +
                " " + Format(right) + " " + Format(top - (capHeight * 0.35D)) +
                " " + Format(right) + " " + Format(top + capHeight) +
                " L " + Format(right) + " " + Format(bottom - capHeight) +
                " C " + Format(right) + " " + Format(bottom + (capHeight * 0.35D)) +
                " " + Format(left) + " " + Format(bottom + (capHeight * 0.35D)) +
                " " + Format(left) + " " + Format(bottom - capHeight) +
                " Z");
            if (!string.IsNullOrEmpty(transform)) {
                writer.WriteAttributeString("transform", transform);
            }

            WriteShapeStyle(writer, shape, scale);
            writer.WriteEndElement();

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-database-seam", "true");
            writer.WriteAttributeString(
                "d",
                "M " + Format(left) + " " + Format(top + capHeight) +
                " C " + Format(left) + " " + Format(top + (capHeight * 2.35D)) +
                " " + Format(right) + " " + Format(top + (capHeight * 2.35D)) +
                " " + Format(right) + " " + Format(top + capHeight));
            if (!string.IsNullOrEmpty(transform)) {
                writer.WriteAttributeString("transform", transform);
            }

            WriteShapeStyle(writer, shape, scale, noFill: true);
            writer.WriteEndElement();
        }

        private static void WriteStencilArtwork(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            string? stencilKey = VisioStencilArtwork.GetKey(shape);
            if (string.IsNullOrEmpty(stencilKey)) {
                return;
            }

            double placementScale = string.IsNullOrWhiteSpace(shape.Text) ? 0.58D : 0.34D;
            double iconSize = Math.Max(0.08D, Math.Min(shape.Width, shape.Height) * placementScale);
            double localCx = shape.Width / 2D;
            double localCy = string.IsNullOrWhiteSpace(shape.Text)
                ? shape.Height / 2D
                : shape.Height - Math.Min(shape.Height * 0.28D, iconSize * 0.72D);
            (double cx, double cy) = GetPagePoint(shape, localCx, localCy);
            (double x, double y) = ToSvg(page, cx, cy, scale);
            double size = iconSize * scale;
            Color color = VisioStencilArtwork.ResolveColor(shape, 210);

            writer.WriteStartElement("g", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-stencil-artwork", "true");
            writer.WriteAttributeString("data-officeimo-stencil-key", stencilKey);
            writer.WriteAttributeString("opacity", "0.42");
            if (Math.Abs(shape.Angle) > 1e-9) {
                writer.WriteAttributeString("transform", FormatTextRotation(shape.Angle, x, y));
            }

            switch (stencilKey) {
                case "person":
                    WriteSvgCircle(writer, x, y - size * 0.18D, size * 0.16D, color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
                    WriteSvgPath(writer, "M " + Format(x - size * 0.27D) + " " + Format(y + size * 0.29D) +
                                         " Q " + Format(x) + " " + Format(y + size * 0.02D) +
                                         " " + Format(x + size * 0.27D) + " " + Format(y + size * 0.29D), color, fill: false, strokeWidth: Math.Max(1D, size * 0.055D));
                    break;
                case "data":
                    WriteSvgCylinder(writer, x, y, size, color);
                    break;
                case "security":
                    WriteSvgShield(writer, x, y, size, color);
                    break;
                case "compute":
                    WriteSvgRect(writer, x - size * 0.34D, y - size * 0.24D, size * 0.68D, size * 0.48D, color, fill: false, strokeWidth: Math.Max(1D, size * 0.045D));
                    WriteSvgLine(writer, x - size * 0.22D, y - size * 0.06D, x + size * 0.22D, y - size * 0.06D, color, Math.Max(1D, size * 0.04D));
                    WriteSvgLine(writer, x - size * 0.22D, y + size * 0.08D, x + size * 0.22D, y + size * 0.08D, color, Math.Max(1D, size * 0.04D));
                    break;
                case "cloud":
                    WriteSvgPath(writer, BuildCloudPath(x, y, size), color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
                    break;
                case "container":
                    WriteSvgHex(writer, x, y, size, color);
                    break;
                case "event":
                    WriteSvgLine(writer, x - size * 0.32D, y - size * 0.16D, x + size * 0.28D, y - size * 0.16D, color, Math.Max(1D, size * 0.045D));
                    WriteSvgLine(writer, x - size * 0.32D, y, x + size * 0.18D, y, color, Math.Max(1D, size * 0.045D));
                    WriteSvgLine(writer, x - size * 0.32D, y + size * 0.16D, x + size * 0.28D, y + size * 0.16D, color, Math.Max(1D, size * 0.045D));
                    break;
                case "monitoring":
                    WriteSvgPath(writer, "M " + Format(x - size * 0.36D) + " " + Format(y) +
                                         " L " + Format(x - size * 0.14D) + " " + Format(y) +
                                         " L " + Format(x - size * 0.04D) + " " + Format(y - size * 0.22D) +
                                         " L " + Format(x + size * 0.09D) + " " + Format(y + size * 0.2D) +
                                         " L " + Format(x + size * 0.19D) + " " + Format(y) +
                                         " L " + Format(x + size * 0.36D) + " " + Format(y), color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
                    break;
            }

            writer.WriteEndElement();
        }

        private static bool WritePackagePreviewArtwork(XmlWriter writer, VisioPage page, VisioShape shape, double scale) {
            if (!VisioPackagePreviewArtwork.TryGetBrowserImage(shape, out VisioPreviewImage image)) {
                return false;
            }

            double placementScale = string.IsNullOrWhiteSpace(shape.Text) ? 0.64D : 0.42D;
            double imageWidth = Math.Max(0.01D, shape.Width * placementScale);
            double imageHeight = Math.Max(0.01D, shape.Height * placementScale);
            double localCx = shape.Width / 2D;
            double localCy = string.IsNullOrWhiteSpace(shape.Text)
                ? shape.Height / 2D
                : shape.Height - Math.Min(shape.Height * 0.3D, imageHeight * 0.72D);
            (double cx, double cy) = GetPagePoint(shape, localCx, localCy);
            (double centerX, double centerY) = ToSvg(page, cx, cy, scale);
            double width = imageWidth * scale;
            double height = imageHeight * scale;
            double x = centerX - (width / 2D);
            double y = centerY - (height / 2D);

            writer.WriteStartElement("image", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-package-preview-artwork", "true");
            writer.WriteAttributeString("x", Format(x));
            writer.WriteAttributeString("y", Format(y));
            writer.WriteAttributeString("width", Format(width));
            writer.WriteAttributeString("height", Format(height));
            writer.WriteAttributeString("preserveAspectRatio", "xMidYMid meet");
            if (Math.Abs(shape.Angle) > 1e-9) {
                writer.WriteAttributeString("transform", FormatTextRotation(shape.Angle, centerX, centerY));
            }

            writer.WriteAttributeString("href", "data:" + image.ContentType + ";base64," + Convert.ToBase64String(image.Data));
            writer.WriteEndElement();
            return true;
        }

        private static void WriteShapeStyle(XmlWriter writer, VisioShape shape, double scale, bool noFill = false, bool noLine = false) {
            if (noFill || shape.FillPattern == 0 || shape.FillColor.A == 0) {
                writer.WriteAttributeString("fill", "none");
            } else {
                WriteColor(writer, "fill", shape.FillColor);
            }

            if (noLine || shape.LinePattern == 0 || shape.LineWeight <= 0D || shape.LineColor.A == 0) {
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
            double textWidth = Math.Max(0.05D, style?.TextWidth ?? shape.Width);
            double textHeight = Math.Max(0.05D, style?.TextHeight ?? shape.Height);
            double horizontalMargins = (style?.LeftMargin ?? 0.05D) + (style?.RightMargin ?? 0.05D);
            double verticalMargins = (style?.TopMargin ?? 0.03D) + (style?.BottomMargin ?? 0.03D);
            WriteText(
                writer,
                shape.Text!,
                x,
                y,
                style,
                defaultSize: 10D,
                scale: scale,
                rotateRadians: shape.Angle + (style?.TextAngle ?? 0D),
                maxWidth: Math.Max(12D, (textWidth - horizontalMargins) * scale),
                maxHeight: Math.Max(8D, (textHeight - verticalMargins) * scale),
                drawLabelBackground: false);
        }

        private static void WriteConnector(XmlWriter writer, VisioPage page, VisioConnector connector, VisioSvgSaveOptions options, double scale, VisioRenderLabelLayout? labelLayout) {
            List<(double X, double Y)> points = GetConnectorPoints(connector);
            writer.WriteStartElement("g", SvgNamespace);
            writer.WriteAttributeString("data-visio-connector-id", connector.Id);

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("d", BuildOpenPath(page, points, scale));
            writer.WriteAttributeString("fill", "none");
            bool visibleLine = connector.LinePattern != 0 && connector.LineWeight > 0D && connector.LineColor.A > 0;
            double strokeWidth = Math.Max(connector.LineWeight * scale, 0.75D);
            if (!visibleLine) {
                writer.WriteAttributeString("stroke", "none");
            } else {
                WriteColor(writer, "stroke", connector.LineColor);
                writer.WriteAttributeString("stroke-width", Format(strokeWidth));
                writer.WriteAttributeString("stroke-linecap", "round");
                writer.WriteAttributeString("stroke-linejoin", "round");
                if (connector.LinePattern != 1) {
                    writer.WriteAttributeString("stroke-dasharray", Format(6D) + " " + Format(4D));
                }
            }

            writer.WriteEndElement();

            if (visibleLine) {
                if (connector.BeginArrow.HasValue && connector.BeginArrow.Value != EndArrow.None && TryGetArrowSegment(points, fromStart: true, out (double X, double Y) beginTip, out (double X, double Y) beginFrom)) {
                    WriteArrow(writer, page, beginTip, beginFrom, scale, connector.LineColor, strokeWidth, "start");
                }

                if (connector.EndArrow.HasValue && connector.EndArrow.Value != EndArrow.None && TryGetArrowSegment(points, fromStart: false, out (double X, double Y) endTip, out (double X, double Y) endFrom)) {
                    WriteArrow(writer, page, endTip, endFrom, scale, connector.LineColor, strokeWidth, "end");
                }
            }

            if (options.RenderConnectorLabels && !string.IsNullOrEmpty(connector.Label)) {
                VisioRenderConnectorLabelPlacement label = labelLayout?.Resolve(connector, points) ?? ResolveConnectorLabel(connector, points);
                (double x, double y) = ToSvg(page, label.X, label.Y, scale);
                double maxWidth = label.Width * scale;
                double maxHeight = label.Height * scale;
                WriteText(
                    writer,
                    connector.Label!,
                    x,
                    y,
                    connector.TextStyle,
                    defaultSize: 9D,
                    scale,
                    rotateRadians: 0D,
                    maxWidth,
                    maxHeight,
                    drawLabelBackground: true,
                    labelAdjusted: label.Adjusted);
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

        private static VisioRenderConnectorLabelPlacement ResolveConnectorLabel(VisioConnector connector, IReadOnlyList<(double X, double Y)> points) {
            (double x, double y) = ResolveConnectorLabelPoint(connector, points);
            VisioConnectorLabelPlacement? placement = connector.LabelPlacement;
            double width = Math.Max(0.6D, connector.TextStyle?.TextWidth ?? placement?.Width ?? 1.35D);
            double height = Math.Max(0.18D, connector.TextStyle?.TextHeight ?? placement?.Height ?? 0.34D);
            return new VisioRenderConnectorLabelPlacement(x, y, width, height, adjusted: false);
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

        private static void WriteText(
            XmlWriter writer,
            string text,
            double x,
            double y,
            VisioTextStyle? style,
            double defaultSize,
            double scale,
            double rotateRadians,
            double maxWidth = 0D,
            double maxHeight = 0D,
            bool drawLabelBackground = false,
            bool labelAdjusted = false) {
            double fontSize = PointsToSvgPixels(style?.Size ?? defaultSize, scale);
            double availableWidth = IsFinitePositive(maxWidth) ? maxWidth : double.PositiveInfinity;
            double availableHeight = IsFinitePositive(maxHeight) ? maxHeight : double.PositiveInfinity;
            TextLayout layout = CreateTextLayout(text, fontSize, availableWidth, availableHeight);
            fontSize = layout.FontSize;
            double anchorX = ResolveTextAnchorX(x, availableWidth, style?.HorizontalAlignment);
            double top = ResolveTextTop(y, layout.Height, availableHeight, style?.VerticalAlignment);

            Color? backgroundColor = ResolveTextBackground(style, drawLabelBackground);
            if (backgroundColor.HasValue) {
                double padX = Math.Max(3D, fontSize * 0.22D);
                double padY = Math.Max(2D, fontSize * 0.16D);
                double backgroundLeft = GetAlignedTextLeft(anchorX, layout.Width, style?.HorizontalAlignment) - padX;
                writer.WriteStartElement("rect", SvgNamespace);
                writer.WriteAttributeString("data-officeimo-text-background", "true");
                if (drawLabelBackground) {
                    writer.WriteAttributeString("data-officeimo-connector-label-background", "true");
                }

                if (labelAdjusted) {
                    writer.WriteAttributeString("data-officeimo-label-adjusted", "true");
                }

                writer.WriteAttributeString("x", Format(backgroundLeft));
                writer.WriteAttributeString("y", Format(top - padY));
                writer.WriteAttributeString("width", Format(layout.Width + (padX * 2D)));
                writer.WriteAttributeString("height", Format(layout.Height + (padY * 2D)));
                if (Math.Abs(rotateRadians) > 1e-9) {
                    writer.WriteAttributeString("transform", FormatTextRotation(rotateRadians, x, y));
                }

                WriteColor(writer, "fill", backgroundColor.Value);
                writer.WriteEndElement();
            }

            writer.WriteStartElement("text", SvgNamespace);
            if (labelAdjusted) {
                writer.WriteAttributeString("data-officeimo-label-adjusted", "true");
            }

            writer.WriteAttributeString("x", Format(anchorX));
            writer.WriteAttributeString("y", Format(top + (fontSize / 2D)));
            writer.WriteAttributeString("font-family", string.IsNullOrWhiteSpace(style?.FontFamily) ? "Aptos, Calibri, Arial, sans-serif" : style!.FontFamily);
            writer.WriteAttributeString("font-size", Format(fontSize));
            writer.WriteAttributeString("text-anchor", GetTextAnchor(style));
            writer.WriteAttributeString("dominant-baseline", "middle");
            WriteColor(writer, "fill", style?.Color ?? Color.FromRgb(17, 24, 39));
            if (style?.Bold == true) writer.WriteAttributeString("font-weight", "700");
            if (style?.Italic == true) writer.WriteAttributeString("font-style", "italic");
            if (style?.Underline == true) writer.WriteAttributeString("text-decoration", "underline");
            if (Math.Abs(rotateRadians) > 1e-9) {
                writer.WriteAttributeString("transform", FormatTextRotation(rotateRadians, x, y));
            }

            for (int i = 0; i < layout.Lines.Length; i++) {
                writer.WriteStartElement("tspan", SvgNamespace);
                writer.WriteAttributeString("x", Format(anchorX));
                writer.WriteAttributeString("dy", i == 0 ? "0" : Format(layout.LineHeight));
                writer.WriteString(layout.Lines[i]);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private static void WriteArrow(
            XmlWriter writer,
            VisioPage page,
            (double X, double Y) tip,
            (double X, double Y) from,
            double scale,
            Color color,
            double strokeWidth,
            string position) {
            (double tipX, double tipY) = ToSvg(page, tip.X, tip.Y, scale);
            (double fromX, double fromY) = ToSvg(page, from.X, from.Y, scale);
            double angle = Math.Atan2(tipY - fromY, tipX - fromX);
            double length = Math.Max(strokeWidth * 4D, 8D);
            double wing = Math.PI / 7D;
            double x1 = tipX - (Math.Cos(angle - wing) * length);
            double y1 = tipY - (Math.Sin(angle - wing) * length);
            double x2 = tipX - (Math.Cos(angle + wing) * length);
            double y2 = tipY - (Math.Sin(angle + wing) * length);

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("data-officeimo-connector-arrow", position);
            writer.WriteAttributeString("d", "M " + Format(tipX) + " " + Format(tipY) +
                                             " L " + Format(x1) + " " + Format(y1) +
                                             " L " + Format(x2) + " " + Format(y2) + " Z");
            WriteColor(writer, "fill", color);
            writer.WriteAttributeString("stroke", "none");
            writer.WriteEndElement();
        }

        private static bool TryGetArrowSegment(
            IReadOnlyList<(double X, double Y)> points,
            bool fromStart,
            out (double X, double Y) tip,
            out (double X, double Y) from) {
            if (points.Count < 2) {
                tip = default;
                from = default;
                return false;
            }

            if (fromStart) {
                tip = points[0];
                for (int i = 1; i < points.Count; i++) {
                    if (Distance(tip, points[i]) > 1e-6D) {
                        from = points[i];
                        return true;
                    }
                }
            } else {
                tip = points[points.Count - 1];
                for (int i = points.Count - 2; i >= 0; i--) {
                    if (Distance(tip, points[i]) > 1e-6D) {
                        from = points[i];
                        return true;
                    }
                }
            }

            from = default;
            return false;
        }

        private static TextLayout CreateTextLayout(string text, double fontSize, double maxWidth, double maxHeight) {
            string[] lines = WrapText(text, fontSize, maxWidth);
            double lineHeight = fontSize * 1.2D;
            double measuredWidth = MeasureMaxLineWidth(lines, fontSize);
            double measuredHeight = Math.Max(fontSize, ((lines.Length - 1) * lineHeight) + fontSize);
            double scaleDown = Math.Min(1D, Math.Min(maxWidth / Math.Max(measuredWidth, 1D), maxHeight / Math.Max(measuredHeight, 1D)));
            if (scaleDown < 0.98D) {
                fontSize = Math.Max(5D, fontSize * scaleDown);
                lines = WrapText(text, fontSize, maxWidth);
                lineHeight = fontSize * 1.2D;
                measuredWidth = MeasureMaxLineWidth(lines, fontSize);
                measuredHeight = Math.Max(fontSize, ((lines.Length - 1) * lineHeight) + fontSize);
            }

            return new TextLayout(lines, fontSize, lineHeight, measuredWidth, measuredHeight);
        }

        private static string[] WrapText(string text, double fontSize, double maxWidth) {
            string[] sourceLines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            List<string> output = new();
            foreach (string sourceLine in sourceLines) {
                string line = sourceLine.Trim();
                if (line.Length == 0) {
                    output.Add(string.Empty);
                    continue;
                }

                string[] words = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                string current = string.Empty;
                for (int i = 0; i < words.Length; i++) {
                    string word = words[i];
                    if (EstimateTextWidth(word, fontSize) > maxWidth) {
                        if (current.Length > 0) {
                            output.Add(current);
                            current = string.Empty;
                        }

                        foreach (string part in BreakWord(word, fontSize, maxWidth)) {
                            output.Add(part);
                        }

                        continue;
                    }

                    string candidate = current.Length == 0 ? word : current + " " + word;
                    if (current.Length > 0 && EstimateTextWidth(candidate, fontSize) > maxWidth) {
                        output.Add(current);
                        current = word;
                    } else {
                        current = candidate;
                    }
                }

                if (current.Length > 0) {
                    output.Add(current);
                }
            }

            return output.Count == 0 ? new[] { string.Empty } : output.ToArray();
        }

        private static IEnumerable<string> BreakWord(string word, double fontSize, double maxWidth) {
            StringBuilder part = new();
            foreach (char c in word) {
                string candidate = part.ToString() + c;
                if (part.Length > 0 && EstimateTextWidth(candidate, fontSize) > maxWidth) {
                    yield return part.ToString();
                    part.Clear();
                }

                part.Append(c);
            }

            if (part.Length > 0) {
                yield return part.ToString();
            }
        }

        private static double MeasureMaxLineWidth(IReadOnlyList<string> lines, double fontSize) {
            double max = 0D;
            for (int i = 0; i < lines.Count; i++) {
                max = Math.Max(max, EstimateTextWidth(lines[i], fontSize));
            }

            return max;
        }

        private static double EstimateTextWidth(string text, double fontSize) {
            double width = 0D;
            foreach (char c in text) {
                if (char.IsWhiteSpace(c)) {
                    width += fontSize * 0.32D;
                } else if ("ilI.,'!:;|".IndexOf(c) >= 0) {
                    width += fontSize * 0.26D;
                } else if ("MW@#%&".IndexOf(c) >= 0) {
                    width += fontSize * 0.86D;
                } else if (char.IsDigit(c)) {
                    width += fontSize * 0.56D;
                } else {
                    width += fontSize * 0.54D;
                }
            }

            return width;
        }

        private static double ResolveTextAnchorX(double centerX, double maxWidth, VisioTextHorizontalAlignment? alignment) {
            if (!IsFinitePositive(maxWidth)) {
                return centerX;
            }

            switch (alignment) {
                case VisioTextHorizontalAlignment.Left:
                    return centerX - (maxWidth / 2D);
                case VisioTextHorizontalAlignment.Right:
                    return centerX + (maxWidth / 2D);
                default:
                    return centerX;
            }
        }

        private static double ResolveTextTop(double centerY, double measuredHeight, double maxHeight, VisioTextVerticalAlignment? alignment) {
            if (!IsFinitePositive(maxHeight)) {
                return centerY - (measuredHeight / 2D);
            }

            switch (alignment) {
                case VisioTextVerticalAlignment.Top:
                    return centerY - (maxHeight / 2D);
                case VisioTextVerticalAlignment.Bottom:
                    return centerY + (maxHeight / 2D) - measuredHeight;
                default:
                    return centerY - (measuredHeight / 2D);
            }
        }

        private static double GetAlignedTextLeft(double anchorX, double width, VisioTextHorizontalAlignment? alignment) {
            switch (alignment) {
                case VisioTextHorizontalAlignment.Left:
                    return anchorX;
                case VisioTextHorizontalAlignment.Right:
                    return anchorX - width;
                default:
                    return anchorX - (width / 2D);
            }
        }

        private static bool IsFinitePositive(double value) =>
            value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);

        private static Color? ResolveTextBackground(VisioTextStyle? style, bool drawLabelBackground) {
            if (style?.BackgroundColor.HasValue == true) {
                return ApplyBackgroundTransparency(style.BackgroundColor.Value, style.BackgroundTransparency);
            }

            return drawLabelBackground ? Color.FromRgba(255, 255, 255, 230) : null;
        }

        private static Color ApplyBackgroundTransparency(Color color, double? transparency) {
            if (!transparency.HasValue) {
                return color;
            }

            double clamped = Math.Max(0D, Math.Min(100D, transparency.Value));
            byte alpha = (byte)Math.Round(color.A * (1D - (clamped / 100D)));
            return Color.FromRgba(color.R, color.G, color.B, alpha);
        }

        private static string FormatTextRotation(double radians, double centerX, double centerY) =>
            "rotate(" + Format(RadiansToDegrees(-radians)) + " " + Format(centerX) + " " + Format(centerY) + ")";

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

        private static string BuildPath(VisioPage page, VisioShape shape, IReadOnlyList<(double X, double Y)> localPoints, double scale, bool isClosed) {
            StringBuilder builder = new();
            for (int i = 0; i < localPoints.Count; i++) {
                (double absX, double absY) = GetPagePoint(shape, localPoints[i].X, localPoints[i].Y);
                (double x, double y) = ToSvg(page, absX, absY, scale);
                builder.Append(i == 0 ? "M " : " L ");
                builder.Append(Format(x)).Append(' ').Append(Format(y));
            }

            if (isClosed) {
                builder.Append(" Z");
            }

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

        private static void WriteSvgCircle(XmlWriter writer, double cx, double cy, double radius, Color color, bool fill, double strokeWidth) {
            writer.WriteStartElement("circle", SvgNamespace);
            writer.WriteAttributeString("cx", Format(cx));
            writer.WriteAttributeString("cy", Format(cy));
            writer.WriteAttributeString("r", Format(radius));
            if (fill) {
                WriteColor(writer, "fill", color);
                writer.WriteAttributeString("stroke", "none");
            } else {
                writer.WriteAttributeString("fill", "none");
                WriteColor(writer, "stroke", color);
                writer.WriteAttributeString("stroke-width", Format(strokeWidth));
                writer.WriteAttributeString("stroke-linecap", "round");
                writer.WriteAttributeString("stroke-linejoin", "round");
            }

            writer.WriteEndElement();
        }

        private static void WriteSvgRect(XmlWriter writer, double x, double y, double width, double height, Color color, bool fill, double strokeWidth) {
            writer.WriteStartElement("rect", SvgNamespace);
            writer.WriteAttributeString("x", Format(x));
            writer.WriteAttributeString("y", Format(y));
            writer.WriteAttributeString("width", Format(width));
            writer.WriteAttributeString("height", Format(height));
            writer.WriteAttributeString("rx", Format(Math.Min(width, height) * 0.08D));
            if (fill) {
                WriteColor(writer, "fill", color);
                writer.WriteAttributeString("stroke", "none");
            } else {
                writer.WriteAttributeString("fill", "none");
                WriteColor(writer, "stroke", color);
                writer.WriteAttributeString("stroke-width", Format(strokeWidth));
                writer.WriteAttributeString("stroke-linecap", "round");
                writer.WriteAttributeString("stroke-linejoin", "round");
            }

            writer.WriteEndElement();
        }

        private static void WriteSvgLine(XmlWriter writer, double x1, double y1, double x2, double y2, Color color, double strokeWidth) {
            writer.WriteStartElement("line", SvgNamespace);
            writer.WriteAttributeString("x1", Format(x1));
            writer.WriteAttributeString("y1", Format(y1));
            writer.WriteAttributeString("x2", Format(x2));
            writer.WriteAttributeString("y2", Format(y2));
            WriteColor(writer, "stroke", color);
            writer.WriteAttributeString("stroke-width", Format(strokeWidth));
            writer.WriteAttributeString("stroke-linecap", "round");
            writer.WriteEndElement();
        }

        private static void WriteSvgPath(XmlWriter writer, string data, Color color, bool fill, double strokeWidth) {
            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("d", data);
            if (fill) {
                WriteColor(writer, "fill", color);
                writer.WriteAttributeString("stroke", "none");
            } else {
                writer.WriteAttributeString("fill", "none");
                WriteColor(writer, "stroke", color);
                writer.WriteAttributeString("stroke-width", Format(strokeWidth));
                writer.WriteAttributeString("stroke-linecap", "round");
                writer.WriteAttributeString("stroke-linejoin", "round");
            }

            writer.WriteEndElement();
        }

        private static void WriteSvgCylinder(XmlWriter writer, double x, double y, double size, Color color) {
            double width = size * 0.62D;
            double height = size * 0.58D;
            double left = x - width / 2D;
            double top = y - height / 2D;
            WriteSvgPath(writer, "M " + Format(left) + " " + Format(top + height * 0.18D) +
                                 " C " + Format(left) + " " + Format(top - height * 0.02D) +
                                 " " + Format(left + width) + " " + Format(top - height * 0.02D) +
                                 " " + Format(left + width) + " " + Format(top + height * 0.18D) +
                                 " L " + Format(left + width) + " " + Format(top + height * 0.82D) +
                                 " C " + Format(left + width) + " " + Format(top + height * 1.02D) +
                                 " " + Format(left) + " " + Format(top + height * 1.02D) +
                                 " " + Format(left) + " " + Format(top + height * 0.82D) +
                                 " Z", color, fill: false, strokeWidth: Math.Max(1D, size * 0.045D));
            WriteSvgPath(writer, "M " + Format(left) + " " + Format(top + height * 0.18D) +
                                 " C " + Format(left) + " " + Format(top + height * 0.38D) +
                                 " " + Format(left + width) + " " + Format(top + height * 0.38D) +
                                 " " + Format(left + width) + " " + Format(top + height * 0.18D), color, fill: false, strokeWidth: Math.Max(1D, size * 0.045D));
        }

        private static void WriteSvgShield(XmlWriter writer, double x, double y, double size, Color color) {
            WriteSvgPath(writer, "M " + Format(x) + " " + Format(y - size * 0.36D) +
                                 " L " + Format(x + size * 0.3D) + " " + Format(y - size * 0.22D) +
                                 " L " + Format(x + size * 0.22D) + " " + Format(y + size * 0.22D) +
                                 " L " + Format(x) + " " + Format(y + size * 0.38D) +
                                 " L " + Format(x - size * 0.22D) + " " + Format(y + size * 0.22D) +
                                 " L " + Format(x - size * 0.3D) + " " + Format(y - size * 0.22D) +
                                 " Z", color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
        }

        private static void WriteSvgHex(XmlWriter writer, double x, double y, double size, Color color) {
            double r = size * 0.36D;
            WriteSvgPath(writer, "M " + Format(x) + " " + Format(y - r) +
                                 " L " + Format(x + r * 0.86D) + " " + Format(y - r * 0.5D) +
                                 " L " + Format(x + r * 0.86D) + " " + Format(y + r * 0.5D) +
                                 " L " + Format(x) + " " + Format(y + r) +
                                 " L " + Format(x - r * 0.86D) + " " + Format(y + r * 0.5D) +
                                 " L " + Format(x - r * 0.86D) + " " + Format(y - r * 0.5D) +
                                 " Z", color, fill: false, strokeWidth: Math.Max(1D, size * 0.05D));
        }

        private static string BuildCloudPath(double x, double y, double size) =>
            "M " + Format(x - size * 0.34D) + " " + Format(y + size * 0.12D) +
            " C " + Format(x - size * 0.48D) + " " + Format(y + size * 0.1D) +
            " " + Format(x - size * 0.45D) + " " + Format(y - size * 0.18D) +
            " " + Format(x - size * 0.2D) + " " + Format(y - size * 0.16D) +
            " C " + Format(x - size * 0.11D) + " " + Format(y - size * 0.42D) +
            " " + Format(x + size * 0.22D) + " " + Format(y - size * 0.35D) +
            " " + Format(x + size * 0.24D) + " " + Format(y - size * 0.1D) +
            " C " + Format(x + size * 0.48D) + " " + Format(y - size * 0.12D) +
            " " + Format(x + size * 0.51D) + " " + Format(y + size * 0.14D) +
            " " + Format(x + size * 0.3D) + " " + Format(y + size * 0.14D) +
            " Z";

        private sealed class TextLayout {
            internal TextLayout(string[] lines, double fontSize, double lineHeight, double width, double height) {
                Lines = lines;
                FontSize = fontSize;
                LineHeight = lineHeight;
                Width = width;
                Height = height;
            }

            internal string[] Lines { get; }

            internal double FontSize { get; }

            internal double LineHeight { get; }

            internal double Width { get; }

            internal double Height { get; }
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
