using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using Color = OfficeIMO.Drawing.OfficeColor;


namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
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
                (double labelCenterX, double labelCenterY) = ResolveTextBoxCenter(label.X, label.Y, label.Width, label.Height, connector.TextStyle);
                (double x, double y) = ToSvg(page, labelCenterX, labelCenterY, scale);
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
    }
}
