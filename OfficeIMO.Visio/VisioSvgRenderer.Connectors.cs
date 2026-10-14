using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
        private static void WriteConnector(XmlWriter writer, VisioPage page, VisioConnector connector, VisioSvgSaveOptions options, VisioRenderProjection projection, VisioRenderLabelLayout? labelLayout, VisioNativeTextStyleResolver textStyles) {
            List<(double X, double Y)> points = GetConnectorPoints(connector);
            writer.WriteStartElement("g", SvgNamespace);
            writer.WriteAttributeString("data-visio-connector-id", connector.Id);

            writer.WriteStartElement("path", SvgNamespace);
            writer.WriteAttributeString("d", BuildOpenPath(page, points, projection));
            writer.WriteAttributeString("fill", "none");
            bool visibleLine = VisioConnectorGeometry.HasVisibleLine(connector);
            double strokeWidth = Math.Max(connector.LineWeight * projection.PhysicalDensity, 0.75D);
            if (!visibleLine) {
                writer.WriteAttributeString("stroke", "none");
            } else {
                OfficeSvgFormatting.WriteColorAttribute(writer, "stroke", connector.LineColor);
                writer.WriteNumberAttribute("stroke-width", strokeWidth);
                writer.WriteStrokeLineCapAttribute(OfficeStrokeLineCap.Round);
                writer.WriteStrokeLineJoinAttribute(OfficeStrokeLineJoin.Round);
                writer.WriteStrokeDashStyleAttribute(OfficeStrokeDashStyleMapper.FromVisioLinePattern(connector.LinePattern), strokeWidth);
            }

            writer.WriteEndElement();

            if (visibleLine) {
                if (connector.BeginArrow.HasValue && connector.BeginArrow.Value != EndArrow.None && OfficeGeometry.TryGetArrowheadSegment(points, fromStart: true, out (double X, double Y) beginTip, out (double X, double Y) beginFrom)) {
                    WriteArrow(writer, page, beginTip, beginFrom, projection, connector.LineColor, strokeWidth, "start");
                }

                if (connector.EndArrow.HasValue && connector.EndArrow.Value != EndArrow.None && OfficeGeometry.TryGetArrowheadSegment(points, fromStart: false, out (double X, double Y) endTip, out (double X, double Y) endFrom)) {
                    WriteArrow(writer, page, endTip, endFrom, projection, connector.LineColor, strokeWidth, "end");
                }
            }

            if (options.RenderConnectorLabels && !string.IsNullOrEmpty(connector.Label)) {
                VisioRenderConnectorLabelPlacement label = labelLayout?.Resolve(connector, points) ?? VisioRenderLabelLayout.ResolveUnadjusted(connector, points, projection);
                (double labelCenterX, double labelCenterY) = VisioConnectorGeometry.GetLabelCenter(connector, label.X, label.Y, label.Width, label.Height);
                (double x, double y) = ToSvg(page, labelCenterX, labelCenterY, projection);
                double maxWidth = label.Width * projection.GeometryDensity;
                double maxHeight = label.Height * projection.GeometryDensity;
                WriteText(
                    writer,
                    connector.Label!,
                    x,
                    y,
                    connector.TextStyle,
                    defaultSize: 9D,
                    projection.PhysicalDensity,
                    rotateRadians: VisioConnectorLabelFrame.ResolveAngle(connector),
                    maxWidth,
                    maxHeight,
                    drawLabelBackground: true,
                    labelAdjusted: label.Adjusted,
                    richText: VisioRichTextProjection.Create(page, connector, projection.PhysicalDensity, options.CancellationToken, textStyles),
                    cancellationToken: options.CancellationToken);
            }

            writer.WriteEndElement();
        }

        private static List<(double X, double Y)> GetConnectorPoints(VisioConnector connector) {
            return VisioConnectorGeometry.GetPoints(connector);
        }



    }
}
