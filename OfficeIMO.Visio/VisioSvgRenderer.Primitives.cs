using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using Color = OfficeIMO.Drawing.OfficeColor;


namespace OfficeIMO.Visio {
    internal static partial class VisioSvgRenderer {
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
    }
}
