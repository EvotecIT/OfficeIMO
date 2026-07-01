using System;
using System.Xml;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared XML writer helpers for simple SVG primitives used by OfficeIMO renderers.
/// </summary>
public static class OfficeSvgPrimitiveWriter {
    /// <summary>
    /// Writes a filled or stroked SVG circle.
    /// </summary>
    public static void WriteCircle(XmlWriter writer, string svgNamespace, double cx, double cy, double radius, OfficeColor color, bool fill, double strokeWidth) {
        if (writer == null) {
            throw new ArgumentNullException(nameof(writer));
        }

        writer.WriteStartElement("circle", svgNamespace);
        writer.WriteNumberAttribute("cx", cx);
        writer.WriteNumberAttribute("cy", cy);
        writer.WriteNumberAttribute("r", radius);
        WriteFillOrStroke(writer, color, fill, strokeWidth);
        writer.WriteEndElement();
    }

    /// <summary>
    /// Writes a filled or stroked SVG rectangle.
    /// </summary>
    public static void WriteRectangle(XmlWriter writer, string svgNamespace, double x, double y, double width, double height, OfficeColor color, bool fill, double strokeWidth, double cornerRadius = 0D) {
        if (writer == null) {
            throw new ArgumentNullException(nameof(writer));
        }

        writer.WriteStartElement("rect", svgNamespace);
        writer.WriteNumberAttribute("x", x);
        writer.WriteNumberAttribute("y", y);
        writer.WriteNumberAttribute("width", width);
        writer.WriteNumberAttribute("height", height);
        if (cornerRadius > 0D) {
            writer.WriteNumberAttribute("rx", cornerRadius);
        }

        WriteFillOrStroke(writer, color, fill, strokeWidth);
        writer.WriteEndElement();
    }

    /// <summary>
    /// Writes a stroked SVG line with rounded line caps.
    /// </summary>
    public static void WriteLine(XmlWriter writer, string svgNamespace, double x1, double y1, double x2, double y2, OfficeColor color, double strokeWidth) {
        if (writer == null) {
            throw new ArgumentNullException(nameof(writer));
        }

        writer.WriteStartElement("line", svgNamespace);
        writer.WriteNumberAttribute("x1", x1);
        writer.WriteNumberAttribute("y1", y1);
        writer.WriteNumberAttribute("x2", x2);
        writer.WriteNumberAttribute("y2", y2);
        OfficeSvgFormatting.WriteColorAttribute(writer, "stroke", color);
        writer.WriteNumberAttribute("stroke-width", strokeWidth);
        writer.WriteStrokeLineCapAttribute(OfficeStrokeLineCap.Round);
        writer.WriteEndElement();
    }

    /// <summary>
    /// Writes a filled or stroked SVG path.
    /// </summary>
    public static void WritePath(XmlWriter writer, string svgNamespace, string data, OfficeColor color, bool fill, double strokeWidth) {
        if (writer == null) {
            throw new ArgumentNullException(nameof(writer));
        }

        if (string.IsNullOrEmpty(data)) {
            return;
        }

        writer.WriteStartElement("path", svgNamespace);
        writer.WriteAttributeString("d", data);
        WriteFillOrStroke(writer, color, fill, strokeWidth);
        writer.WriteEndElement();
    }

    private static void WriteFillOrStroke(XmlWriter writer, OfficeColor color, bool fill, double strokeWidth) {
        if (fill) {
            OfficeSvgFormatting.WriteColorAttribute(writer, "fill", color);
            writer.WriteAttributeString("stroke", "none");
            return;
        }

        writer.WriteAttributeString("fill", "none");
        OfficeSvgFormatting.WriteColorAttribute(writer, "stroke", color);
        writer.WriteNumberAttribute("stroke-width", strokeWidth);
        writer.WriteStrokeLineCapAttribute(OfficeStrokeLineCap.Round);
        writer.WriteStrokeLineJoinAttribute(OfficeStrokeLineJoin.Round);
    }
}
