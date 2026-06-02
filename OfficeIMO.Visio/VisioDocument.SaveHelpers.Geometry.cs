using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Save-time helper methods for <see cref="VisioDocument"/>.
    /// </summary>
    public partial class VisioDocument {

        private static void WriteGeometryHeaderRow(XmlWriter writer, string ns) {
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "Geometry");
            WriteCellValue(writer, ns, "NoFill", "0");
            WriteCellValue(writer, ns, "NoLine", "0");
            WriteCellValue(writer, ns, "NoShow", "0");
            WriteCellValue(writer, ns, "NoSnap", "0");
            WriteCellValue(writer, ns, "NoQuickDrag", "0");
            writer.WriteEndElement();
        }

        private static void WriteXForm(XmlWriter writer, string ns, VisioShape shape, double width, double height) {
            WriteXForm(writer, ns, shape.PinX, shape.PinY, width, height, shape.LocPinX, shape.LocPinY, shape.Angle);
        }

        private static void WriteXForm(XmlWriter writer, string ns, double pinX, double pinY, double width, double height, double locPinX, double locPinY, double angle) {
            WriteCell(writer, ns, "PinX", pinX);
            WriteCell(writer, ns, "PinY", pinY);
            WriteCell(writer, ns, "Width", width);
            WriteCell(writer, ns, "Height", height);
            WriteCell(writer, ns, "LocPinX", locPinX);
            WriteCell(writer, ns, "LocPinY", locPinY);
            WriteCell(writer, ns, "Angle", angle);
        }

        private static void WriteXForm1D(XmlWriter writer, string ns, double beginX, double beginY, double endX, double endY) {
            writer.WriteStartElement("XForm1D", ns);
            writer.WriteElementString("BeginX", ns, ToVisioString(beginX));
            writer.WriteElementString("BeginY", ns, ToVisioString(beginY));
            writer.WriteElementString("EndX", ns, ToVisioString(endX));
            writer.WriteElementString("EndY", ns, ToVisioString(endY));
            writer.WriteEndElement();

            WriteCell(writer, ns, "BeginX", beginX);
            WriteCell(writer, ns, "BeginY", beginY);
            WriteCell(writer, ns, "EndX", endX);
            WriteCell(writer, ns, "EndY", endY);
        }

        private static void WriteRectangleGeometry(XmlWriter writer, string ns, double width, double height) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteEllipseGeometry(XmlWriter writer, string ns, double width, double height) {
            double rx = width / 2.0;
            double ry = height / 2.0;
            double centerX = rx;
            double centerY = ry;
            const int segments = 24;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", centerX + rx);
            WriteCell(writer, ns, "Y", centerY);
            writer.WriteEndElement();

            for (int i = 1; i <= segments; i++) {
                double angle = (Math.PI * 2D * i) / segments;
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("T", "LineTo");
                WriteCell(writer, ns, "X", centerX + (Math.Cos(angle) * rx));
                WriteCell(writer, ns, "Y", centerY + (Math.Sin(angle) * ry));
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private static void WriteDiamondGeometry(XmlWriter writer, string ns, double width, double height) {
            double midX = width / 2.0;
            double midY = height / 2.0;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", midY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", midY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteTriangleGeometry(XmlWriter writer, string ns, double width, double height) {
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width / 2.0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WritePentagonGeometry(XmlWriter writer, string ns, double width, double height) {
            double midX = width / 2.0;
            double shoulderY = height * 0.62;
            double lowerInset = width * 0.2;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", shoulderY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - lowerInset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", lowerInset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", shoulderY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteParallelogramGeometry(XmlWriter writer, string ns, double width, double height) {
            double offset = Math.Min(width / 4.0, Math.Max(width / 10.0, height / 3.0));
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", offset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - offset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", offset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteHexagonGeometry(XmlWriter writer, string ns, double width, double height) {
            double inset = Math.Min(width / 4.0, Math.Max(width / 8.0, height / 4.0));
            double midY = height / 2.0;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - inset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", midY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", midY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteTrapezoidGeometry(XmlWriter writer, string ns, double width, double height) {
            double inset = Math.Min(width / 5.0, Math.Max(width / 10.0, height / 4.0));
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width - inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", inset);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteOffPageReferenceGeometry(XmlWriter writer, string ns, double width, double height) {
            double midX = width / 2.0;
            double shoulderY = height * 0.45;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");
            WriteGeometryHeaderRow(writer, ns);

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", shoulderY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", midX);
            WriteCell(writer, ns, "Y", 0);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", shoulderY);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "LineTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", height);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }
    }
}
