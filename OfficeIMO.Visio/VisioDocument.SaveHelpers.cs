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
        private static string ToVisioString(double value) {
            string text = Math.Round(value, 15).ToString("F15", CultureInfo.InvariantCulture);
            return text.TrimEnd('0').TrimEnd('.');
        }

        private static void WriteCell(XmlWriter writer, string ns, string name, double value) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", ToVisioString(value));
            writer.WriteEndElement();
        }

        private static void WriteCell(XmlWriter writer, string ns, string name, double value, string? unit, string? formula) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", ToVisioString(value));
            if (!string.IsNullOrEmpty(unit)) writer.WriteAttributeString("U", unit);
            if (!string.IsNullOrEmpty(formula)) writer.WriteAttributeString("F", formula);
            writer.WriteEndElement();
        }

        private static void WriteCellValue(XmlWriter writer, string ns, string name, string value) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", value);
            writer.WriteEndElement();
        }

        private static void WriteCellValue(XmlWriter writer, string ns, string name, string value, string? unit, string? formula) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", value);
            if (!string.IsNullOrEmpty(unit)) writer.WriteAttributeString("U", unit);
            if (!string.IsNullOrEmpty(formula)) writer.WriteAttributeString("F", formula);
            writer.WriteEndElement();
        }

        private static void WriteXForm(XmlWriter writer, string ns, VisioShape shape, double width, double height) {
            WriteCell(writer, ns, "PinX", shape.PinX);
            WriteCell(writer, ns, "PinY", shape.PinY);
            WriteCell(writer, ns, "Width", width);
            WriteCell(writer, ns, "Height", height);
            WriteCell(writer, ns, "LocPinX", shape.LocPinX);
            WriteCell(writer, ns, "LocPinY", shape.LocPinY);
            WriteCell(writer, ns, "Angle", shape.Angle);
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

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "Geometry");
            WriteCellValue(writer, ns, "NoFill", "0");
            WriteCellValue(writer, ns, "NoLine", "0");
            WriteCellValue(writer, ns, "NoShow", "0");
            WriteCellValue(writer, ns, "NoSnap", "0");
            WriteCellValue(writer, ns, "NoQuickDrag", "0");
            writer.WriteEndElement();

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
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");

            // Match rectangle behavior by emitting a Geometry row with default flags
            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "Geometry");
            WriteCellValue(writer, ns, "NoFill", "0");
            WriteCellValue(writer, ns, "NoLine", "0");
            WriteCellValue(writer, ns, "NoShow", "0");
            WriteCellValue(writer, ns, "NoSnap", "0");
            WriteCellValue(writer, ns, "NoQuickDrag", "0");
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "MoveTo");
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", ry);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "EllipticalArcTo");
            WriteCell(writer, ns, "X", width);
            WriteCell(writer, ns, "Y", ry);
            WriteCell(writer, ns, "A", ry);
            WriteCell(writer, ns, "B", width);
            writer.WriteEndElement();

            writer.WriteStartElement("Row", ns);
            writer.WriteAttributeString("T", "EllipticalArcTo");
            // Explicitly close back to the leftmost point to avoid viewer quirks
            WriteCell(writer, ns, "X", 0);
            WriteCell(writer, ns, "Y", ry);
            WriteCell(writer, ns, "A", ry);
            WriteCell(writer, ns, "B", width);
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        private static void WriteDiamondGeometry(XmlWriter writer, string ns, double width, double height) {
            double midX = width / 2.0;
            double midY = height / 2.0;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Geometry");
            writer.WriteAttributeString("IX", "0");

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

        private static void WriteConnectionSection(XmlWriter writer, string ns, IList<VisioConnectionPoint> points) {
            if (points.Count == 0) return;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Connection");
            for (int i = 0; i < points.Count; i++) {
                VisioConnectionPoint cp = points[i];
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("IX", XmlConvert.ToString(i));
                WriteCell(writer, ns, "X", cp.X);
                WriteCell(writer, ns, "Y", cp.Y);
                WriteCell(writer, ns, "DirX", cp.DirX);
                WriteCell(writer, ns, "DirY", cp.DirY);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private static void WriteDataSection(XmlWriter writer, string ns, IDictionary<string, string> data) {
            if (data.Count == 0) return;
            writer.WriteStartElement("Section", ns);
            writer.WriteAttributeString("N", "Prop");
            foreach (var kv in data) {
                writer.WriteStartElement("Row", ns);
                writer.WriteAttributeString("N", kv.Key);
                writer.WriteStartElement("Cell", ns);
                writer.WriteAttributeString("N", "Value");
                writer.WriteAttributeString("V", kv.Value);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        private static void WritePageCell(XmlWriter writer, string ns, string name, double value, string? unit = null, string? formula = null) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);

            bool isMillimeters = string.Equals(unit, "MM", StringComparison.OrdinalIgnoreCase);
            double serializedValue = isMillimeters
                ? Math.Round(value * 25.4, 8, MidpointRounding.AwayFromZero)
                : value;

            writer.WriteAttributeString("V", ToVisioString(serializedValue));

            if (!string.IsNullOrEmpty(unit)) {
                writer.WriteAttributeString("U", isMillimeters ? "MM" : unit);
            }

            if (formula != null) writer.WriteAttributeString("F", formula);
            writer.WriteEndElement();
        }

        private static void WriteTextElement(XmlWriter writer, string ns, string? text) {
            if (!string.IsNullOrEmpty(text)) writer.WriteElementString("Text", ns, text);
        }

        private static string GetConnectionCell(VisioShape shape, VisioConnectionPoint? point) {
            if (point == null) return "PinX";
            int index = shape.ConnectionPoints.IndexOf(point);
            return index >= 0 ? $"Connections.X{index + 1}" : "PinX";
        }

        private static XDocument CreateVisioDocumentXml(bool requestRecalcOnOpen) {
            XNamespace ns = VisioNamespace;
            XElement settings = new(ns + "DocumentSettings",
                new XAttribute("TopPage", 0),
                new XAttribute("DefaultTextStyle", 0),
                new XAttribute("DefaultLineStyle", 0),
                new XAttribute("DefaultFillStyle", 0),
                new XAttribute("DefaultGuideStyle", 4),
                new XElement(ns + "GlueSettings", 9),
                new XElement(ns + "SnapSettings", 295),
                new XElement(ns + "SnapExtensions", 34),
                new XElement(ns + "SnapAngles"),
                new XElement(ns + "DynamicGridEnabled", 1),
                new XElement(ns + "ProtectStyles", 0),
                new XElement(ns + "ProtectShapes", 0),
                new XElement(ns + "ProtectMasters", 0),
                new XElement(ns + "ProtectBkgnds", 0));
            if (requestRecalcOnOpen) settings.Add(new XElement(ns + "RelayoutAndRerouteUponOpen", 1));
            XElement styleSheets = new(ns + "StyleSheets",
                new XElement(ns + "StyleSheet",
                    new XAttribute("ID", 0),
                    new XAttribute("Name", "No Style"),
                    new XAttribute("NameU", "No Style"),
                    new XElement(ns + "Cell", new XAttribute("N", "EnableLineProps"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "EnableFillProps"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "EnableTextProps"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "LineWeight"), new XAttribute("V", "0.01041666666666667")),
                    new XElement(ns + "Cell", new XAttribute("N", "LineColor"), new XAttribute("V", "0")),
                    new XElement(ns + "Cell", new XAttribute("N", "LinePattern"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "FillForegnd"), new XAttribute("V", "1")),
                    new XElement(ns + "Cell", new XAttribute("N", "FillPattern"), new XAttribute("V", "1"))),
                new XElement(ns + "StyleSheet",
                    new XAttribute("ID", 1),
                    new XAttribute("Name", "Normal"),
                    new XAttribute("NameU", "Normal"),
                    new XAttribute("BasedOn", 0),
                    new XAttribute("LineStyle", 0),
                    new XAttribute("FillStyle", 0),
                    new XAttribute("TextStyle", 0),
                    new XElement(ns + "Cell", new XAttribute("N", "LinePattern"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "LineColor"), new XAttribute("V", "#000000")),
                    new XElement(ns + "Cell", new XAttribute("N", "FillPattern"), new XAttribute("V", 1)),
                    new XElement(ns + "Cell", new XAttribute("N", "FillForegnd"), new XAttribute("V", "#FFFFFF"))),
                new XElement(ns + "StyleSheet",
                    new XAttribute("ID", 2),
                    new XAttribute("Name", "Connector"),
                    new XAttribute("NameU", "Connector"),
                    new XAttribute("BasedOn", 1),
                    new XAttribute("LineStyle", 0),
                    new XAttribute("FillStyle", 0),
                    new XAttribute("TextStyle", 0),
                    new XElement(ns + "Cell", new XAttribute("N", "EndArrow"), new XAttribute("V", 0))));

            return new XDocument(
                new XElement(ns + "VisioDocument",
                    settings,
                    new XElement(ns + "Colors"),
                    new XElement(ns + "FaceNames"),
                    styleSheets));
        }

        private static void FixContentTypes(string filePath, int masterCount, bool includeTheme, IEnumerable<string> pagePartNames) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("File path cannot be null or whitespace.", nameof(filePath));
            }

            if (pagePartNames is null) {
                throw new ArgumentNullException(nameof(pagePartNames));
            }

            using FileStream zipStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(zipStream, ZipArchiveMode.Update);
            ZipArchiveEntry? entry = archive.GetEntry("[Content_Types].xml");
            entry?.Delete();
            ZipArchiveEntry newEntry = archive.CreateEntry("[Content_Types].xml");
            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement root = new(ct + "Types",
                new XElement(ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                new XElement(ct + "Default", new XAttribute("Extension", "emf"), new XAttribute("ContentType", "image/x-emf")));

            HashSet<string> overridePartNames = new(StringComparer.OrdinalIgnoreCase);
            void AddOverride(string partName, string contentType) {
                if (string.IsNullOrWhiteSpace(partName)) {
                    return;
                }

                string normalizedPartName = NormalizePartName(partName);

                if (overridePartNames.Add(normalizedPartName)) {
                    root.Add(new XElement(ct + "Override",
                        new XAttribute("PartName", normalizedPartName),
                        new XAttribute("ContentType", contentType)));
                }
            }

            AddOverride("/visio/document.xml", DocumentContentType);
            AddOverride("/visio/pages/pages.xml", PagesContentType);
            AddOverride("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
            AddOverride("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
            AddOverride("/docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml");
            AddOverride("/docProps/thumbnail.emf", "image/x-emf");
            AddOverride("/visio/windows.xml", WindowsContentType);

            foreach (string partName in pagePartNames) {
                AddOverride(partName, PageContentType);
            }
            if (includeTheme) {
                AddOverride("/visio/theme/theme1.xml", ThemeContentType);
            }
            if (masterCount > 0) {
                AddOverride("/visio/masters/masters.xml", "application/vnd.ms-visio.masters+xml");
                for (int i = 1; i <= masterCount; i++) {
                    AddOverride($"/visio/masters/master{i}.xml", "application/vnd.ms-visio.master+xml");
                }
            }
            XDocument doc = new(root);
            using StreamWriter writer = new(newEntry.Open());
            writer.Write(doc.Declaration + Environment.NewLine + doc.ToString(SaveOptions.DisableFormatting));
        }

        private static string NormalizePartName(string partName) {
            if (partName is null) {
                throw new ArgumentNullException(nameof(partName));
            }

            return "/" + partName.TrimStart('/');
        }
    }
}
