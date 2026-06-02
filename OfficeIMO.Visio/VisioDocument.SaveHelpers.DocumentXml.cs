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

        private static void WritePageCell(XmlWriter writer, string ns, string name, double value, string? unit = null, string? formula = null) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", XmlConvert.ToString(value));
            if (unit != null) writer.WriteAttributeString("U", unit);
            if (formula != null) writer.WriteAttributeString("F", formula);
            writer.WriteEndElement();
        }

        private static void WriteTextElement(XmlWriter writer, string ns, string? text, XElement? preservedTextElement = null, string? preservedTextValue = null) {
            if (preservedTextElement != null &&
                string.Equals(text ?? string.Empty, preservedTextValue ?? string.Empty, StringComparison.Ordinal)) {
                XElement clone = new(preservedTextElement);
                using var reader = clone.CreateReader();
                writer.WriteNode(reader, false);
                return;
            }

            if (!string.IsNullOrEmpty(text)) {
                writer.WriteElementString("Text", ns, text);
            }
        }

        private static string GetConnectionCell(VisioShape shape, VisioConnectionPoint? point, string? preservedCell = null) {
            if (point == null) {
                return string.IsNullOrWhiteSpace(preservedCell) ? "PinX" : preservedCell!;
            }

            Dictionary<VisioConnectionPoint, int> pointIndices = BuildConnectionPointIndices(shape.ConnectionPoints);
            return pointIndices.TryGetValue(point, out int index)
                ? $"Connections.X{index + 1}"
                : string.IsNullOrWhiteSpace(preservedCell) ? "PinX" : preservedCell!;
        }

        private static XDocument CreateVisioDocumentXml(
            bool requestRecalcOnOpen,
            IEnumerable<XAttribute>? preservedDocumentAttributes = null,
            IEnumerable<XElement>? preservedDocumentElements = null,
            IEnumerable<XAttribute>? preservedDocumentSettingsAttributes = null,
            IEnumerable<XElement>? preservedDocumentSettingsElements = null,
            IEnumerable<XAttribute>? preservedColorsAttributes = null,
            IEnumerable<XElement>? preservedColorsElements = null,
            IEnumerable<XAttribute>? preservedFaceNamesAttributes = null,
            IEnumerable<XElement>? preservedFaceNamesElements = null,
            IEnumerable<XAttribute>? preservedStyleSheetsAttributes = null,
            IEnumerable<XElement>? preservedStyleSheetsElements = null,
            IDictionary<string, PreservedStyleSheetData>? preservedGeneratedStyleSheets = null,
            IEnumerable<XElement>? preservedAdditionalStyleSheets = null) {
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
            foreach (XAttribute attribute in preservedDocumentSettingsAttributes ?? Enumerable.Empty<XAttribute>()) {
                settings.Add(new XAttribute(attribute));
            }
            if (requestRecalcOnOpen) settings.Add(new XElement(ns + "RelayoutAndRerouteUponOpen", 1));
            foreach (XElement element in preservedDocumentSettingsElements ?? Enumerable.Empty<XElement>()) {
                settings.Add(new XElement(element));
            }
            XElement colors = new(ns + "Colors");
            foreach (XAttribute attribute in preservedColorsAttributes ?? Enumerable.Empty<XAttribute>()) {
                colors.Add(new XAttribute(attribute));
            }
            foreach (XElement element in preservedColorsElements ?? Enumerable.Empty<XElement>()) {
                colors.Add(new XElement(element));
            }
            XElement faceNames = new(ns + "FaceNames");
            foreach (XAttribute attribute in preservedFaceNamesAttributes ?? Enumerable.Empty<XAttribute>()) {
                faceNames.Add(new XAttribute(attribute));
            }
            foreach (XElement element in preservedFaceNamesElements ?? Enumerable.Empty<XElement>()) {
                faceNames.Add(new XElement(element));
            }
            XElement styleSheets = new(ns + "StyleSheets");
            foreach (XAttribute attribute in preservedStyleSheetsAttributes ?? Enumerable.Empty<XAttribute>()) {
                styleSheets.Add(new XAttribute(attribute));
            }
            foreach (XElement element in preservedStyleSheetsElements ?? Enumerable.Empty<XElement>()) {
                styleSheets.Add(new XElement(element));
            }
            styleSheets.Add(CreateGeneratedStyleSheet(ns, "0", preservedGeneratedStyleSheets));
            styleSheets.Add(CreateGeneratedStyleSheet(ns, "1", preservedGeneratedStyleSheets));
            styleSheets.Add(CreateGeneratedStyleSheet(ns, "2", preservedGeneratedStyleSheets));
            foreach (XElement styleSheet in preservedAdditionalStyleSheets ?? Enumerable.Empty<XElement>()) {
                styleSheets.Add(new XElement(styleSheet));
            }

            XElement root = new(ns + "VisioDocument");
            foreach (XAttribute attribute in preservedDocumentAttributes ?? Enumerable.Empty<XAttribute>()) {
                root.Add(new XAttribute(attribute));
            }
            foreach (XElement element in preservedDocumentElements ?? Enumerable.Empty<XElement>()) {
                root.Add(new XElement(element));
            }
            root.Add(settings);
            root.Add(colors);
            root.Add(faceNames);
            root.Add(styleSheets);

            return new XDocument(root);
        }

        private static XElement CreateGeneratedStyleSheet(XNamespace ns, string styleSheetId, IDictionary<string, PreservedStyleSheetData>? preservedGeneratedStyleSheets) {
            XElement styleSheet = styleSheetId switch {
                "0" => new XElement(ns + "StyleSheet",
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
                "1" => new XElement(ns + "StyleSheet",
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
                "2" => new XElement(ns + "StyleSheet",
                    new XAttribute("ID", 2),
                    new XAttribute("Name", "Connector"),
                    new XAttribute("NameU", "Connector"),
                    new XAttribute("BasedOn", 1),
                    new XAttribute("LineStyle", 0),
                    new XAttribute("FillStyle", 0),
                    new XAttribute("TextStyle", 0),
                    new XElement(ns + "Cell", new XAttribute("N", "EndArrow"), new XAttribute("V", 0))),
                _ => throw new InvalidOperationException($"Unsupported generated style sheet id '{styleSheetId}'.")
            };

            if (preservedGeneratedStyleSheets != null &&
                preservedGeneratedStyleSheets.TryGetValue(styleSheetId, out PreservedStyleSheetData? preserved)) {
                foreach (XAttribute attribute in preserved.Attributes) {
                    styleSheet.Add(new XAttribute(attribute));
                }

                foreach (XElement element in preserved.ChildElements) {
                    styleSheet.Add(new XElement(element));
                }
            }

            return styleSheet;
        }
    }
}
