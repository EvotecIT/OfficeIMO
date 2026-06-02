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

        private static void WriteStringCell(XmlWriter writer, string ns, string name, string value, string? formula = null) {
            writer.WriteStartElement("Cell", ns);
            writer.WriteAttributeString("N", name);
            writer.WriteAttributeString("V", value);
            if (!string.IsNullOrEmpty(formula)) writer.WriteAttributeString("F", formula);
            writer.WriteEndElement();
        }

        private static void WriteAttribute(XmlWriter writer, XAttribute attribute) {
            XNamespace attributeNamespace = attribute.Name.Namespace;
            string? namespaceName = attributeNamespace == XNamespace.None ? null : attributeNamespace.NamespaceName;
            writer.WriteAttributeString(null, attribute.Name.LocalName, namespaceName, attribute.Value);
        }
    }
}
