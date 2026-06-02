using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {

        private static void WritePreservedConnectorCells(XmlWriter writer, IEnumerable<XElement> preservedCells) => WritePreservedElements(writer, preservedCells);

        private static void WritePreservedConnectorSections(XmlWriter writer, IEnumerable<XElement> preservedSections) => WritePreservedElements(writer, preservedSections);

        private static void WritePreservedAttributes(XmlWriter writer, IEnumerable<XAttribute> preservedAttributes) {
            foreach (XAttribute attribute in preservedAttributes) {
                writer.WriteAttributeString(
                    attribute.Name.LocalName,
                    attribute.Name.NamespaceName.Length == 0 ? null : attribute.Name.NamespaceName,
                    attribute.Value);
            }
        }

        private static void WritePreservedElements(XmlWriter writer, IEnumerable<XElement> preservedElements) {
            foreach (XElement element in preservedElements) {
                XElement clone = new(element);
                using var reader = clone.CreateReader();
                writer.WriteNode(reader, false);
            }
        }

        private static void WriteShapeGeometry(XmlWriter writer, string ns, IEnumerable<XElement> preservedGeometrySections, string? nameU, double width, double height, bool writeGeneratedGeometryWhenEmpty = true) {
            if (WritePreservedGeometrySections(writer, preservedGeometrySections)) {
                return;
            }

            if (writeGeneratedGeometryWhenEmpty) {
                WriteMasterGeometry(writer, ns, nameU, width, height);
            }
        }

        private static bool WritePreservedGeometrySections(XmlWriter writer, IEnumerable<XElement> preservedGeometrySections) {
            bool wroteGeometry = false;
            foreach (XElement section in preservedGeometrySections) {
                XElement clone = new(section);
                using var reader = clone.CreateReader();
                writer.WriteNode(reader, false);
                wroteGeometry = true;
            }

            return wroteGeometry;
        }
    }
}
