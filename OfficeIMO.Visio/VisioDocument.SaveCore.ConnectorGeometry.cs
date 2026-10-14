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

        private static void WriteConnectorGeometry(XmlWriter writer, string ns, VisioConnector connector, double startX, double startY, double endX, double endY) {
            VisioShape shape = VisioConnectorGeometry.CreateShape(connector);
            WritePreservedGeometrySections(writer, shape.PreservedGeometrySections);
        }

        private static readonly string[] ConnectorTransformCells = { "PinX", "PinY", "Width", "Height", "LocPinX", "LocPinY", "Angle", "FlipX", "FlipY" };

        private static void WriteConnectorTransform(XmlWriter writer, string ns, VisioConnector connector, string? cellName = null, XElement? original = null) {
            VisioShape shape = VisioConnectorGeometry.CreateShape(connector);
            bool native = connector.NativeGeometry?.AppliesTo(connector) == true;
            double[] values = { shape.PinX, shape.PinY, shape.Width, shape.Height, shape.LocPinX, shape.LocPinY, shape.Angle,
                native && connector.NativeGeometry!.FlipX ? 1 : 0, native && connector.NativeGeometry!.FlipY ? 1 : 0 };
            for (int i = 0; i < ConnectorTransformCells.Length; i++) {
                string name = ConnectorTransformCells[i];
                if (cellName != null && name != cellName) continue;
                if (original == null) WriteCell(writer, ns, name, values[i]);
                else {
                    XElement cell = new(original);
                    if (!double.TryParse((string?)cell.Attribute("V"), NumberStyles.Float, CultureInfo.InvariantCulture, out double before) || Math.Abs(before - values[i]) > 1e-10)
                        cell.Attribute("F")?.Remove();
                    cell.SetAttributeValue("V", ToVisioString(values[i]));
                    cell.WriteTo(writer);
                }
            }
        }

        private static void WritePreservedConnectAttributes(XmlWriter writer, IEnumerable<XAttribute> preservedAttributes) {
            foreach (XAttribute attribute in preservedAttributes) {
                writer.WriteAttributeString(
                    attribute.Name.LocalName,
                    attribute.Name.NamespaceName.Length == 0 ? null : attribute.Name.NamespaceName,
                    attribute.Value);
            }
        }

        private static void WriteConnectElement(
            XmlWriter writer,
            string ns,
            IReadOnlyDictionary<string, string> persistedIds,
            VisioConnector connector,
            VisioConnectorEndpointScope endpointScope) {
            if ((endpointScope == VisioConnectorEndpointScope.Start ? connector.From : connector.To) == null) return;
            IEnumerable<XAttribute> preservedAttributes = endpointScope == VisioConnectorEndpointScope.Start
                ? connector.PreservedBeginConnectAttributes
                : connector.PreservedEndConnectAttributes;
            IEnumerable<XName> attributeOrder = endpointScope == VisioConnectorEndpointScope.Start
                ? connector.PreservedBeginConnectAttributeOrder
                : connector.PreservedEndConnectAttributeOrder;

            writer.WriteStartElement("Connect", ns);
            WriteOrderedConnectAttributes(writer, persistedIds, connector, endpointScope, preservedAttributes, attributeOrder);
            writer.WriteEndElement();
        }

        private static void WriteOrderedConnectAttributes(
            XmlWriter writer,
            IReadOnlyDictionary<string, string> persistedIds,
            VisioConnector connector,
            VisioConnectorEndpointScope endpointScope,
            IEnumerable<XAttribute> preservedAttributes,
            IEnumerable<XName> attributeOrder) {
            XName fromSheetName = XName.Get("FromSheet");
            XName fromCellName = XName.Get("FromCell");
            XName toSheetName = XName.Get("ToSheet");
            XName toCellName = XName.Get("ToCell");

            List<(XName Name, string Value)> standardAttributeList = new() {
                (fromSheetName, GetPersistedId(persistedIds, connector.Id)),
                (fromCellName, endpointScope == VisioConnectorEndpointScope.Start ? "BeginX" : "EndX"),
                (toSheetName, endpointScope == VisioConnectorEndpointScope.Start
                    ? GetPersistedId(persistedIds, connector.From!.Id)
                    : GetPersistedId(persistedIds, connector.To!.Id)),
                (toCellName, endpointScope == VisioConnectorEndpointScope.Start
                    ? GetConnectionCell(connector.From!, connector.FromConnectionPoint, connector.PreservedFromConnectionCell)
                    : GetConnectionCell(connector.To!, connector.ToConnectionPoint, connector.PreservedToConnectionCell))
            };
            Dictionary<XName, string> standardAttributes = standardAttributeList.ToDictionary(attribute => attribute.Name, attribute => attribute.Value);

            List<XAttribute> preservedAttributeList = preservedAttributes.ToList();
            Dictionary<XName, XAttribute> preservedByName = preservedAttributeList.ToDictionary(attribute => attribute.Name, attribute => attribute);
            HashSet<XName> writtenNames = new();

            foreach (XName attributeName in attributeOrder) {
                if (attributeName == null) {
                    continue;
                }

                if (standardAttributes.TryGetValue(attributeName, out string? standardValue)) {
                    WriteAttribute(writer, attributeName, standardValue);
                    writtenNames.Add(attributeName);
                    continue;
                }

                if (preservedByName.TryGetValue(attributeName, out XAttribute? preservedAttribute)) {
                    WriteAttribute(writer, preservedAttribute.Name, preservedAttribute.Value);
                    writtenNames.Add(attributeName);
                }
            }

            foreach ((XName Name, string Value) standardAttribute in standardAttributeList) {
                if (writtenNames.Add(standardAttribute.Name)) {
                    WriteAttribute(writer, standardAttribute.Name, standardAttribute.Value);
                }
            }

            foreach (XAttribute preservedAttribute in preservedAttributeList) {
                if (writtenNames.Add(preservedAttribute.Name)) {
                    WriteAttribute(writer, preservedAttribute.Name, preservedAttribute.Value);
                }
            }
        }

        private static void WriteAttribute(XmlWriter writer, XName attributeName, string value) {
            writer.WriteAttributeString(
                attributeName.LocalName,
                attributeName.NamespaceName.Length == 0 ? null : attributeName.NamespaceName,
                value);
        }
    }
}
