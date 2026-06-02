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

        private static void ReserveRawMasterInstanceChildIds(VisioShape shape, VisioMaster master, Action<string> reserve) {
            XElement? rootShape = FindFirstMasterShape(master.RawMasterContentXml!);
            if (rootShape == null) {
                return;
            }

            foreach (XElement childShape in GetRawMasterChildShapes(rootShape)) {
                ReserveRawMasterInstanceChildId(shape.Id, childShape, reserve);
            }
        }

        private static void ReserveRawMasterInstanceChildId(string instanceShapeId, XElement masterShape, Action<string> reserve) {
            string? masterShapeId = masterShape.Attribute("ID")?.Value;
            if (string.IsNullOrWhiteSpace(masterShapeId)) {
                return;
            }

            reserve(GetRawMasterInstanceChildKey(instanceShapeId, masterShapeId!));
            foreach (XElement childShape in GetRawMasterChildShapes(masterShape)) {
                ReserveRawMasterInstanceChildId(instanceShapeId, childShape, reserve);
            }
        }

        private static void WriteRawMasterInstanceChildShapes(XmlWriter writer, string ns, VisioShape shape, VisioMaster master, IReadOnlyDictionary<string, string> persistedIds) {
            XElement? rootShape = FindFirstMasterShape(master.RawMasterContentXml!);
            if (rootShape == null) {
                return;
            }

            List<XElement> childShapes = GetRawMasterChildShapes(rootShape).ToList();
            if (childShapes.Count == 0) {
                return;
            }

            writer.WriteStartElement("Shapes", ns);
            foreach (XElement childShape in childShapes) {
                WriteRawMasterInstanceChildShape(writer, ns, shape.Id, childShape, persistedIds);
            }
            writer.WriteEndElement();
        }

        private static void WriteRawMasterInstanceChildShape(XmlWriter writer, string ns, string instanceShapeId, XElement masterShape, IReadOnlyDictionary<string, string> persistedIds) {
            string? masterShapeId = masterShape.Attribute("ID")?.Value;
            if (string.IsNullOrWhiteSpace(masterShapeId)) {
                return;
            }

            writer.WriteStartElement("Shape", ns);
            writer.WriteAttributeString("ID", GetPersistedId(persistedIds, GetRawMasterInstanceChildKey(instanceShapeId, masterShapeId!)));
            WriteRawMasterInstanceChildAttribute(writer, masterShape, "NameU");
            WriteRawMasterInstanceChildAttribute(writer, masterShape, "IsCustomNameU");
            WriteRawMasterInstanceChildAttribute(writer, masterShape, "Name");
            WriteRawMasterInstanceChildAttribute(writer, masterShape, "IsCustomName");
            string type = masterShape.Attribute("Type")?.Value ?? (GetRawMasterChildShapes(masterShape).Any() ? "Group" : "Shape");
            writer.WriteAttributeString("Type", type);
            writer.WriteAttributeString("MasterShape", masterShapeId);

            List<XElement> childShapes = GetRawMasterChildShapes(masterShape).ToList();
            if (childShapes.Count > 0) {
                writer.WriteStartElement("Shapes", ns);
                foreach (XElement childShape in childShapes) {
                    WriteRawMasterInstanceChildShape(writer, ns, instanceShapeId, childShape, persistedIds);
                }
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private static void WriteRawMasterInstanceChildAttribute(XmlWriter writer, XElement masterShape, string attributeName) {
            string? value = masterShape.Attribute(attributeName)?.Value;
            if (!string.IsNullOrWhiteSpace(value)) {
                writer.WriteAttributeString(attributeName, value);
            }
        }

        private static IEnumerable<XElement> GetRawMasterChildShapes(XElement masterShape) {
            XElement? shapes = masterShape
                .Elements()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "Shapes", StringComparison.OrdinalIgnoreCase));
            return shapes == null
                ? Enumerable.Empty<XElement>()
                : shapes.Elements().Where(element => string.Equals(element.Name.LocalName, "Shape", StringComparison.OrdinalIgnoreCase));
        }

        private static string GetRawMasterInstanceChildKey(string instanceShapeId, string masterShapeId) {
            return instanceShapeId + "::raw-master-shape::" + masterShapeId;
        }

        private static string GetPersistedId(IReadOnlyDictionary<string, string> persistedIds, string originalId) {
            return persistedIds.TryGetValue(originalId, out string? persistedId) ? persistedId : originalId;
        }

        private static KeyValuePair<string, string>? GetOriginalIdEntry(IReadOnlyDictionary<string, string> persistedIds, string originalId) {
            string persistedId = GetPersistedId(persistedIds, originalId);
            return string.Equals(persistedId, originalId, StringComparison.Ordinal)
                ? null
                : new KeyValuePair<string, string>(OriginalIdPropName, originalId);
        }
    }
}
