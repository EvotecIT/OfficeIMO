using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSmartArt {
        private const string PersistedDiagramNamespace =
            "http://schemas.microsoft.com/office/drawing/2008/diagram";

        private bool TrySynchronizePersistedBasicProcessTopology(
            XDocument dataDocument,
            (XNamespace dgm, XNamespace a) ns,
            PowerPointSmartArtTopology current,
            IReadOnlyList<PowerPointSmartArtNode> requested,
            out DiagramPersistLayoutPart? persistPart,
            out XDocument? persistDocument,
            out string diagnostic) {
            persistPart = null;
            persistDocument = null;
            diagnostic = string.Empty;

            XElement? dataModelExtension = dataDocument.Descendants()
                .FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName,
                        "dataModelExt", StringComparison.Ordinal)
                    && string.Equals(element.Name.NamespaceName,
                        PersistedDiagramNamespace, StringComparison.Ordinal));
            if (dataModelExtension == null) {
                return true;
            }

            string? relationshipId = (string?)dataModelExtension
                .Attribute("relId");
            if (current.Kind != OfficeDiagramKind.Process
                || string.IsNullOrWhiteSpace(relationshipId)
                || !_slidePart.TryGetPartById(relationshipId!,
                    out OpenXmlPart? relatedPart)
                || relatedPart is not DiagramPersistLayoutPart drawingPart) {
                diagnostic = "The producer SmartArt presentation graph cannot be synchronized safely and remains preservation-only for topology changes.";
                return false;
            }

            try {
                XDocument drawingDocument;
                using (Stream stream = drawingPart.GetStream(FileMode.Open,
                           FileAccess.Read)) {
                    drawingDocument = PowerPointXmlReader
                        .LoadPackagePartXml(stream);
                }

                if (!TrySynchronizeBasicProcessPresentationGraph(
                        dataDocument, drawingDocument, ns, current,
                        requested)) {
                    diagnostic = "The producer SmartArt presentation graph does not match the supported linear Basic Process model and remains preservation-only for topology changes.";
                    return false;
                }

                persistPart = drawingPart;
                persistDocument = drawingDocument;
                return true;
            } catch (Exception exception) when (
                exception is InvalidOperationException
                || exception is FormatException) {
                diagnostic = "The producer SmartArt presentation graph could not be synchronized without changing diagram meaning: "
                    + exception.Message;
                return false;
            }
        }

        private static bool TrySynchronizeBasicProcessPresentationGraph(
            XDocument dataDocument,
            XDocument drawingDocument,
            (XNamespace dgm, XNamespace a) ns,
            PowerPointSmartArtTopology current,
            IReadOnlyList<PowerPointSmartArtNode> requested) {
            XNamespace dsp = PersistedDiagramNamespace;
            Dictionary<string, PowerPointSmartArtNode> currentById = current
                .Nodes.ToDictionary(node => node.Id, StringComparer.Ordinal);
            PowerPointSmartArtNode[] orderedRequested = requested
                .OrderBy(node => node.Order).ToArray();

            List<PersistedPresentationPoint> presentationPoints =
                ReadPersistedPresentationPoints(dataDocument, ns.dgm);
            PersistedPresentationPoint[] nodeSlots = presentationPoints
                .Where(point => string.Equals(point.Name, "node",
                    StringComparison.Ordinal))
                .OrderBy(point => point.Slot).ToArray();
            PersistedPresentationPoint[] transitionSlots = presentationPoints
                .Where(point => string.Equals(point.Name, "sibTrans",
                    StringComparison.Ordinal))
                .OrderBy(point => point.Slot).ToArray();
            if (nodeSlots.Length != orderedRequested.Length
                || transitionSlots.Length != Math.Max(0,
                    orderedRequested.Length - 1)
                || nodeSlots.Select(point => point.SourceId)
                    .Distinct(StringComparer.Ordinal).Count()
                    != nodeSlots.Length
                || !new HashSet<string>(nodeSlots.Select(point =>
                        point.SourceId), StringComparer.Ordinal)
                    .SetEquals(orderedRequested.Select(node => node.Id))) {
                return false;
            }

            Dictionary<string, PersistedPresentationPoint>
                nodePresentationBySemanticId = nodeSlots.ToDictionary(
                    point => point.SourceId, StringComparer.Ordinal);
            Dictionary<int, PersistedDrawingGeometry> geometryByPosition =
                new();
            for (int index = 0; index < nodeSlots.Length; index++) {
                XElement? shape = FindPersistedDrawingShape(drawingDocument,
                    dsp, nodeSlots[index].Id);
                XElement? transform = shape?.Element(dsp + "spPr")?
                    .Element(ns.a + "xfrm");
                XElement? textTransform = shape?.Element(dsp + "txXfrm");
                if (transform == null || textTransform == null) {
                    return false;
                }
                geometryByPosition.Add(index,
                    new PersistedDrawingGeometry(transform, textTransform));
            }

            Dictionary<string, XElement> semanticConnections = dataDocument
                .Descendants(ns.dgm + "cxn")
                .Where(connection => connection.Attribute("type") == null)
                .Where(connection => connection.Attribute("destId") != null)
                .GroupBy(connection => (string)connection.Attribute("destId")!,
                    StringComparer.Ordinal)
                .Where(group => group.Count() == 1)
                .ToDictionary(group => group.Key, group => group.Single(),
                    StringComparer.Ordinal);
            if (orderedRequested.Any(node =>
                    !semanticConnections.ContainsKey(node.Id))) {
                return false;
            }

            for (int index = 0; index < orderedRequested.Length; index++) {
                PowerPointSmartArtNode requestedNode = orderedRequested[index];
                PersistedPresentationPoint presentationPoint =
                    nodePresentationBySemanticId[requestedNode.Id];
                presentationPoint.ParentConnection.SetAttributeValue("srcOrd",
                    nodeSlots[index].Slot.ToString(CultureInfo.InvariantCulture));

                XElement? shape = FindPersistedDrawingShape(drawingDocument,
                    dsp, presentationPoint.Id);
                XElement? transform = shape?.Element(dsp + "spPr")?
                    .Element(ns.a + "xfrm");
                XElement? textTransform = shape?.Element(dsp + "txXfrm");
                if (shape == null || transform == null
                    || textTransform == null) {
                    return false;
                }
                transform.ReplaceWith(new XElement(
                    geometryByPosition[index].Transform));
                textTransform.ReplaceWith(new XElement(
                    geometryByPosition[index].TextTransform));

                if (!string.Equals(currentById[requestedNode.Id].Text,
                        requestedNode.Text, StringComparison.Ordinal)) {
                    XElement[] textRuns = shape.Descendants(ns.a + "t")
                        .ToArray();
                    if (textRuns.Length != 1
                        || shape.Descendants(ns.a + "br").Any()) {
                        return false;
                    }
                    textRuns[0].Value = requestedNode.Text;
                }

                if (index >= transitionSlots.Length) {
                    continue;
                }
                string? siblingTransitionId = (string?)semanticConnections[
                    requestedNode.Id].Attribute("sibTransId");
                if (string.IsNullOrWhiteSpace(siblingTransitionId)
                    || !TryRemapPresentationPoint(dataDocument, ns.dgm,
                        transitionSlots[index].Id, siblingTransitionId!)) {
                    return false;
                }
                foreach (PersistedPresentationPoint child in presentationPoints
                             .Where(point => string.Equals(point.ParentId,
                                 transitionSlots[index].Id,
                                 StringComparison.Ordinal))) {
                    if (!TryRemapPresentationPoint(dataDocument, ns.dgm,
                            child.Id, siblingTransitionId!)) {
                        return false;
                    }
                }
            }

            return true;
        }

        private static List<PersistedPresentationPoint>
            ReadPersistedPresentationPoints(XDocument dataDocument,
                XNamespace dgm) {
            var points = new List<PersistedPresentationPoint>();
            foreach (XElement point in dataDocument.Descendants(dgm + "pt")
                         .Where(point => string.Equals(
                             (string?)point.Attribute("type"), "pres",
                             StringComparison.Ordinal))) {
                string? id = (string?)point.Attribute("modelId");
                string? name = (string?)point.Element(dgm + "prSet")?
                    .Attribute("presName");
                if (string.IsNullOrWhiteSpace(id)
                    || string.IsNullOrWhiteSpace(name)) {
                    continue;
                }
                XElement[] parentConnections = dataDocument
                    .Descendants(dgm + "cxn")
                    .Where(connection => string.Equals(
                        (string?)connection.Attribute("type"), "presParOf",
                        StringComparison.Ordinal)
                        && string.Equals(
                            (string?)connection.Attribute("destId"), id,
                            StringComparison.Ordinal)).ToArray();
                XElement[] sourceConnections = dataDocument
                    .Descendants(dgm + "cxn")
                    .Where(connection => string.Equals(
                        (string?)connection.Attribute("type"), "presOf",
                        StringComparison.Ordinal)
                        && string.Equals(
                            (string?)connection.Attribute("destId"), id,
                            StringComparison.Ordinal)).ToArray();
                if (parentConnections.Length != 1
                    || sourceConnections.Length != 1
                    || !uint.TryParse((string?)parentConnections[0]
                            .Attribute("srcOrd"), NumberStyles.None,
                        CultureInfo.InvariantCulture, out uint slot)
                    || string.IsNullOrWhiteSpace((string?)parentConnections[0]
                        .Attribute("srcId"))
                    || string.IsNullOrWhiteSpace((string?)sourceConnections[0]
                        .Attribute("srcId"))) {
                    continue;
                }
                points.Add(new PersistedPresentationPoint(id!, name!, slot,
                    (string)parentConnections[0].Attribute("srcId")!,
                    (string)sourceConnections[0].Attribute("srcId")!,
                    parentConnections[0]));
            }
            return points;
        }

        private static bool TryRemapPresentationPoint(
            XDocument dataDocument, XNamespace dgm,
            string presentationPointId, string semanticPointId) {
            XElement[] mappings = dataDocument.Descendants(dgm + "cxn")
                .Where(connection => string.Equals(
                    (string?)connection.Attribute("type"), "presOf",
                    StringComparison.Ordinal)
                    && string.Equals((string?)connection.Attribute("destId"),
                        presentationPointId, StringComparison.Ordinal))
                .ToArray();
            if (mappings.Length != 1) {
                return false;
            }
            mappings[0].SetAttributeValue("srcId", semanticPointId);
            return true;
        }

        private static XElement? FindPersistedDrawingShape(
            XDocument drawingDocument, XNamespace dsp, string modelId) =>
            drawingDocument.Descendants(dsp + "sp").SingleOrDefault(shape =>
                string.Equals((string?)shape.Attribute("modelId"), modelId,
                    StringComparison.Ordinal));

        private static void SaveDiagramTopologyAtomically(
            DiagramDataPart dataPart, XDocument dataDocument,
            DiagramPersistLayoutPart persistPart,
            XDocument persistDocument) {
            byte[] originalData = ReadPartBytes(dataPart);
            byte[] originalPersistedDrawing = ReadPartBytes(persistPart);
            byte[] updatedData = SerializeXml(dataDocument);
            byte[] updatedPersistedDrawing = SerializeXml(persistDocument);
            try {
                WritePartBytes(dataPart, updatedData);
                WritePartBytes(persistPart, updatedPersistedDrawing);
            } catch {
                try { WritePartBytes(dataPart, originalData); } catch { }
                try {
                    WritePartBytes(persistPart, originalPersistedDrawing);
                } catch { }
                throw;
            }
        }

        private static byte[] ReadPartBytes(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open,
                FileAccess.Read);
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }

        private static byte[] SerializeXml(XDocument document) {
            using var buffer = new MemoryStream();
            document.Save(buffer);
            return buffer.ToArray();
        }

        private static void WritePartBytes(OpenXmlPart part, byte[] bytes) {
            using Stream stream = part.GetStream(FileMode.Create,
                FileAccess.Write);
            stream.Write(bytes, 0, bytes.Length);
        }

        private sealed class PersistedPresentationPoint {
            internal PersistedPresentationPoint(string id, string name,
                uint slot, string parentId, string sourceId,
                XElement parentConnection) {
                Id = id;
                Name = name;
                Slot = slot;
                ParentId = parentId;
                SourceId = sourceId;
                ParentConnection = parentConnection;
            }

            internal string Id { get; }
            internal string Name { get; }
            internal uint Slot { get; }
            internal string ParentId { get; }
            internal string SourceId { get; }
            internal XElement ParentConnection { get; }
        }

        private sealed class PersistedDrawingGeometry {
            internal PersistedDrawingGeometry(XElement transform,
                XElement textTransform) {
                Transform = new XElement(transform);
                TextTransform = new XElement(textTransform);
            }

            internal XElement Transform { get; }
            internal XElement TextTransform { get; }
        }
    }
}
