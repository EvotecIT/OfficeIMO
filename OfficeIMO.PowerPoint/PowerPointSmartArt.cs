using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a SmartArt diagram on a PowerPoint slide.
    /// </summary>
    public class PowerPointSmartArt : PowerPointShape {
        private readonly SlidePart _slidePart;

        internal PowerPointSmartArt(GraphicFrame graphicFrame, SlidePart slidePart) : base(graphicFrame) {
            _slidePart = slidePart;
        }

        private GraphicFrame GraphicFrame => (GraphicFrame)Element;

        /// <summary>
        ///     Gets the number of editable SmartArt nodes.
        /// </summary>
        public int NodeCount => LoadNodeTextBodiesWithPart().textBodies.Count;

        /// <summary>Gets all editable SmartArt node texts in data-model order.</summary>
        public IReadOnlyList<string> GetNodeTexts() {
            var (_, ns, textBodies, _) = LoadNodeTextBodiesWithPart();
            return textBodies.Select(body => ReadNodeText(body, ns.a))
                .ToList().AsReadOnly();
        }

        /// <summary>
        /// Tries to expose the current SmartArt content through the shared
        /// dependency-free semantic diagram contract.
        /// </summary>
        public bool TryGetOfficeDiagramSnapshot(
            out OfficeDiagramSnapshot snapshot) {
            try {
                var (xdoc, ns, textBodies, _) = LoadNodeTextBodiesWithPart();
                IReadOnlyList<string> nodes = textBodies
                    .Select(body => ReadNodeText(body, ns.a))
                    .Where(text => !string.IsNullOrWhiteSpace(text))
                    .ToArray();
                XElement? properties = xdoc.Descendants(ns.dgm + "prSet")
                    .FirstOrDefault(element =>
                        element.Attribute("loCatId") != null
                        || element.Attribute("loTypeId") != null);
                string category = ((string?)properties?.Attribute("loCatId")
                    ?? (string?)properties?.Attribute("loTypeId")
                    ?? string.Empty).ToLowerInvariant();
                OfficeDiagramKind kind;
                if (category.IndexOf("hierarchy", StringComparison.Ordinal) >= 0) {
                    kind = OfficeDiagramKind.Hierarchy;
                } else if (category.IndexOf("cycle", StringComparison.Ordinal) >= 0) {
                    kind = OfficeDiagramKind.Cycle;
                } else if (category.IndexOf("matrix", StringComparison.Ordinal) >= 0) {
                    kind = OfficeDiagramKind.Matrix;
                } else if (category.IndexOf("pyramid", StringComparison.Ordinal) >= 0) {
                    kind = OfficeDiagramKind.Pyramid;
                } else if (category.IndexOf("relationship", StringComparison.Ordinal) >= 0) {
                    kind = OfficeDiagramKind.Relationship;
                } else if (category.IndexOf("list", StringComparison.Ordinal) >= 0) {
                    kind = OfficeDiagramKind.List;
                } else {
                    kind = OfficeDiagramKind.Process;
                }
                snapshot = new OfficeDiagramSnapshot(Name, kind, nodes,
                    WidthPoints, HeightPoints);
                return true;
            } catch {
                snapshot = null!;
                return false;
            }
        }

        /// <summary>
        ///     Gets the text of an editable SmartArt node.
        /// </summary>
        public string GetNodeText(int index) {
            var (_, ns, textBodies, _) = LoadNodeTextBodiesWithPart();
            if (index < 0 || index >= textBodies.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return ReadNodeText(textBodies[index], ns.a);
        }

        /// <summary>
        ///     Replaces the text of an editable SmartArt node.
        /// </summary>
        public void SetNodeText(int index, string text) {
            var (xdoc, ns, textBodies, dataPart) =
                LoadNodeTextBodiesWithPart();
            if (index < 0 || index >= textBodies.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            List<XElement> paragraphs = textBodies[index].Elements(ns.a + "p")
                .ToList();
            XElement paragraph = paragraphs[0];
            paragraph.RemoveNodes();
            paragraph.Add(new XElement(ns.a + "r",
                new XElement(ns.a + "t", text ?? string.Empty)));
            paragraph.Add(new XElement(ns.a + "endParaRPr", new XAttribute("lang", "en-US")));
            for (int paragraphIndex = 1;
                 paragraphIndex < paragraphs.Count; paragraphIndex++) {
                paragraphs[paragraphIndex].Remove();
            }
            SaveDiagramData(dataPart, xdoc);
        }

        private (XDocument xdoc, (XNamespace dgm, XNamespace a) ns,
            List<XElement> textBodies, DiagramDataPart dataPart)
            LoadNodeTextBodiesWithPart() {
            DiagramDataPart dataPart = GetDiagramDataPart();
            XDocument xdoc = LoadDiagramXDocument(dataPart);
            XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            List<XElement> textBodies = xdoc
                .Descendants(dgm + "pt")
                .Where(point => point.Attribute("type") == null)
                .Select(point => point.Element(dgm + "t")
                    ?? point.Element(dgm + "txBody"))
                .Where(body => body?.Elements(a + "p").Any() == true)
                .Cast<XElement>()
                .ToList();

            return (xdoc, (dgm, a), textBodies, dataPart);
        }

        private static string ReadNodeText(XElement textBody, XNamespace a) =>
            string.Join("\n", textBody.Elements(a + "p")
                .Select(paragraph => string.Concat(paragraph
                    .Descendants(a + "t")
                    .Select(text => (string?)text ?? string.Empty))));

        private DiagramDataPart GetDiagramDataPart() {
            Dgm.RelationshipIds relationshipIds = GraphicFrame.Graphic?.GraphicData?.GetFirstChild<Dgm.RelationshipIds>()
                ?? throw new InvalidOperationException("SmartArt relationship ids were not found.");
            string? dataPartId = relationshipIds.DataPart?.Value;
            if (string.IsNullOrWhiteSpace(dataPartId)) {
                throw new InvalidOperationException("SmartArt data relationship was not found.");
            }

            return _slidePart.GetPartById(dataPartId!) as DiagramDataPart
                ?? throw new InvalidOperationException("SmartArt diagram data part was not found.");
        }

        private static XDocument LoadDiagramXDocument(DiagramDataPart dataPart) {
            using Stream stream = dataPart.GetStream(FileMode.Open, FileAccess.Read);
            return PowerPointXmlReader.LoadPackagePartXml(stream);
        }

        private static void SaveDiagramData(DiagramDataPart dataPart, XDocument xdoc) {
            using Stream stream = dataPart.GetStream(FileMode.Create, FileAccess.Write);
            xdoc.Save(stream);
        }
    }
}
