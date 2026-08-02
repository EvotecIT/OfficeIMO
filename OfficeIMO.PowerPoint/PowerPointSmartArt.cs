using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a SmartArt diagram on a PowerPoint slide.
    /// </summary>
    public class PowerPointSmartArt : PowerPointShape {
        private const string SimpleQuickStyleId =
            "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1";
        private const string AccentOneColorStyleId =
            "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2";
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
                XElement? properties = xdoc.Descendants(ns.dgm + "prSet")
                    .FirstOrDefault(element =>
                        element.Attribute("loCatId") != null
                        || element.Attribute("loTypeId") != null);
                string category = ((string?)properties?.Attribute("loCatId")
                    ?? (string?)properties?.Attribute("loTypeId")
                    ?? string.Empty).ToLowerInvariant();
                if (!TryResolveDiagramKind(category, out OfficeDiagramKind kind)) {
                    snapshot = null!;
                    return false;
                }
                if (!TryReadRepresentableTopology(xdoc, ns, textBodies,
                        kind, out IReadOnlyList<string> nodes)) {
                    snapshot = null!;
                    return false;
                }
                if (!TryReadRepresentableStyle(properties,
                        out OfficeDiagramStyle style)) {
                    snapshot = null!;
                    return false;
                }
                if (!HasRepresentableLayoutDefinition(kind,
                        textBodies.Count)) {
                    snapshot = null!;
                    return false;
                }
                snapshot = new OfficeDiagramSnapshot(Name, kind, nodes,
                    WidthPoints, HeightPoints, style);
                return true;
            } catch {
                snapshot = null!;
                return false;
            }
        }

        private bool HasRepresentableLayoutDefinition(OfficeDiagramKind kind,
            int nodeCount) {
            Dgm.RelationshipIds? relationshipIds = GraphicFrame
                .Descendants<Dgm.RelationshipIds>().SingleOrDefault();
            string? relationshipId = relationshipIds?.LayoutPart?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)
                || !_slidePart.TryGetPartById(relationshipId!,
                    out OpenXmlPart? relatedPart)
                || relatedPart is not DiagramLayoutDefinitionPart layoutPart
                || layoutPart.LayoutDefinition == null
                || HeightPoints <= 0D) {
                return false;
            }
            PowerPointSmartArtType type = kind switch {
                OfficeDiagramKind.Process => PowerPointSmartArtType.BasicProcess,
                OfficeDiagramKind.Hierarchy => PowerPointSmartArtType.BasicHierarchy,
                OfficeDiagramKind.Cycle => PowerPointSmartArtType.BasicCycle,
                OfficeDiagramKind.List => PowerPointSmartArtType.BasicList,
                OfficeDiagramKind.Matrix => PowerPointSmartArtType.BasicMatrix,
                OfficeDiagramKind.Pyramid => PowerPointSmartArtType.BasicPyramid,
                OfficeDiagramKind.Relationship => PowerPointSmartArtType.BasicRelationship,
                _ => throw new ArgumentOutOfRangeException(nameof(kind))
            };
            Dgm.LayoutDefinition expected = PowerPointSlide
                .CreateSmartArtLayoutDefinition(type, nodeCount,
                    WidthPoints / HeightPoints);
            return string.Equals(layoutPart.LayoutDefinition.OuterXml,
                expected.OuterXml, StringComparison.Ordinal);
        }

        private bool TryReadRepresentableStyle(XElement? properties,
            out OfficeDiagramStyle style) {
            style = null!;
            if (!string.Equals((string?)properties?.Attribute("qsTypeId"),
                    SimpleQuickStyleId, StringComparison.Ordinal)
                || !string.Equals((string?)properties?.Attribute("csTypeId"),
                    AccentOneColorStyleId, StringComparison.Ordinal)) {
                return false;
            }
            if (!HasRepresentableStyleDefinitions()) return false;

            A.ColorScheme? colorScheme = _slidePart.ThemeOverridePart?
                    .ThemeOverride?.ColorScheme
                ?? _slidePart.SlideLayoutPart?.ThemeOverridePart?
                    .ThemeOverride?.ColorScheme
                ?? _slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?
                    .Theme?.ThemeElements?.ColorScheme;
            A.FontScheme? fontScheme = _slidePart.ThemeOverridePart?
                    .ThemeOverride?.FontScheme
                ?? _slidePart.SlideLayoutPart?.ThemeOverridePart?
                    .ThemeOverride?.FontScheme
                ?? _slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?
                    .Theme?.ThemeElements?.FontScheme;
            OfficeColor? accent = OfficeOpenXmlThemeColorResolver
                .ResolveSchemeColor(colorScheme, "accent1");
            OfficeColor? light = OfficeOpenXmlThemeColorResolver
                .ResolveSchemeColor(colorScheme, "light1");
            string? fontFamily = fontScheme?.MinorFont?.LatinFont?
                .Typeface?.Value;
            if (!accent.HasValue || !light.HasValue
                || string.IsNullOrWhiteSpace(fontFamily)) {
                return false;
            }

            style = new OfficeDiagramStyle(fontFamily!,
                new[] { accent.Value }, light.Value, light.Value,
                accent.Value);
            return true;
        }

        private bool HasRepresentableStyleDefinitions() {
            Dgm.RelationshipIds? relationshipIds = GraphicFrame
                .Descendants<Dgm.RelationshipIds>().SingleOrDefault();
            string? styleRelationshipId = relationshipIds?.StylePart?.Value;
            string? colorRelationshipId = relationshipIds?.ColorPart?.Value;
            if (string.IsNullOrWhiteSpace(styleRelationshipId)
                || string.IsNullOrWhiteSpace(colorRelationshipId)
                || !_slidePart.TryGetPartById(styleRelationshipId!,
                    out OpenXmlPart? relatedStyle)
                || relatedStyle is not DiagramStylePart stylePart
                || stylePart.StyleDefinition == null
                || !_slidePart.TryGetPartById(colorRelationshipId!,
                    out OpenXmlPart? relatedColors)
                || relatedColors is not DiagramColorsPart colorsPart
                || colorsPart.ColorsDefinition == null) {
                return false;
            }
            return MatchesSmartArtDefinition(
                    stylePart.StyleDefinition.OuterXml,
                    PowerPointSlide.SmartArtStyleDefinitionXml)
                && MatchesSmartArtDefinition(
                    colorsPart.ColorsDefinition.OuterXml,
                    PowerPointSlide.SmartArtColorsDefinitionXml);
        }

        private static bool MatchesSmartArtDefinition(string actual,
            string expected) => XNode.DeepEquals(
                NormalizeSmartArtDefinition(XElement.Parse(actual)),
                NormalizeSmartArtDefinition(XElement.Parse(expected)));

        private static XElement NormalizeSmartArtDefinition(
            XElement element) {
            var normalized = new XElement(element.Name);
            foreach (XAttribute attribute in element.Attributes()
                         .Where(attribute => !attribute.IsNamespaceDeclaration)
                         .OrderBy(attribute => attribute.Name.NamespaceName,
                             StringComparer.Ordinal)
                         .ThenBy(attribute => attribute.Name.LocalName,
                             StringComparer.Ordinal)) {
                normalized.Add(new XAttribute(attribute.Name,
                    attribute.Value));
            }
            foreach (XNode node in element.Nodes()) {
                if (node is XElement child) {
                    normalized.Add(NormalizeSmartArtDefinition(child));
                } else if (node is XText text
                           && !string.IsNullOrWhiteSpace(text.Value)) {
                    normalized.Add(new XText(text.Value));
                }
            }
            return normalized;
        }

        private static bool TryResolveDiagramKind(string category,
            out OfficeDiagramKind kind) {
            category = category.Trim();
            if (category == "hierarchy"
                || category.EndsWith("/layout/hierarchy1", StringComparison.Ordinal)) {
                kind = OfficeDiagramKind.Hierarchy;
            } else if (category == "cycle"
                || category.EndsWith("/layout/cycle2", StringComparison.Ordinal)) {
                kind = OfficeDiagramKind.Cycle;
            } else if (category == "matrix"
                || category.EndsWith("/layout/matrix3", StringComparison.Ordinal)) {
                kind = OfficeDiagramKind.Matrix;
            } else if (category == "pyramid"
                || category.EndsWith("/layout/pyramid1", StringComparison.Ordinal)) {
                kind = OfficeDiagramKind.Pyramid;
            } else if (category == "relationship"
                || category.EndsWith("/layout/radial1", StringComparison.Ordinal)) {
                kind = OfficeDiagramKind.Relationship;
            } else if (category == "list"
                || category.EndsWith("/layout/default", StringComparison.Ordinal)) {
                kind = OfficeDiagramKind.List;
            } else if (category == "process"
                || category.EndsWith("/layout/process1", StringComparison.Ordinal)) {
                kind = OfficeDiagramKind.Process;
            } else {
                kind = default;
                return false;
            }
            return true;
        }

        private static bool TryReadRepresentableTopology(
            XDocument xdoc,
            (XNamespace dgm, XNamespace a) ns,
            IReadOnlyList<XElement> textBodies,
            OfficeDiagramKind kind,
            out IReadOnlyList<string> nodes) {
            nodes = Array.Empty<string>();
            if (!TryCreateSemanticNodeMap(xdoc, ns, textBodies,
                    out Dictionary<string, (int Index, XElement TextBody)> nodeById,
                    out HashSet<string> documentIds,
                    out Dictionary<string, string> parentByNode,
                    out Dictionary<string, uint> sourceOrderByNode)) {
                return false;
            }
            if (nodeById.Values.Any(node => string.IsNullOrWhiteSpace(
                    ReadNodeText(node.TextBody, ns.a)))) return false;
            if (parentByNode.Count != nodeById.Count) return false;

            bool rooted = kind == OfficeDiagramKind.Hierarchy
                || kind == OfficeDiagramKind.Relationship;
            if (!rooted) {
                if (parentByNode.Values.Any(parent =>
                        !documentIds.Contains(parent))) {
                    return false;
                }
                nodes = OrderSemanticNodes(nodeById, parentByNode,
                        sourceOrderByNode, documentIds)
                    .Select(node => ReadNodeText(node.Value.TextBody, ns.a))
                    .ToArray();
                return true;
            }

            string[] roots = parentByNode
                .Where(pair => documentIds.Contains(pair.Value))
                .Select(pair => pair.Key)
                .ToArray();
            if (roots.Length != 1) return false;
            string rootId = roots[0];
            if (parentByNode.Any(pair => pair.Key != rootId
                    && !string.Equals(pair.Value, rootId,
                        StringComparison.Ordinal))) {
                return false;
            }
            nodes = OrderSemanticNodes(nodeById, parentByNode,
                    sourceOrderByNode, documentIds)
                .Select(node => ReadNodeText(node.Value.TextBody, ns.a))
                .ToArray();
            return true;
        }

        private static bool TryCreateSemanticNodeMap(
            XDocument xdoc,
            (XNamespace dgm, XNamespace a) ns,
            IReadOnlyList<XElement> textBodies,
            out Dictionary<string, (int Index, XElement TextBody)> nodeById,
            out HashSet<string> documentIds,
            out Dictionary<string, string> parentByNode,
            out Dictionary<string, uint> sourceOrderByNode) {
            nodeById = new Dictionary<string, (int, XElement)>(
                StringComparer.Ordinal);
            documentIds = new HashSet<string>(StringComparer.Ordinal);
            parentByNode = new Dictionary<string, string>(StringComparer.Ordinal);
            sourceOrderByNode = new Dictionary<string, uint>(StringComparer.Ordinal);

            List<XElement> nodePoints = xdoc.Descendants(ns.dgm + "pt")
                .Where(point => point.Attribute("type") == null)
                .Where(point => {
                    XElement? body = point.Element(ns.dgm + "t")
                        ?? point.Element(ns.dgm + "txBody");
                    return body != null && textBodies.Contains(body);
                })
                .ToList();
            if (nodePoints.Count == 0 || nodePoints.Count != textBodies.Count) {
                return false;
            }

            for (int index = 0; index < nodePoints.Count; index++) {
                string? modelId = (string?)nodePoints[index].Attribute("modelId");
                XElement? textBody = nodePoints[index].Element(ns.dgm + "t")
                    ?? nodePoints[index].Element(ns.dgm + "txBody");
                if (string.IsNullOrWhiteSpace(modelId) || textBody == null
                    || nodeById.ContainsKey(modelId!)) {
                    return false;
                }
                nodeById.Add(modelId!, (index, textBody));
            }

            documentIds.UnionWith(xdoc.Descendants(ns.dgm + "pt")
                .Where(point => string.Equals((string?)point.Attribute("type"),
                    "doc", StringComparison.OrdinalIgnoreCase))
                .Select(point => (string?)point.Attribute("modelId"))
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Cast<string>());
            if (documentIds.Count == 0) return false;

            foreach (XElement connection in xdoc.Descendants(ns.dgm + "cxn")) {
                string? type = (string?)connection.Attribute("type");
                if (!string.IsNullOrWhiteSpace(type)
                    && !string.Equals(type, "parOf",
                        StringComparison.OrdinalIgnoreCase)) continue;

                string? destination = (string?)connection.Attribute("destId");
                if (string.IsNullOrWhiteSpace(destination)
                    || !nodeById.ContainsKey(destination!)) continue;

                string? source = (string?)connection.Attribute("srcId");
                if (string.IsNullOrWhiteSpace(source)
                    || (!documentIds.Contains(source!)
                        && !nodeById.ContainsKey(source!))
                    || parentByNode.ContainsKey(destination!)) return false;

                parentByNode.Add(destination!, source!);
                string? sourceOrder = (string?)connection.Attribute("srcOrd");
                if (string.IsNullOrWhiteSpace(sourceOrder)) continue;
                if (!uint.TryParse(sourceOrder, NumberStyles.None,
                        CultureInfo.InvariantCulture, out uint parsedOrder)) {
                    return false;
                }
                sourceOrderByNode.Add(destination!, parsedOrder);
            }
            return true;
        }

        private static IReadOnlyList<KeyValuePair<string,
            (int Index, XElement TextBody)>> OrderSemanticNodes(
            IReadOnlyDictionary<string, (int Index, XElement TextBody)> nodes,
            IReadOnlyDictionary<string, string> parentByNode,
            IReadOnlyDictionary<string, uint> sourceOrderByNode,
            IReadOnlyCollection<string> documentIds) {
            var ordered = new List<KeyValuePair<string,
                (int Index, XElement TextBody)>>();
            var visited = new HashSet<string>(StringComparer.Ordinal);

            void AppendNode(KeyValuePair<string,
                (int Index, XElement TextBody)> node) {
                if (!visited.Add(node.Key)) return;
                ordered.Add(node);
                foreach (KeyValuePair<string, (int Index, XElement TextBody)> child
                         in OrderSiblingNodes(nodes.Where(candidate =>
                             parentByNode.TryGetValue(candidate.Key,
                                 out string? parent)
                             && string.Equals(parent, node.Key,
                                 StringComparison.Ordinal)), sourceOrderByNode)) {
                    AppendNode(child);
                }
            }

            foreach (KeyValuePair<string, (int Index, XElement TextBody)> root
                     in OrderSiblingNodes(nodes.Where(node =>
                         parentByNode.TryGetValue(node.Key, out string? parent)
                         && documentIds.Contains(parent)), sourceOrderByNode)) {
                AppendNode(root);
            }
            foreach (KeyValuePair<string, (int Index, XElement TextBody)> node
                     in OrderSiblingNodes(nodes.Where(node =>
                         !visited.Contains(node.Key)), sourceOrderByNode)) {
                AppendNode(node);
            }
            return ordered;
        }

        private static IOrderedEnumerable<KeyValuePair<string,
            (int Index, XElement TextBody)>> OrderSiblingNodes(
            IEnumerable<KeyValuePair<string,
                (int Index, XElement TextBody)>> nodes,
            IReadOnlyDictionary<string, uint> sourceOrderByNode) =>
            nodes.OrderBy(node => sourceOrderByNode.ContainsKey(node.Key)
                    ? 0
                    : 1)
                .ThenBy(node => sourceOrderByNode.TryGetValue(node.Key,
                    out uint sourceOrder) ? sourceOrder : uint.MaxValue)
                .ThenBy(node => node.Value.Index);

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
            string normalizedText = text ?? string.Empty;
            PowerPointXmlValueValidator.ValidateCharacters(normalizedText,
                nameof(text), "SmartArt node text");
            if (string.IsNullOrWhiteSpace(normalizedText)) {
                throw new ArgumentException(
                    "SmartArt node text cannot be empty.", nameof(text));
            }
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
                new XElement(ns.a + "t", normalizedText)));
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

            if (TryCreateSemanticNodeMap(xdoc, (dgm, a), textBodies,
                    out Dictionary<string, (int Index, XElement TextBody)> nodeById,
                    out HashSet<string> documentIds,
                    out Dictionary<string, string> parentByNode,
                    out Dictionary<string, uint> sourceOrderByNode)) {
                textBodies = OrderSemanticNodes(nodeById, parentByNode,
                        sourceOrderByNode, documentIds)
                    .Select(node => node.Value.TextBody)
                    .ToList();
            }

            return (xdoc, (dgm, a), textBodies, dataPart);
        }

        private static string ReadNodeText(XElement textBody, XNamespace a) =>
            string.Join("\n", textBody.Elements(a + "p")
                .Select(paragraph => string.Concat(paragraph.Descendants()
                    .Where(element => element.Name == a + "t"
                        || element.Name == a + "br")
                    .Select(element => element.Name == a + "br"
                        ? "\n"
                        : (string?)element ?? string.Empty))));

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
