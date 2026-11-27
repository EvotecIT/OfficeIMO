using OfficeIMO.Core.Internal.OpenXml;
using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a SmartArt diagram on a PowerPoint slide.
    /// </summary>
    public partial class PowerPointSmartArt : PowerPointShape {
        private const string SimpleQuickStyleId =
            "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1";
        private const string AccentOneColorStyleId =
            "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2";
        private static readonly XNamespace DiagramNamespace =
            "http://schemas.openxmlformats.org/drawingml/2006/diagram";

        internal PowerPointSmartArt(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
        }

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
                    Width, Height, style);
                return true;
            } catch {
                snapshot = null!;
                return false;
            }
        }

        private bool HasRepresentableLayoutDefinition(OfficeDiagramKind kind,
            int nodeCount) {
            string? relationshipId = SingleRelationshipIds()
                .Attribute(OfficeXmlNamespaces.Relationships, "lo");
            if (string.IsNullOrWhiteSpace(relationshipId)
                || !TryGetDiagramPart(relationshipId, DiagramLayoutRelationshipType,
                    out PowerPointPartRef layoutPart)
                || ReadRootMarkup(layoutPart) is not string layoutMarkup
                || XElement.Parse(layoutMarkup) is not XElement layout
                || layout.Name != DiagramNamespace + "layoutDef"
                || Height.ToPoints() <= 0D) {
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
            double aspectRatio = Width.ToPoints() / Height.ToPoints();
            if (type == PowerPointSmartArtType.BasicList) {
                XElement? storedAspectRatio = layout.Descendants(DiagramNamespace + "param")
                    .FirstOrDefault(parameter => string.Equals(
                        (string?)parameter.Attribute("type"), "ar",
                        StringComparison.Ordinal));
                if (OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse((string?)storedAspectRatio?.Attribute("val"), NumberStyles.Float, CultureInfo.InvariantCulture, out double storedRatio)
                    && storedRatio > 0D && !double.IsInfinity(storedRatio)) {
                    aspectRatio = storedRatio;
                }
            }
            string expected = CanonicalLayoutDefinitionXml(type, nodeCount,
                aspectRatio);
            // The SDK writes the canonical layout exactly as built; a package tree writes the same markup its own way.
            if (string.Equals(layoutMarkup, expected, StringComparison.Ordinal)
                || XNode.DeepEquals(layout, XElement.Parse(expected))) {
                return true;
            }
            return type == PowerPointSmartArtType.BasicProcess
                && IsMicrosoftBasicProcessLayout(layout);
        }

        private static bool IsMicrosoftBasicProcessLayout(XElement layout) {
            const string diagramNamespace =
                "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            const string relationshipNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const string markupCompatibilityNamespace =
                "http://schemas.openxmlformats.org/markup-compatibility/2006";
            const string xmlNamespace =
                "http://www.w3.org/XML/1998/namespace";
            const string basicProcessLayoutId =
                "urn:microsoft.com/office/officeart/2005/8/layout/process1";
            if (!string.Equals((string?)layout.Attribute("uniqueId"),
                    basicProcessLayoutId, StringComparison.Ordinal)) {
                return false;
            }
            if (layout.DescendantsAndSelf().Any(element =>
                    !string.Equals(element.Name.NamespaceName, diagramNamespace,
                        StringComparison.Ordinal)
                    || element.Attributes().Any(attribute =>
                        !attribute.IsNamespaceDeclaration
                        && !string.IsNullOrEmpty(attribute.Name.NamespaceName)
                        && !string.Equals(attribute.Name.NamespaceName,
                            relationshipNamespace, StringComparison.Ordinal)
                        && !string.Equals(attribute.Name.NamespaceName,
                            markupCompatibilityNamespace,
                            StringComparison.Ordinal)
                        && !string.Equals(attribute.Name.NamespaceName,
                            xmlNamespace, StringComparison.Ordinal)))) {
                return false;
            }
            bool processCategory = layout.Descendants(DiagramNamespace + "cat")
                .Any(category => string.Equals((string?)category.Attribute("type"),
                    "process", StringComparison.OrdinalIgnoreCase));
            bool linearAlgorithm = layout.Descendants(DiagramNamespace + "alg")
                .Any(algorithm => string.Equals((string?)algorithm.Attribute("type"),
                    "lin", StringComparison.Ordinal));
            bool nodeIterator = layout.Descendants(DiagramNamespace + "forEach")
                .Any(iterator =>
                    ((string?)iterator.Attribute("axis") ?? string.Empty)
                        .Split(' ').Contains("ch")
                    && ((string?)iterator.Attribute("ptType") ?? string.Empty)
                        .Split(' ').Contains("node"));
            string?[] shapeTypes = layout.Descendants(DiagramNamespace + "shape")
                .Select(shape => (string?)shape.Attribute("type")).ToArray();
            bool nodeShape = shapeTypes.Length == 4
                && shapeTypes[0] == null
                && string.Equals(shapeTypes[1], "roundRect",
                    StringComparison.Ordinal)
                && string.Equals(shapeTypes[2], "conn",
                    StringComparison.Ordinal)
                && string.Equals(shapeTypes[3], "conn",
                    StringComparison.Ordinal);
            return processCategory && linearAlgorithm && nodeIterator
                && nodeShape;
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

            PowerPointThemeScope theme = PowerPointThemeScope.For(HostPartRef);
            OfficeColor? accent = OfficeXmlThemeColors
                .ResolveSchemeColor(theme.ColorScheme, "accent1");
            OfficeColor? light = OfficeXmlThemeColors
                .ResolveSchemeColor(theme.ColorScheme, "light1");
            string? fontFamily = theme.FontScheme
                .Child(OfficeXmlNamespaces.DrawingML, "minorFont")
                .Child(OfficeXmlNamespaces.DrawingML, "latin")
                .Attribute("typeface");
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
            OfficeXmlNode relationshipIds = SingleRelationshipIds();
            string? styleRelationshipId = relationshipIds
                .Attribute(OfficeXmlNamespaces.Relationships, "qs");
            string? colorRelationshipId = relationshipIds
                .Attribute(OfficeXmlNamespaces.Relationships, "cs");
            if (string.IsNullOrWhiteSpace(styleRelationshipId)
                || string.IsNullOrWhiteSpace(colorRelationshipId)
                || !TryGetDiagramPart(styleRelationshipId,
                    DiagramQuickStyleRelationshipType, out PowerPointPartRef stylePart)
                || ReadRootMarkup(stylePart) is not string styleDefinition
                || !TryGetDiagramPart(colorRelationshipId,
                    DiagramColorsRelationshipType, out PowerPointPartRef colorsPart)
                || ReadRootMarkup(colorsPart) is not string colorsDefinition) {
                return false;
            }
            return MatchesSmartArtDefinition(styleDefinition,
                    PowerPointSlide.SmartArtStyleDefinitionXml)
                && MatchesSmartArtDefinition(colorsDefinition,
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

            var textBodySet = new HashSet<XElement>(textBodies);

            List<XElement> nodePoints = xdoc.Descendants(ns.dgm + "pt")
                .Where(point => point.Attribute("type") == null)
                .Where(point => {
                    XElement? body = point.Element(ns.dgm + "t")
                        ?? point.Element(ns.dgm + "txBody");
                    return body != null && textBodySet.Contains(body);
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
            var childrenByParent = new Dictionary<string, List<KeyValuePair<string,
                (int Index, XElement TextBody)>>>(StringComparer.Ordinal);
            var rootNodes = new List<KeyValuePair<string,
                (int Index, XElement TextBody)>>();

            foreach (KeyValuePair<string, (int Index, XElement TextBody)> node in nodes) {
                if (!parentByNode.TryGetValue(node.Key, out string? parent)) continue;
                if (documentIds.Contains(parent)) {
                    rootNodes.Add(node);
                    continue;
                }
                if (!childrenByParent.TryGetValue(parent, out List<KeyValuePair<string,
                        (int Index, XElement TextBody)>>? children)) {
                    children = new List<KeyValuePair<string,
                        (int Index, XElement TextBody)>>();
                    childrenByParent.Add(parent, children);
                }
                children.Add(node);
            }

            void AppendNodes(IEnumerable<KeyValuePair<string,
                (int Index, XElement TextBody)>> startingNodes) {
                var stack = new Stack<KeyValuePair<string,
                    (int Index, XElement TextBody)>>();
                KeyValuePair<string, (int Index, XElement TextBody)>[] orderedStarts =
                    OrderSiblingNodes(startingNodes, sourceOrderByNode).ToArray();
                for (int index = orderedStarts.Length - 1; index >= 0; index--) {
                    stack.Push(orderedStarts[index]);
                }

                while (stack.Count > 0) {
                    KeyValuePair<string, (int Index, XElement TextBody)> node =
                        stack.Pop();
                    if (!visited.Add(node.Key)) continue;
                    ordered.Add(node);
                    if (!childrenByParent.TryGetValue(node.Key,
                            out List<KeyValuePair<string,
                                (int Index, XElement TextBody)>>? children)) continue;
                    KeyValuePair<string, (int Index, XElement TextBody)>[] orderedChildren =
                        OrderSiblingNodes(children, sourceOrderByNode).ToArray();
                    for (int index = orderedChildren.Length - 1; index >= 0; index--) {
                        stack.Push(orderedChildren[index]);
                    }
                }
            }

            AppendNodes(rootNodes);
            AppendNodes(nodes.Where(node => !visited.Contains(node.Key)));
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
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, textBodies.Count);

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
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, textBodies.Count);

            bool updateCachedDrawing = TryPrepareCachedNodeTextEdit(xdoc, textBodies[index], normalizedText,
                out PowerPointPartRef persistPart, out XDocument? persistDocument);
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
            if (updateCachedDrawing) SaveDiagramTopologyAtomically(dataPart, xdoc, persistPart, persistDocument!);
            else SaveDiagramData(dataPart, xdoc);
        }

        private (XDocument xdoc, (XNamespace dgm, XNamespace a) ns,
            List<XElement> textBodies, PowerPointPartRef dataPart)
            LoadNodeTextBodiesWithPart() {
            PowerPointPartRef dataPart = GetDiagramDataPart();
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

        private PowerPointPartRef GetDiagramDataPart() {
            OfficeXmlNode relationshipIds = GraphicDataRelationshipIds;
            if (relationshipIds.IsNone) {
                throw new InvalidOperationException("SmartArt relationship ids were not found.");
            }
            string? dataPartId = relationshipIds.Attribute(OfficeXmlNamespaces.Relationships, "dm");
            if (string.IsNullOrWhiteSpace(dataPartId)) {
                throw new InvalidOperationException("SmartArt data relationship was not found.");
            }
            // An id the slide has no relationship for is an argument error, as the SDK reports it.
            if (!HostPartRef.TryGetRelationship(dataPartId!, out _)) {
                throw new ArgumentOutOfRangeException(nameof(dataPartId), "The slide has no relationship with the SmartArt data id.");
            }
            return TryGetDiagramPart(dataPartId, DiagramDataRelationshipType, out PowerPointPartRef dataPart)
                ? dataPart
                : throw new InvalidOperationException("SmartArt diagram data part was not found.");
        }

        private static XDocument LoadDiagramXDocument(PowerPointPartRef dataPart) => LoadPartXml(dataPart);

        private static void SaveDiagramData(PowerPointPartRef dataPart, XDocument xdoc) => WritePartXml(dataPart, xdoc);
    }
}
