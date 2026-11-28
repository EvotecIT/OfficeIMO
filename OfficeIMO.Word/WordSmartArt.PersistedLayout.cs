using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    public partial class WordSmartArt {
        private const string PersistedDrawingNamespace = "http://schemas.microsoft.com/office/drawing/2008/diagram";

        internal IReadOnlyList<WordSmartArtPersistedShape> GetPersistedLayoutShapes() {
            try {
                object dataPart = GetDiagramDataPart();
                XDocument data = LoadDiagramXDocument(dataPart);
                string? persistRelationshipId = GetPersistedLayoutRelationshipId(data);
                if (string.IsNullOrWhiteSpace(persistRelationshipId)) {
                    return Array.Empty<WordSmartArtPersistedShape>();
                }

                object owner = WordPartOwnership.ResolveOwner(_document, _drawing);
                OfficeXmlNode drawing = WordPartOwnership.GetDiagramPersistedLayout(owner, persistRelationshipId!);
                if (drawing.IsNone) {
                    return Array.Empty<WordSmartArtPersistedShape>();
                }

                Dictionary<string, string> nodeTextById = GetNodeTextByModelId(data);
                Dictionary<string, string> sourceByPresentationId = GetPresentationSourceByDestinationId(data);
                var shapes = new List<WordSmartArtPersistedShape>();
                foreach (OfficeXmlNode shape in PersistedShapes(drawing)) {
                    string? modelId = shape.Attribute("modelId");
                    if (string.IsNullOrWhiteSpace(modelId)) {
                        continue;
                    }

                    string shapeModelId = modelId!;

                    OfficeXmlNode shapeProperties = shape.Child(PersistedDrawingNamespace, "spPr");
                    OfficeXmlNode transform = shapeProperties.IsNone ? default : shapeProperties.Child(OfficeXmlNamespaces.DrawingML, "xfrm");
                    OfficeXmlNode offset = transform.IsNone ? default : transform.Child(OfficeXmlNamespaces.DrawingML, "off");
                    OfficeXmlNode extents = transform.IsNone ? default : transform.Child(OfficeXmlNamespaces.DrawingML, "ext");
                    long? x = ReadInt64(offset, "x");
                    long? y = x == null ? null : ReadInt64(offset, "y");
                    long? cx = y == null ? null : ReadInt64(extents, "cx");
                    long? cy = cx == null ? null : ReadInt64(extents, "cy");
                    if (x == null || y == null || cx == null || cy == null || cx.Value <= 0L || cy.Value <= 0L) {
                        continue;
                    }

                    OfficeXmlNode geometry = shapeProperties.IsNone ? default : shapeProperties.Child(OfficeXmlNamespaces.DrawingML, "prstGeom");
                    string preset = (geometry.IsNone ? null : geometry.Attribute("prst")) ?? string.Empty;
                    if (preset.Length == 0) {
                        continue;
                    }

                    string text = string.Empty;
                    if (nodeTextById.TryGetValue(shapeModelId, out string? directText)) {
                        text = directText;
                    } else if (sourceByPresentationId.TryGetValue(shapeModelId, out string? sourceModelId) &&
                        nodeTextById.TryGetValue(sourceModelId, out string? sourceText)) {
                        text = sourceText;
                    }

                    string? rotationText = transform.IsNone ? null : transform.Attribute("rot");
                    double rotation = rotationText != null
                        ? XmlConvert.ToInt32(rotationText) / 60000D
                        : 0D;

                    shapes.Add(new WordSmartArtPersistedShape(
                        shapeModelId,
                        preset,
                        text ?? string.Empty,
                        Helpers.ConvertEmusToPoints(x.Value),
                        Helpers.ConvertEmusToPoints(y.Value),
                        Helpers.ConvertEmusToPoints(cx.Value),
                        Helpers.ConvertEmusToPoints(cy.Value),
                        rotation));
                }

                return shapes;
            } catch {
                return Array.Empty<WordSmartArtPersistedShape>();
            }
        }

        // An integer attribute as the SDK's typed value reads it: null when absent, a FormatException when malformed.
        private static long? ReadInt64(OfficeXmlNode element, string localName) {
            string? value = element.IsNone ? null : element.Attribute(localName);
            return value == null ? null : XmlConvert.ToInt64(value);
        }

        /// <summary>The <c>dsp:sp</c> shapes of a persisted layout drawing, in document order.</summary>
        /// <summary>
        /// The shapes of a persisted layout drawing in document order: the <c>dsp:sp</c> children of its shape tree and
        /// of nested groups. Extension lists and other markup are not searched, as a shape there is not part of the layout.
        /// </summary>
        private static List<OfficeXmlNode> PersistedShapes(OfficeXmlNode root) {
            var shapes = new List<OfficeXmlNode>();
            OfficeXmlNode tree = root.Child(PersistedDrawingNamespace, "spTree");
            if (!tree.IsNone) CollectPersistedShapes(tree, shapes);
            return shapes;
        }

        private static void CollectPersistedShapes(OfficeXmlNode container, List<OfficeXmlNode> shapes) {
            for (OfficeXmlNode child = container.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.Is(PersistedDrawingNamespace, "sp")) shapes.Add(child);
                else if (child.Is(PersistedDrawingNamespace, "grpSp")) CollectPersistedShapes(child, shapes);
            }
        }

        private static string? GetPersistedLayoutRelationshipId(XDocument data) {
            return data.Descendants()
                .SelectMany(element => element.Attributes())
                .FirstOrDefault(attribute => string.Equals(attribute.Name.LocalName, "relId", StringComparison.OrdinalIgnoreCase))
                ?.Value;
        }

        private static Dictionary<string, string> GetNodeTextByModelId(XDocument data) {
            XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement point in data.Descendants(dgm + "pt")) {
                string? modelId = (string?)point.Attribute("modelId");
                if (string.IsNullOrWhiteSpace(modelId)) {
                    continue;
                }

                string? type = (string?)point.Attribute("type");
                if (!string.IsNullOrEmpty(type) && !string.Equals(type, "node", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                XElement? textRoot = point.Element(dgm + "t") ?? point.Element(dgm + "txBody");
                if (textRoot == null) {
                    continue;
                }

                string text = string.Concat(textRoot.Descendants(a + "t").Select(element => (string?)element ?? string.Empty));
                if (text.Length > 0) {
                    result[modelId!] = text;
                }
            }

            return result;
        }

        private static Dictionary<string, string> GetPresentationSourceByDestinationId(XDocument data) {
            XNamespace dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement connection in data.Descendants(dgm + "cxn")) {
                string? type = (string?)connection.Attribute("type");
                if (!string.Equals(type, "presOf", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                string? sourceId = (string?)connection.Attribute("srcId");
                string? destinationId = (string?)connection.Attribute("destId");
                if (!string.IsNullOrWhiteSpace(sourceId) && !string.IsNullOrWhiteSpace(destinationId)) {
                    result[destinationId!] = sourceId!;
                }
            }

            return result;
        }
    }

    internal sealed class WordSmartArtPersistedShape {
        internal WordSmartArtPersistedShape(
            string modelId,
            string presetName,
            string text,
            double x,
            double y,
            double width,
            double height,
            double rotationDegrees) {
            ModelId = modelId;
            PresetName = presetName;
            Text = text;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            RotationDegrees = rotationDegrees;
        }

        internal string ModelId { get; }

        internal string PresetName { get; }

        internal string Text { get; }

        internal double X { get; }

        internal double Y { get; }

        internal double Width { get; }

        internal double Height { get; }

        internal double RotationDegrees { get; }
    }
}
