using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Dsp = DocumentFormat.OpenXml.Office.Drawing;

namespace OfficeIMO.Word {
    public partial class WordSmartArt {
        internal IReadOnlyList<WordSmartArtPersistedShape> GetPersistedLayoutShapes() {
            try {
                DiagramDataPart dataPart = GetDiagramDataPart();
                XDocument data = LoadDiagramXDocument(dataPart);
                string? persistRelationshipId = GetPersistedLayoutRelationshipId(data);
                if (string.IsNullOrWhiteSpace(persistRelationshipId)) {
                    return Array.Empty<WordSmartArtPersistedShape>();
                }

                var mainPart = _document._wordprocessingDocument.MainDocumentPart!;
                if (mainPart.GetPartById(persistRelationshipId!) is not DiagramPersistLayoutPart persistPart) {
                    return Array.Empty<WordSmartArtPersistedShape>();
                }

                Dsp.Drawing? drawing = persistPart.Drawing;
                if (drawing == null) {
                    return Array.Empty<WordSmartArtPersistedShape>();
                }

                Dictionary<string, string> nodeTextById = GetNodeTextByModelId(data);
                Dictionary<string, string> sourceByPresentationId = GetPresentationSourceByDestinationId(data);
                var shapes = new List<WordSmartArtPersistedShape>();
                foreach (Dsp.Shape shape in drawing.Descendants<Dsp.Shape>()) {
                    string? modelId = shape.ModelId?.Value;
                    if (string.IsNullOrWhiteSpace(modelId)) {
                        continue;
                    }

                    string shapeModelId = modelId!;

                    Dsp.ShapeProperties? shapeProperties = shape.GetFirstChild<Dsp.ShapeProperties>();
                    A.Transform2D? transform = shapeProperties?.GetFirstChild<A.Transform2D>();
                    A.Offset? offset = transform?.GetFirstChild<A.Offset>();
                    A.Extents? extents = transform?.GetFirstChild<A.Extents>();
                    if (offset?.X?.Value == null ||
                        offset.Y?.Value == null ||
                        extents?.Cx?.Value == null ||
                        extents.Cy?.Value == null ||
                        extents.Cx.Value <= 0L ||
                        extents.Cy.Value <= 0L) {
                        continue;
                    }

                    A.PresetGeometry? geometry = shapeProperties?.GetFirstChild<A.PresetGeometry>();
                    string preset = geometry?.Preset?.InnerText ?? string.Empty;
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

                    double rotation = transform?.Rotation?.Value is int rotationValue
                        ? rotationValue / 60000D
                        : 0D;

                    shapes.Add(new WordSmartArtPersistedShape(
                        shapeModelId,
                        preset,
                        text ?? string.Empty,
                        Helpers.ConvertEmusToPoints(offset.X.Value),
                        Helpers.ConvertEmusToPoints(offset.Y.Value),
                        Helpers.ConvertEmusToPoints(extents.Cx.Value),
                        Helpers.ConvertEmusToPoints(extents.Cy.Value),
                        rotation));
                }

                return shapes;
            } catch {
                return Array.Empty<WordSmartArtPersistedShape>();
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
