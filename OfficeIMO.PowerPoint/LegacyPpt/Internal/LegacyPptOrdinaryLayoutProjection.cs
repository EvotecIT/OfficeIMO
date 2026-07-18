using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Captures the synthetic editable layout scope projected from a classic
    /// slide layout signature so supported changes can be materialized without
    /// treating it as an independent binary persist object.
    /// </summary>
    internal sealed class LegacyPptOrdinaryLayoutProjection {
        private readonly IReadOnlyDictionary<uint, string> _shapesById;

        private LegacyPptOrdinaryLayoutProjection(SlideLayoutPart layoutPart,
            IReadOnlyDictionary<uint, string> shapesById) {
            PartUri = layoutPart.Uri.ToString();
            ShapeTreeFingerprint = CreateShapeTreeFingerprint(layoutPart);
            TypeFingerprint = CreateTypeFingerprint(layoutPart);
            _shapesById = new ReadOnlyDictionary<uint, string>(
                shapesById.ToDictionary(pair => pair.Key,
                    pair => pair.Value));
        }

        internal string PartUri { get; }

        internal string ShapeTreeFingerprint { get; }

        internal string TypeFingerprint { get; }

        internal bool ShapeTreeMatches(SlideLayoutPart layoutPart) =>
            string.Equals(ShapeTreeFingerprint,
                CreateShapeTreeFingerprint(layoutPart),
                StringComparison.Ordinal);

        internal bool TypeMatches(SlideLayoutPart layoutPart) =>
            string.Equals(TypeFingerprint,
                CreateTypeFingerprint(layoutPart),
                StringComparison.Ordinal);

        internal bool TryGetAddedShapeIds(SlideLayoutPart layoutPart,
            out IReadOnlyList<uint> addedShapeIds) {
            addedShapeIds = Array.Empty<uint>();
            if (!TryCreateShapeSnapshot(layoutPart,
                    out IReadOnlyDictionary<uint, string> current)) {
                return false;
            }
            foreach (KeyValuePair<uint, string> expected in _shapesById) {
                if (!current.TryGetValue(expected.Key,
                        out string? currentShape)
                    || !string.Equals(expected.Value, currentShape,
                        StringComparison.Ordinal)) {
                    return false;
                }
            }
            addedShapeIds = current.Keys.Where(id =>
                    !_shapesById.ContainsKey(id))
                .ToArray();
            return true;
        }

        internal static IReadOnlyList<LegacyPptOrdinaryLayoutProjection>
            Create(PowerPointPresentation presentation,
                IReadOnlyList<LegacyPptMasterProjection> titleMasters) {
            var titlePartUris = new HashSet<string>(titleMasters.Select(
                master => master.MasterPartUri), StringComparer.Ordinal);
            var result = new List<LegacyPptOrdinaryLayoutProjection>();
            foreach (SlideLayoutPart layout in presentation.OpenXmlDocument
                         .PresentationPart?.SlideMasterParts.SelectMany(
                             master => master.SlideLayoutParts)
                     ?? Enumerable.Empty<SlideLayoutPart>()) {
                if (titlePartUris.Contains(layout.Uri.ToString())
                    || !TryCreateShapeSnapshot(layout,
                        out IReadOnlyDictionary<uint, string> shapes)) {
                    continue;
                }
                result.Add(new LegacyPptOrdinaryLayoutProjection(layout,
                    shapes));
            }
            return result;
        }

        private static string CreateShapeTreeFingerprint(
            SlideLayoutPart layoutPart) => layoutPart.SlideLayout?
                .CommonSlideData?.ShapeTree?.OuterXml ?? string.Empty;

        private static string CreateTypeFingerprint(
            SlideLayoutPart layoutPart) => layoutPart.SlideLayout?.Type?
                .InnerText ?? string.Empty;

        private static bool TryCreateShapeSnapshot(
            SlideLayoutPart layoutPart,
            out IReadOnlyDictionary<uint, string> snapshot) {
            var result = new Dictionary<uint, string>();
            snapshot = result;
            foreach (OpenXmlElement element in layoutPart.SlideLayout?
                         .CommonSlideData?.ShapeTree?.ChildElements
                     ?? Enumerable.Empty<OpenXmlElement>()) {
                if (element is P.NonVisualGroupShapeProperties
                    or P.GroupShapeProperties) continue;
                uint? id = GetShapeId(element);
                if (!id.HasValue || result.ContainsKey(id.Value)) {
                    return false;
                }
                OpenXmlElement normalized = element.CloneNode(true);
                if (normalized is P.Shape shape
                    && IsHeaderFooterPlaceholder(shape)) {
                    shape.TextBody?.RemoveAllChildren<A.Paragraph>();
                }
                result.Add(id.Value, normalized.OuterXml);
            }
            return true;
        }

        private static uint? GetShapeId(OpenXmlElement element) =>
            element switch {
                P.Shape shape => shape.NonVisualShapeProperties?
                    .NonVisualDrawingProperties?.Id?.Value,
                P.Picture picture => picture.NonVisualPictureProperties?
                    .NonVisualDrawingProperties?.Id?.Value,
                P.ConnectionShape connector => connector
                    .NonVisualConnectionShapeProperties?
                    .NonVisualDrawingProperties?.Id?.Value,
                P.GroupShape group => group.NonVisualGroupShapeProperties?
                    .NonVisualDrawingProperties?.Id?.Value,
                P.GraphicFrame frame => frame
                    .NonVisualGraphicFrameProperties?
                    .NonVisualDrawingProperties?.Id?.Value,
                _ => null
            };

        private static bool IsHeaderFooterPlaceholder(P.Shape shape) {
            P.PlaceholderValues? type = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape?
                .Type?.Value;
            return type == P.PlaceholderValues.DateAndTime
                || type == P.PlaceholderValues.Footer
                || type == P.PlaceholderValues.SlideNumber;
        }
    }
}
