using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointEffectiveRunStyleResolver {
        private static OfficeXmlNode ShapePlaceholder(OfficeXmlNode shape) =>
            shape.Is(Presentation, "sp")
                ? shape.Child(Presentation, "nvSpPr").Child(Presentation, "nvPr").Child(Presentation, "ph")
                : default;

        /// <summary>Enumerates local, layout and master placeholder shapes in native inheritance order.</summary>
        private static IEnumerable<OfficeXmlNode> EnumeratePlaceholderShapes(OfficeXmlNode shape, PowerPointPartRef layout) {
            if (!shape.Is(Presentation, "sp")) yield break;
            yield return shape;
            OfficeXmlNode placeholder = ShapePlaceholder(shape);
            if (placeholder.IsNone || layout.IsNone) yield break;
            OfficeXmlNode layoutShape = FindPlaceholderShape(layout, placeholder, matchIndex: true);
            if (layoutShape.Exists && layoutShape != shape) yield return layoutShape;

            OfficeXmlNode masterPlaceholder = ShapePlaceholder(layoutShape);
            OfficeXmlNode masterShape = FindPlaceholderShape(layout.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType),
                masterPlaceholder.Exists ? masterPlaceholder : placeholder, matchIndex: false);
            if (masterShape.Exists && masterShape != shape) yield return masterShape;
        }

        private static void AddPlaceholderDefaults(List<OfficeXmlNode> sources, OfficeXmlNode shape, int level) {
            OfficeXmlNode body = shape.Child(Presentation, "txBody");
            if (PlaceholderParagraphProperties(body, level).Child(Drawing, "defRPr") is { Exists: true } defaults)
                sources.Add(defaults);
            sources.AddRange(FindDefaultRunProperties(body.Child(Drawing, "lstStyle"), level));
        }

        private static OfficeXmlNode PlaceholderParagraphProperties(OfficeXmlNode body, int level) {
            foreach (OfficeXmlNode paragraph in body.Children(Drawing, "p")) {
                OfficeXmlNode properties = paragraph.Child(Drawing, "pPr");
                if (ParagraphLevel(properties) == level) return properties;
            }

            return default;
        }

        private static OfficeXmlNode FindPlaceholderShape(PowerPointPartRef owner, OfficeXmlNode source, bool matchIndex) {
            OfficeXmlNode selected = default;
            OfficeXmlNode tree = owner.Root.Child(Presentation, "cSld").Child(Presentation, "spTree");
            uint index = source.UInt32Attribute("idx") ?? 0U;
            string type = NormalizePlaceholderType(source);
            foreach (OfficeXmlNode shape in tree.Children(Presentation, "sp")) {
                OfficeXmlNode candidate = ShapePlaceholder(shape);
                if (candidate.IsNone || (matchIndex
                    ? (candidate.UInt32Attribute("idx") ?? 0U) != index
                    : NormalizePlaceholderType(candidate) != type)) continue;
                // Ambiguous source declarations cannot establish an effective inherited value.
                if (selected.Exists) return default;
                selected = shape;
            }

            return selected;
        }

        private static string NormalizePlaceholderType(OfficeXmlNode value) => value.Attribute("type") switch {
            "ctrTitle" => "title",
            { } type => type,
            _ => "obj"
        };
    }
}
