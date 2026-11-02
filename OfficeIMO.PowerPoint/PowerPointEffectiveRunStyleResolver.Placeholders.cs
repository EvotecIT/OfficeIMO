using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointEffectiveRunStyleResolver {
        /// <summary>Enumerates local, layout and master placeholder shapes in native inheritance order.</summary>
        private static IEnumerable<P.Shape> EnumeratePlaceholderShapes(P.Shape shape, SlideLayoutPart? layout) {
            yield return shape;
            P.PlaceholderShape? placeholder = shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
            if (placeholder == null) yield break;
            P.Shape? layoutShape = FindPlaceholderShape(layout?.SlideLayout?.CommonSlideData?.ShapeTree,
                placeholder, matchIndex: true);
            if (layoutShape != null && !ReferenceEquals(layoutShape, shape)) yield return layoutShape;
            P.PlaceholderShape masterPlaceholder = layoutShape?.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() ?? placeholder;
            P.Shape? masterShape = FindPlaceholderShape(layout?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree,
                masterPlaceholder, matchIndex: false);
            if (masterShape != null && !ReferenceEquals(masterShape, shape)) yield return masterShape;
        }

        private static void AddPlaceholderDefaults(List<A.TextCharacterPropertiesType> sources, P.Shape? shape, int level) {
            A.DefaultRunProperties? paragraphDefaults = shape?.TextBody?.Elements<A.Paragraph>()
                .FirstOrDefault(paragraph => (paragraph.ParagraphProperties?.Level?.Value ?? 0) == level)?
                .ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>();
            if (paragraphDefaults != null) sources.Add(paragraphDefaults);
            sources.AddRange(FindDefaultRunProperties(shape?.TextBody?.ListStyle, level));
        }

        private static P.Shape? FindPlaceholderShape(P.ShapeTree? tree, P.PlaceholderShape source, bool matchIndex) {
            IEnumerable<P.Shape> matches = tree?.Elements<P.Shape>() ?? Enumerable.Empty<P.Shape>();
            matches = matches.Where(shape => {
                P.PlaceholderShape? candidate = shape.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
                return candidate != null && (matchIndex
                    ? (candidate.Index?.Value ?? 0U) == (source.Index?.Value ?? 0U)
                    : NormalizePlaceholderType(candidate.Type?.Value) == NormalizePlaceholderType(source.Type?.Value));
            });
            // Ambiguous source declarations cannot establish an effective inherited value.
            P.Shape[] selected = matches.Take(2).ToArray();
            return selected.Length == 1 ? selected[0] : null;
        }

        private static P.PlaceholderValues NormalizePlaceholderType(P.PlaceholderValues? value) =>
            value == P.PlaceholderValues.CenteredTitle ? P.PlaceholderValues.Title : value ?? P.PlaceholderValues.Object;
    }
}
