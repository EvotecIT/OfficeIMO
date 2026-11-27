using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static IReadOnlyList<PowerPointShape> ReadSlideShapesForWrite(
            PowerPointSlide slide, out string? unsupportedReason) {
            ArgumentNullException.ThrowIfNull(slide);
            unsupportedReason = null;
            PowerPointShape[] slideShapes = slide.Shapes.ToArray();
            PowerPointDocument presentation = OwnerDocumentForWrite(slide);
            EnsureShapeWrappersCanBind(presentation);
            if (!ShowsLayoutShapes(slide.PartRef.Root)) {
                return slideShapes;
            }

            PowerPointPartRef layoutPart = slide.PartRef.RelatedPart(
                PowerPointPartRef.SlideLayoutRelationshipType);
            OfficeXmlNode tree = layoutPart.Root.Path(Pml, "cSld", "spTree");
            if (layoutPart.IsNone || tree.IsNone) return slideShapes;

            var result = new List<PowerPointShape>(
                checked(tree.Children().Count + slideShapes.Length));
            foreach (OfficeXmlNode element in tree.Children()) {
                if (IsGroupProperties(element)) continue;
                PowerPointShape? shape = WrapShapeForWrite(presentation,
                    layoutPart, element);
                if (shape == null) {
                    unsupportedReason ??=
                        $"The slide layout contains '{element.LocalName}' content that is not yet materialized by the native binary writer.";
                    continue;
                }
                if (!IsPlaceholderOverridden(shape, slideShapes)) result.Add(shape);
            }
            result.AddRange(slideShapes);
            return result;
        }

        // The group's own non-visual and visual properties, which are not shapes of the tree.
        private static bool IsGroupProperties(OfficeXmlNode element) =>
            element.Is(Pml, "nvGrpSpPr") || element.Is(Pml, "grpSpPr");

        internal static IReadOnlyList<PowerPointShape>
            FlattenShapeTreeForWrite(IEnumerable<PowerPointShape> shapes,
                out string? unsupportedReason) =>
            FlattenShapeTreeForWrite(shapes, depth: 0,
                out unsupportedReason);

        private static IReadOnlyList<PowerPointShape>
            FlattenShapeTreeForWrite(IEnumerable<PowerPointShape> shapes,
                int depth, out string? unsupportedReason) {
            ArgumentNullException.ThrowIfNull(shapes);
            var result = new List<PowerPointShape>();
            unsupportedReason = null;
            foreach (PowerPointShape shape in shapes) {
                result.Add(shape);
                if (shape is not PowerPointGroupShape group) continue;
                if (depth >= MaximumGroupNestingDepth) {
                    unsupportedReason = $"Binary PowerPoint conversion supports at most {MaximumGroupNestingDepth} nested group levels.";
                    return result;
                }
                IReadOnlyList<PowerPointShape> children =
                    ReadGroupChildrenForWrite(group,
                        out string? childReason);
                if (childReason != null) {
                    unsupportedReason = childReason;
                    return result;
                }
                IReadOnlyList<PowerPointShape> descendants =
                    FlattenShapeTreeForWrite(children, checked(depth + 1),
                        out childReason);
                if (childReason != null) {
                    unsupportedReason = childReason;
                    return result;
                }
                result.AddRange(descendants);
            }
            return result;
        }

        private static bool ShowsLayoutShapes(OfficeXmlNode slide) =>
            slide.Bool("showMasterSp") != false;

        internal static bool IsLayoutShape(PowerPointShape shape) =>
            shape != null && shape.XmlElement.HasAncestor(Pml, "sldLayout");

        internal static bool FollowsMasterObjects(LegacyPptPartView slidePart,
            bool layoutIsIndependentMaster = false) =>
            slidePart.Root.Bool("showMasterSp") != false
            && (layoutIsIndependentMaster
                || slidePart.SlideLayout.Root.Bool("showMasterSp") != false);

        private static bool IsPlaceholderOverridden(PowerPointShape inheritedShape,
            IReadOnlyList<PowerPointShape> slideShapes) {
            if (!TryGetPlaceholderSignature(inheritedShape,
                    out PowerPointPlaceholderType? inheritedType,
                    out uint? inheritedIndex)) return false;
            return slideShapes.Any(shape => TryGetPlaceholderSignature(shape,
                    out PowerPointPlaceholderType? slideType, out uint? slideIndex)
                && PlaceholderSignaturesMatch(inheritedType, inheritedIndex,
                    slideType, slideIndex));
        }

        private static bool TryGetPlaceholderSignature(PowerPointShape shape,
            out PowerPointPlaceholderType? type, out uint? index) {
            type = shape.ShapePlaceholderType;
            index = shape.ShapePlaceholderIndex;
            return type.HasValue || index.HasValue;
        }

        private static bool PlaceholderSignaturesMatch(PowerPointPlaceholderType? leftType,
            uint? leftIndex, PowerPointPlaceholderType? rightType, uint? rightIndex) {
            if (leftIndex.HasValue && rightIndex.HasValue) {
                return leftIndex.Value == rightIndex.Value
                    && (!leftType.HasValue || !rightType.HasValue
                        || leftType.Value == rightType.Value);
            }
            return leftType.HasValue && rightType.HasValue
                && leftType.Value == rightType.Value;
        }
    }
}
