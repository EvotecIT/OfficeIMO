using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static IReadOnlyList<PowerPointShape> ReadSlideShapesForWrite(
            PowerPointSlide slide, out string? unsupportedReason) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            unsupportedReason = null;
            PowerPointShape[] slideShapes = slide.Shapes.ToArray();
            if (!ShowsLayoutShapes(slide.SlidePart.Slide)) {
                return slideShapes;
            }

            SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
            P.ShapeTree? tree = layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
            if (layoutPart == null || tree == null) return slideShapes;

            var result = new List<PowerPointShape>(
                checked(tree.ChildElements.Count + slideShapes.Length));
            foreach (OpenXmlElement element in tree.ChildElements) {
                if (element is P.NonVisualGroupShapeProperties
                    or P.GroupShapeProperties) continue;
                PowerPointShape? shape = WrapInheritedShape(element, layoutPart);
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

        private static PowerPointShape? WrapInheritedShape(OpenXmlElement element,
            OpenXmlPartContainer ownerPart) => element switch {
                P.Shape shape when shape.TextBody != null =>
                    new PowerPointTextBox(shape, ownerPart),
                P.Shape shape => new PowerPointAutoShape(shape),
                _ => null
            };

        private static bool ShowsLayoutShapes(P.Slide? slide) =>
            slide?.ShowMasterShapes?.Value != false;

        internal static bool FollowsMasterObjects(PowerPointSlide slide,
            bool layoutIsIndependentMaster = false) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            return slide.SlidePart.Slide?.ShowMasterShapes?.Value != false
                && (layoutIsIndependentMaster
                    || slide.SlidePart.SlideLayoutPart?.SlideLayout?
                        .ShowMasterShapes?.Value != false);
        }

        private static bool IsPlaceholderOverridden(PowerPointShape inheritedShape,
            IReadOnlyList<PowerPointShape> slideShapes) {
            if (!TryGetPlaceholderSignature(inheritedShape,
                    out P.PlaceholderValues? inheritedType,
                    out uint? inheritedIndex)) return false;
            return slideShapes.Any(shape => TryGetPlaceholderSignature(shape,
                    out P.PlaceholderValues? slideType, out uint? slideIndex)
                && PlaceholderSignaturesMatch(inheritedType, inheritedIndex,
                    slideType, slideIndex));
        }

        private static bool TryGetPlaceholderSignature(PowerPointShape shape,
            out P.PlaceholderValues? type, out uint? index) {
            type = shape.ShapePlaceholderType;
            index = shape.ShapePlaceholderIndex;
            return type.HasValue || index.HasValue;
        }

        private static bool PlaceholderSignaturesMatch(P.PlaceholderValues? leftType,
            uint? leftIndex, P.PlaceholderValues? rightType, uint? rightIndex) {
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
