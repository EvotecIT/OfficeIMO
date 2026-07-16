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
            if (!ShowsLayoutShapes(slide.SlidePart.Slide?.CommonSlideData)) {
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
                PowerPointShape? shape = WrapLayoutShape(element, layoutPart);
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

        private static PowerPointShape? WrapLayoutShape(OpenXmlElement element,
            OpenXmlPartContainer ownerPart) => element switch {
                P.Shape shape when shape.TextBody != null =>
                    new PowerPointTextBox(shape, ownerPart),
                P.Shape shape => new PowerPointAutoShape(shape),
                _ => null
            };

        private static bool ShowsLayoutShapes(P.CommonSlideData? commonSlideData) {
            if (commonSlideData == null) return true;
            string? value = commonSlideData.GetAttributes()
                .FirstOrDefault(attribute => attribute.LocalName == "showMasterSp")
                .Value;
            return string.IsNullOrWhiteSpace(value)
                || value == "1"
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
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
