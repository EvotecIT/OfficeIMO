using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Returns layout and master shapes that are inherited by this slide for export adapters.
        /// </summary>
        internal IReadOnlyList<PowerPointShape> GetInheritedShapesForExport() {
            var shapes = new List<PowerPointShape>();
            if (!ShowsMasterShapes(SlideRoot.CommonSlideData)) {
                return shapes;
            }

            SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
            SlideMasterPart? masterPart = layoutPart?.SlideMasterPart;
            List<PowerPointShape> layoutShapes = CreateInheritedShapes(layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree, layoutPart);
            IReadOnlyList<PowerPointShape> slideShapes = Shapes;

            if (ShowsMasterShapes(layoutPart?.SlideLayout?.CommonSlideData)) {
                foreach (PowerPointShape masterShape in CreateInheritedShapes(masterPart?.SlideMaster?.CommonSlideData?.ShapeTree, masterPart)) {
                    if (!IsPlaceholderOverridden(masterShape, layoutShapes) &&
                        !IsPlaceholderOverridden(masterShape, slideShapes)) {
                        shapes.Add(masterShape);
                    }
                }
            }

            foreach (PowerPointShape layoutShape in layoutShapes) {
                if (!IsPlaceholderOverridden(layoutShape, slideShapes)) {
                    shapes.Add(layoutShape);
                }
            }

            return shapes;
        }

        private static bool ShowsMasterShapes(CommonSlideData? commonSlideData) {
            if (commonSlideData == null) {
                return true;
            }

            string? value = null;
            foreach (OpenXmlAttribute attribute in commonSlideData.GetAttributes()) {
                if (attribute.LocalName == "showMasterSp") {
                    value = attribute.Value;
                    break;
                }
            }

            return string.IsNullOrWhiteSpace(value) ||
                value == "1" ||
                value?.Equals("true", System.StringComparison.OrdinalIgnoreCase) == true;
        }

        private List<PowerPointShape> CreateInheritedShapes(ShapeTree? tree, OpenXmlPartContainer? ownerPart) {
            var shapes = new List<PowerPointShape>();
            if (tree == null || ownerPart == null) {
                return shapes;
            }

            foreach (OpenXmlElement element in tree.ChildElements) {
                PowerPointShape? shape = CreateShapeFromElement(element, ownerPart);
                if (shape != null) {
                    shapes.Add(shape.AttachTo(this));
                }
            }

            return shapes;
        }

        private static bool IsPlaceholderOverridden(PowerPointShape inheritedShape, IReadOnlyList<PowerPointShape> overridingShapes) {
            if (!TryGetPlaceholderSignature(inheritedShape, out PlaceholderValues? inheritedType, out uint? inheritedIndex)) {
                return false;
            }

            foreach (PowerPointShape overridingShape in overridingShapes) {
                if (TryGetPlaceholderSignature(overridingShape, out PlaceholderValues? overridingType, out uint? overridingIndex) &&
                    PlaceholderSignaturesMatch(inheritedType, inheritedIndex, overridingType, overridingIndex)) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetPlaceholderSignature(PowerPointShape shape, out PlaceholderValues? type, out uint? index) {
            type = shape.ShapePlaceholderType;
            index = shape.ShapePlaceholderIndex;
            return type.HasValue || index.HasValue;
        }

        private static bool PlaceholderSignaturesMatch(PlaceholderValues? inheritedType, uint? inheritedIndex, PlaceholderValues? overridingType, uint? overridingIndex) {
            if (inheritedIndex.HasValue && overridingIndex.HasValue) {
                return inheritedIndex.Value == overridingIndex.Value &&
                    (!inheritedType.HasValue || !overridingType.HasValue || inheritedType.Value == overridingType.Value);
            }

            return inheritedType.HasValue &&
                overridingType.HasValue &&
                inheritedType.Value == overridingType.Value;
        }
    }
}
