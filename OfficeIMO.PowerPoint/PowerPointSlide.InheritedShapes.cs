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
            if (!ShowsMasterShapes(SlideRoot)) {
                return shapes;
            }

            SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
            SlideMasterPart? masterPart = layoutPart?.SlideMasterPart;
            List<PowerPointShape> layoutShapes = CreateInheritedShapes(layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree, layoutPart);
            IReadOnlyList<PowerPointShape> slideShapes = Shapes;
            var layoutOverrides = new PlaceholderOverrideIndex(layoutShapes);
            var slideOverrides = new PlaceholderOverrideIndex(slideShapes);

            if (ShowsMasterShapes(layoutPart?.SlideLayout)) {
                // Master placeholders define geometry and styles for concrete slide/layout
                // placeholders. Their editing prompts are not inherited slide content.
                foreach (PowerPointShape masterShape in CreateInheritedShapes(masterPart?.SlideMaster?.CommonSlideData?.ShapeTree, masterPart)) {
                    if (!IsStructuralPlaceholder(masterShape) &&
                        !layoutOverrides.ContainsMatch(masterShape) &&
                        !slideOverrides.ContainsMatch(masterShape)) {
                        shapes.Add(masterShape);
                    }
                }
            }

            foreach (PowerPointShape layoutShape in layoutShapes) {
                if (!slideOverrides.ContainsMatch(layoutShape)) {
                    shapes.Add(layoutShape);
                }
            }

            return shapes;
        }

        private static bool ShowsMasterShapes(Slide? slide) =>
            slide?.ShowMasterShapes?.Value != false;

        private static bool ShowsMasterShapes(SlideLayout? layout) =>
            layout?.ShowMasterShapes?.Value != false;

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

        private static bool IsStructuralPlaceholder(PowerPointShape shape) =>
            TryGetPlaceholderSignature(shape, out _, out _);

        private static bool TryGetPlaceholderSignature(PowerPointShape shape, out PlaceholderValues? type, out uint? index) {
            type = shape.ShapePlaceholderType;
            index = shape.ShapePlaceholderIndex;
            return type.HasValue || index.HasValue;
        }

        private sealed class PlaceholderOverrideIndex {
            private readonly HashSet<uint> _indices = new();
            private readonly HashSet<uint> _untypedIndices = new();
            private readonly Dictionary<uint, HashSet<PlaceholderValues>> _typesByIndex = new();
            private readonly HashSet<PlaceholderValues> _allTypes = new();
            private readonly HashSet<PlaceholderValues> _typesWithoutIndex = new();

            internal PlaceholderOverrideIndex(IReadOnlyList<PowerPointShape> shapes) {
                for (int index = 0; index < shapes.Count; index++) {
                    if (!TryGetPlaceholderSignature(shapes[index], out PlaceholderValues? type, out uint? placeholderIndex)) {
                        continue;
                    }

                    if (type.HasValue) {
                        _allTypes.Add(type.Value);
                    }

                    if (!placeholderIndex.HasValue) {
                        if (type.HasValue) {
                            _typesWithoutIndex.Add(type.Value);
                        }

                        continue;
                    }

                    _indices.Add(placeholderIndex.Value);
                    if (!type.HasValue) {
                        _untypedIndices.Add(placeholderIndex.Value);
                    } else {
                        if (!_typesByIndex.TryGetValue(placeholderIndex.Value, out HashSet<PlaceholderValues>? types)) {
                            types = new HashSet<PlaceholderValues>();
                            _typesByIndex.Add(placeholderIndex.Value, types);
                        }

                        types.Add(type.Value);
                    }
                }
            }

            internal bool ContainsMatch(PowerPointShape inheritedShape) {
                if (!TryGetPlaceholderSignature(inheritedShape, out PlaceholderValues? type, out uint? index)) {
                    return false;
                }

                if (!index.HasValue) {
                    return type.HasValue && _allTypes.Contains(type.Value);
                }

                if (!type.HasValue) {
                    return _indices.Contains(index.Value);
                }

                return _typesWithoutIndex.Contains(type.Value) ||
                    _untypedIndices.Contains(index.Value) ||
                    (_typesByIndex.TryGetValue(index.Value, out HashSet<PlaceholderValues>? types) && types.Contains(type.Value));
            }
        }
    }
}
