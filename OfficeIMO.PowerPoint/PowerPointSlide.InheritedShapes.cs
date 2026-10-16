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

            if (ShowsMasterShapes(layoutPart?.SlideLayout)) {
                // Master placeholders define geometry and styles for concrete slide/layout
                // placeholders. Their editing prompts are not inherited slide content.
                foreach (PowerPointShape masterShape in CreateInheritedShapes(masterPart?.SlideMaster?.CommonSlideData?.ShapeTree, masterPart)) {
                    if (!IsStructuralPlaceholder(masterShape)) {
                        shapes.Add(masterShape);
                    }
                }
            }

            foreach (PowerPointShape layoutShape in layoutShapes) {
                // Layout placeholders also supply geometry/styles and editing prompts.
                // A visible field or text placeholder must be instantiated on the slide.
                if (!IsStructuralPlaceholder(layoutShape)) {
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
                    shapes.Add(shape.AttachTo(this, ownerPart));
                }
            }

            return shapes;
        }

        private static bool IsStructuralPlaceholder(PowerPointShape shape) =>
            shape.HasPlaceholderShape;
    }
}
