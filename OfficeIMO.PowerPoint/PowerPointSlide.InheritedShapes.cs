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
            SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
            SlideMasterPart? masterPart = layoutPart?.SlideMasterPart;

            AddInheritedShapes(masterPart?.SlideMaster?.CommonSlideData?.ShapeTree, masterPart, shapes);
            AddInheritedShapes(layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree, layoutPart, shapes);
            return shapes;
        }

        private void AddInheritedShapes(ShapeTree? tree, OpenXmlPartContainer? ownerPart, List<PowerPointShape> shapes) {
            if (tree == null || ownerPart == null) {
                return;
            }

            foreach (OpenXmlElement element in tree.ChildElements) {
                PowerPointShape? shape = CreateShapeFromElement(element, ownerPart);
                if (shape != null) {
                    shapes.Add(shape);
                }
            }
        }
    }
}
