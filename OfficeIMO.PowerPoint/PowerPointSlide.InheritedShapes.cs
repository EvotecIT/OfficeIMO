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

            if (ShowsMasterShapes(layoutPart?.SlideLayout?.CommonSlideData)) {
                AddInheritedShapes(masterPart?.SlideMaster?.CommonSlideData?.ShapeTree, masterPart, shapes);
            }

            AddInheritedShapes(layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree, layoutPart, shapes);
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
