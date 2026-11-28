using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Returns layout and master shapes that are inherited by this slide for export adapters.
        /// </summary>
        internal IReadOnlyList<PowerPointShape> GetInheritedShapesForExport() {
            var shapes = new List<PowerPointShape>();
            PowerPointPartRef slide = PartRef;
            if (slide.Root.BooleanAttribute("showMasterSp") == false) return shapes;

            PowerPointPartRef layout = slide.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
            if (layout.Root.BooleanAttribute("showMasterSp") != false) {
                AddInheritedShapes(layout.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType), shapes);
            }

            AddInheritedShapes(layout, shapes);
            return shapes;
        }

        private void AddInheritedShapes(PowerPointPartRef part, List<PowerPointShape> shapes) {
            const int presentation = OfficeXmlNamespaces.PresentationML;
            OfficeXmlNode tree = part.Root.Child(presentation, "cSld").Child(presentation, "spTree");
            foreach (OfficeXmlNode element in tree.Children()) {
                PowerPointShape? shape = WrapSlideShape(element, part);
                // Placeholders supply geometry and text styles. Their prompts appear only on the editing canvas.
                if (shape != null && !shape.HasPlaceholderShape) shapes.Add(shape);
            }
        }
    }
}
