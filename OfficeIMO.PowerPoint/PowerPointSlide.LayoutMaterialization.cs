using System;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>Transfers a layout-owned text frame without replacing its native wrapper or body.</summary>
        internal void AdoptLayoutTextBox(PowerPointTextBox box) {
            AdoptLayoutShape(box);
        }

        /// <summary>Moves a materialized shape together with its interaction relationships.</summary>
        internal void AdoptLayoutShape(PowerPointShape box) {
            if (ReferenceEquals(box.OwnerSlide, this)) return;
            var drawing = box.Element.Descendants<NonVisualDrawingProperties>().First();
            foreach (var child in box.Element.ChildElements) {
                // Paragraphs and table rows bind individually: detached later slices still own source IDs.
                if (child is TextBody || box is PowerPointTable && child is A.Graphic) continue;
                PowerPointLayoutNativeText.RehomeRelationships(child, box.RelationshipOwnerPart, _slidePart);
            }
            box.OwnerSlide?.ShapeList.Remove(box);
            if (box.Element.Parent != null) box.Element.Remove();
            drawing.Id = AllocateShapeId();
            drawing.Name = GenerateUniqueName(box.Name ?? "Content");
            SlideRoot.CommonSlideData!.ShapeTree!.Append(box.Element);
            box.AttachTo(this, _slidePart);
            ShapeList.Add(box);
        }
    }
}
