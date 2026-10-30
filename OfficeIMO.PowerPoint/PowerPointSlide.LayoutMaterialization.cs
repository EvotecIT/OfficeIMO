using System;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>Transfers a layout-owned text frame without replacing its native wrapper or body.</summary>
        internal void AdoptLayoutTextBox(PowerPointTextBox box) {
            if (ReferenceEquals(box.OwnerSlide, this)) return;
            box.OwnerSlide?.ShapeList.Remove(box);
            box.Element.Remove();
            var shape = (Shape)box.Element;
            shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = AllocateShapeId();
            shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name = GenerateUniqueName(box.Name ?? "Text");
            SlideRoot.CommonSlideData!.ShapeTree!.Append(shape);
            box.AttachTo(this, _slidePart);
            ShapeList.Add(box);
        }
    }
}
