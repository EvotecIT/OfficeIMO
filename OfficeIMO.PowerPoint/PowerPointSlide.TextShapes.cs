using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Adds a text-bearing shape with the specified preset geometry.
        /// </summary>
        public PowerPointTextBox AddTextShape(
            OfficePresetShapeType shapeType,
            string text,
            OfficeLength? left = null,
            OfficeLength? top = null,
            OfficeLength? width = null,
            OfficeLength? height = null,
            string? name = null) {
            long nativeLeft = (left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(914400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(914400L)).ToEmus();
            if (text == null) throw new ArgumentNullException(nameof(text));

            string shapeName = name ?? GenerateUniqueName(shapeType.ToString());
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = AllocateShapeId(), Name = shapeName },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = nativeLeft, Y = nativeTop }, new A.Extents { Cx = nativeWidth, Cy = nativeHeight }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = shapeType.ToOpenXml() }),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(text)))));

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            return TrackShape(new PowerPointTextBox(shape, _slidePart));
        }

    }
}
