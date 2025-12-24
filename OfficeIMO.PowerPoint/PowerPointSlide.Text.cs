using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds a title textbox to the slide.
        /// </summary>
        public PowerPointTextBox AddTitle(string text, long left = 838200L, long top = 365125L,
            long width = 7772400L, long height = 1470025L) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            string name = GenerateUniqueName("Title");
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = PlaceholderValues.Title })
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(text)))
                )
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            PowerPointTextBox textBox = new(shape);
            _shapes.Add(textBox);
            return textBox;
        }

        /// <summary>
        ///     Adds a textbox with the specified text.
        /// </summary>
        public PowerPointTextBox AddTextBox(string text, long left = 838200L, long top = 2174875L, long width = 7772400L,
            long height = 3962400L) {
            if (text == null) {
                throw new ArgumentNullException(nameof(text));
            }

            string name = GenerateUniqueName("TextBox");
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(text)))
                )
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            PowerPointTextBox textBox = new(shape);
            _shapes.Add(textBox);
            return textBox;
        }

    }
}
