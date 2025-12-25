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
            PowerPointTextBox textBox = new(shape, _slidePart);
            _shapes.Add(textBox);
            return textBox;
        }

        /// <summary>
        ///     Adds a title textbox using a layout box.
        /// </summary>
        public PowerPointTextBox AddTitle(string text, PowerPointLayoutBox layout) {
            return AddTitle(text, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds a title textbox using centimeter measurements.
        /// </summary>
        public PowerPointTextBox AddTitleCm(string text, double leftCm, double topCm, double widthCm, double heightCm) {
            return AddTitle(text,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a title textbox using inch measurements.
        /// </summary>
        public PowerPointTextBox AddTitleInches(string text, double leftInches, double topInches, double widthInches,
            double heightInches) {
            return AddTitle(text,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a title textbox using point measurements.
        /// </summary>
        public PowerPointTextBox AddTitlePoints(string text, double leftPoints, double topPoints, double widthPoints,
            double heightPoints) {
            return AddTitle(text,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
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
            PowerPointTextBox textBox = new(shape, _slidePart);
            _shapes.Add(textBox);
            return textBox;
        }

        /// <summary>
        ///     Adds a textbox using a layout box.
        /// </summary>
        public PowerPointTextBox AddTextBox(string text, PowerPointLayoutBox layout) {
            return AddTextBox(text, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds an empty textbox using a layout box.
        /// </summary>
        public PowerPointTextBox AddTextBox(PowerPointLayoutBox layout) {
            return AddTextBox(string.Empty, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds a textbox using centimeter measurements.
        /// </summary>
        public PowerPointTextBox AddTextBoxCm(string text, double leftCm, double topCm, double widthCm,
            double heightCm) {
            return AddTextBox(text,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a textbox using inch measurements.
        /// </summary>
        public PowerPointTextBox AddTextBoxInches(string text, double leftInches, double topInches, double widthInches,
            double heightInches) {
            return AddTextBox(text,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a textbox using point measurements.
        /// </summary>
        public PowerPointTextBox AddTextBoxPoints(string text, double leftPoints, double topPoints, double widthPoints,
            double heightPoints) {
            return AddTextBox(text,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

    }
}
