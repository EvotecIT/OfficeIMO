using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Adds a text-bearing shape with the specified preset geometry.
        /// </summary>
        public PowerPointTextBox AddTextShape(
            A.ShapeTypeValues shapeType,
            string text,
            long left = 0L,
            long top = 0L,
            long width = 914400L,
            long height = 914400L,
            string? name = null) {
            if (text == null) throw new ArgumentNullException(nameof(text));

            string shapeName = name ?? GenerateUniqueName(shapeType.ToString());
            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = AllocateShapeId(), Name = shapeName },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = shapeType }),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(text)))));

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            return TrackShape(new PowerPointTextBox(shape, _slidePart));
        }

        /// <summary>
        /// Adds a text-bearing shape using centimeter measurements.
        /// </summary>
        public PowerPointTextBox AddTextShapeCm(
            A.ShapeTypeValues shapeType,
            string text,
            double leftCm,
            double topCm,
            double widthCm,
            double heightCm,
            string? name = null) {
            return AddTextShape(
                shapeType,
                text,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm),
                name);
        }

        /// <summary>
        /// Adds a text-bearing shape using inch measurements.
        /// </summary>
        public PowerPointTextBox AddTextShapeInches(
            A.ShapeTypeValues shapeType,
            string text,
            double leftInches,
            double topInches,
            double widthInches,
            double heightInches,
            string? name = null) {
            return AddTextShape(
                shapeType,
                text,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches),
                name);
        }

        /// <summary>
        /// Adds a text-bearing shape using point measurements.
        /// </summary>
        public PowerPointTextBox AddTextShapePoints(
            A.ShapeTypeValues shapeType,
            string text,
            double leftPoints,
            double topPoints,
            double widthPoints,
            double heightPoints,
            string? name = null) {
            return AddTextShape(
                shapeType,
                text,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints),
                name);
        }
    }
}
