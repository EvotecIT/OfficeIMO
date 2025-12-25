using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds an auto shape with the specified geometry.
        /// </summary>
        public PowerPointAutoShape AddShape(A.ShapeTypeValues shapeType, long left = 0L, long top = 0L, long width = 914400L,
            long height = 914400L, string? name = null) {
            string shapeName = name ?? GenerateUniqueName(shapeType.ToString());

            ShapeProperties shapeProperties = new(
                new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = shapeType }
            );

            if (shapeType == A.ShapeTypeValues.Line) {
                shapeProperties.Append(new A.NoFill());
            }

            Shape shape = new(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = shapeName },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                shapeProperties
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(shape);
            PowerPointAutoShape autoShape = new(shape);
            _shapes.Add(autoShape);
            return autoShape;
        }

        /// <summary>
        ///     Adds an auto shape with the specified geometry using centimeter measurements.
        /// </summary>
        public PowerPointAutoShape AddShapeCm(A.ShapeTypeValues shapeType, double leftCm, double topCm, double widthCm,
            double heightCm, string? name = null) {
            return AddShape(shapeType,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm),
                name);
        }

        /// <summary>
        ///     Adds an auto shape with the specified geometry using inch measurements.
        /// </summary>
        public PowerPointAutoShape AddShapeInches(A.ShapeTypeValues shapeType, double leftInches, double topInches,
            double widthInches, double heightInches, string? name = null) {
            return AddShape(shapeType,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches),
                name);
        }

        /// <summary>
        ///     Adds an auto shape with the specified geometry using point measurements.
        /// </summary>
        public PowerPointAutoShape AddShapePoints(A.ShapeTypeValues shapeType, double leftPoints, double topPoints,
            double widthPoints, double heightPoints, string? name = null) {
            return AddShape(shapeType,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints),
                name);
        }

        /// <summary>
        ///     Adds a rectangle shape.
        /// </summary>
        public PowerPointAutoShape AddRectangle(long left, long top, long width, long height, string? name = null) {
            return AddShape(A.ShapeTypeValues.Rectangle, left, top, width, height, name);
        }

        /// <summary>
        ///     Adds a rectangle shape using centimeter measurements.
        /// </summary>
        public PowerPointAutoShape AddRectangleCm(double leftCm, double topCm, double widthCm, double heightCm,
            string? name = null) {
            return AddShapeCm(A.ShapeTypeValues.Rectangle, leftCm, topCm, widthCm, heightCm, name);
        }

        /// <summary>
        ///     Adds a rectangle shape using inch measurements.
        /// </summary>
        public PowerPointAutoShape AddRectangleInches(double leftInches, double topInches, double widthInches,
            double heightInches, string? name = null) {
            return AddShapeInches(A.ShapeTypeValues.Rectangle, leftInches, topInches, widthInches, heightInches, name);
        }

        /// <summary>
        ///     Adds a rectangle shape using point measurements.
        /// </summary>
        public PowerPointAutoShape AddRectanglePoints(double leftPoints, double topPoints, double widthPoints,
            double heightPoints, string? name = null) {
            return AddShapePoints(A.ShapeTypeValues.Rectangle, leftPoints, topPoints, widthPoints, heightPoints, name);
        }

        /// <summary>
        ///     Adds an ellipse shape.
        /// </summary>
        public PowerPointAutoShape AddEllipse(long left, long top, long width, long height, string? name = null) {
            return AddShape(A.ShapeTypeValues.Ellipse, left, top, width, height, name);
        }

        /// <summary>
        ///     Adds an ellipse shape using centimeter measurements.
        /// </summary>
        public PowerPointAutoShape AddEllipseCm(double leftCm, double topCm, double widthCm, double heightCm,
            string? name = null) {
            return AddShapeCm(A.ShapeTypeValues.Ellipse, leftCm, topCm, widthCm, heightCm, name);
        }

        /// <summary>
        ///     Adds an ellipse shape using inch measurements.
        /// </summary>
        public PowerPointAutoShape AddEllipseInches(double leftInches, double topInches, double widthInches,
            double heightInches, string? name = null) {
            return AddShapeInches(A.ShapeTypeValues.Ellipse, leftInches, topInches, widthInches, heightInches, name);
        }

        /// <summary>
        ///     Adds an ellipse shape using point measurements.
        /// </summary>
        public PowerPointAutoShape AddEllipsePoints(double leftPoints, double topPoints, double widthPoints,
            double heightPoints, string? name = null) {
            return AddShapePoints(A.ShapeTypeValues.Ellipse, leftPoints, topPoints, widthPoints, heightPoints, name);
        }

        /// <summary>
        ///     Adds a line shape from the specified start and end points.      
        /// </summary>
        public PowerPointAutoShape AddLine(long startX, long startY, long endX, long endY, string? name = null) {
            long left = Math.Min(startX, endX);
            long top = Math.Min(startY, endY);
            long width = Math.Abs(endX - startX);
            long height = Math.Abs(endY - startY);

            if (width == 0) width = 1;
            if (height == 0) height = 1;

            return AddShape(A.ShapeTypeValues.Line, left, top, width, height, name);
        }

        /// <summary>
        ///     Adds a line shape using centimeter measurements.
        /// </summary>
        public PowerPointAutoShape AddLineCm(double startXCm, double startYCm, double endXCm, double endYCm,
            string? name = null) {
            return AddLine(
                PowerPointUnits.FromCentimeters(startXCm),
                PowerPointUnits.FromCentimeters(startYCm),
                PowerPointUnits.FromCentimeters(endXCm),
                PowerPointUnits.FromCentimeters(endYCm),
                name);
        }

        /// <summary>
        ///     Adds a line shape using inch measurements.
        /// </summary>
        public PowerPointAutoShape AddLineInches(double startXInches, double startYInches, double endXInches,
            double endYInches, string? name = null) {
            return AddLine(
                PowerPointUnits.FromInches(startXInches),
                PowerPointUnits.FromInches(startYInches),
                PowerPointUnits.FromInches(endXInches),
                PowerPointUnits.FromInches(endYInches),
                name);
        }

        /// <summary>
        ///     Adds a line shape using point measurements.
        /// </summary>
        public PowerPointAutoShape AddLinePoints(double startXPoints, double startYPoints, double endXPoints,
            double endYPoints, string? name = null) {
            return AddLine(
                PowerPointUnits.FromPoints(startXPoints),
                PowerPointUnits.FromPoints(startYPoints),
                PowerPointUnits.FromPoints(endXPoints),
                PowerPointUnits.FromPoints(endYPoints),
                name);
        }
    }
}
