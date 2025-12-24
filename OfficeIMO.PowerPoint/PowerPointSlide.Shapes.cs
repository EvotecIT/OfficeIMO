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
        ///     Adds a rectangle shape.
        /// </summary>
        public PowerPointAutoShape AddRectangle(long left, long top, long width, long height, string? name = null) {
            return AddShape(A.ShapeTypeValues.Rectangle, left, top, width, height, name);
        }

        /// <summary>
        ///     Adds an ellipse shape.
        /// </summary>
        public PowerPointAutoShape AddEllipse(long left, long top, long width, long height, string? name = null) {
            return AddShape(A.ShapeTypeValues.Ellipse, left, top, width, height, name);
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
    }
}
