using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents an auto shape without a text body.
    /// </summary>
    public class PowerPointAutoShape : PowerPointShape {
        internal PowerPointAutoShape(Shape shape) : base(shape) {
        }

        private Shape Shape => (Shape)Element;

        /// <summary>
        ///     Gets the preset geometry type of the shape.
        /// </summary>
        public A.ShapeTypeValues? ShapeType => Shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value;

        /// <summary>
        ///     Sets the fill color and returns the shape for chaining.
        /// </summary>
        public PowerPointAutoShape Fill(string color) {
            FillColor = color;
            return this;
        }

        /// <summary>
        ///     Sets the outline color (and optional width in points) and returns the shape for chaining.
        /// </summary>
        public PowerPointAutoShape Stroke(string color, double? widthPoints = null) {
            OutlineColor = color;
            if (widthPoints != null) {
                OutlineWidthPoints = widthPoints;
            }
            return this;
        }
    }
}
