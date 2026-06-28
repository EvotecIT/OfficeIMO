using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a native PowerPoint connection shape.
    /// </summary>
    public class PowerPointConnectionShape : PowerPointShape {
        internal PowerPointConnectionShape(ConnectionShape shape) : base(shape) {
        }

        private ConnectionShape Shape => (ConnectionShape)Element;

        /// <summary>
        ///     Gets the preset geometry type of the connection shape.
        /// </summary>
        public A.ShapeTypeValues? ShapeType => Shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value;

        /// <summary>
        ///     Sets the outline color (and optional width in points) and returns the shape for chaining.
        /// </summary>
        public PowerPointConnectionShape Stroke(string color, double? widthPoints = null) {
            OutlineColor = color;
            if (widthPoints != null) {
                OutlineWidthPoints = widthPoints;
            }

            return this;
        }
    }
}
