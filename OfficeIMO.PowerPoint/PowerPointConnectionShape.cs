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
        public OfficePresetShapeType? ShapeType => Shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value.ToOfficeEnum();

        /// <summary>
        ///     Sets the outline color (and optional physical width) and returns the shape for chaining.
        /// </summary>
        public PowerPointConnectionShape Stroke(string color, OfficeLength? width = null) {
            if (width.HasValue) ToDrawingLineWidth(width.Value, nameof(width));
            OutlineColor = color;
            if (width.HasValue) {
                OutlineWidth = width;
            }

            return this;
        }
    }
}
