using System;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Reusable visual style for Visio shapes.
    /// </summary>
    public sealed class VisioShapeStyle {
        /// <summary>
        /// Initializes a new shape style.
        /// </summary>
        /// <param name="fillColor">Fill color.</param>
        /// <param name="lineColor">Line color.</param>
        /// <param name="lineWeight">Line weight in inches.</param>
        /// <param name="linePattern">Line pattern index.</param>
        /// <param name="fillPattern">Fill pattern index.</param>
        public VisioShapeStyle(Color fillColor, Color lineColor, double lineWeight = VisioShape.DefaultLineWeight, int linePattern = 1, int fillPattern = 1) {
            FillColor = fillColor;
            LineColor = lineColor;
            LineWeight = lineWeight;
            LinePattern = linePattern;
            FillPattern = fillPattern;
        }

        /// <summary>Fill color.</summary>
        public Color FillColor { get; set; }

        /// <summary>Line color.</summary>
        public Color LineColor { get; set; }

        /// <summary>Line weight in inches.</summary>
        public double LineWeight { get; set; }

        /// <summary>Line pattern index.</summary>
        public int LinePattern { get; set; }

        /// <summary>Fill pattern index.</summary>
        public int FillPattern { get; set; }

        /// <summary>Optional text style to apply with this shape style.</summary>
        public VisioTextStyle? TextStyle { get; set; }

        /// <summary>Creates a detached copy of this style.</summary>
        public VisioShapeStyle Clone() {
            return new VisioShapeStyle(FillColor, LineColor, LineWeight, LinePattern, FillPattern) {
                TextStyle = TextStyle?.Clone()
            };
        }

        /// <summary>Applies this style to a shape.</summary>
        /// <param name="shape">Shape to update.</param>
        public void ApplyTo(VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            shape.FillColor = FillColor;
            shape.LineColor = LineColor;
            shape.LineWeight = LineWeight;
            shape.LinePattern = LinePattern;
            shape.FillPattern = FillPattern;
            if (TextStyle != null) {
                shape.TextStyle = TextStyle.Clone();
            }
        }
    }
}
