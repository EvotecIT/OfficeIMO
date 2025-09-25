using SixLabors.ImageSharp;

namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Fluent helper for configuring a shape (text, stroke, fill).
    /// </summary>
    public class VisioFluentShape {
        private readonly VisioShape _s;

        /// <summary>Initializes a new shape wrapper.</summary>
        /// <param name="shape">Underlying shape model.</param>
        internal VisioFluentShape(VisioShape shape) { _s = shape; }

        /// <summary>Sets shape text.</summary>
        /// <param name="text">Text content.</param>
        public VisioFluentShape Text(string text) { _s.Text = text; return this; }

        /// <summary>Sets fill color and optional pattern.</summary>
        /// <param name="color">Fill color.</param>
        /// <param name="pattern">Fill pattern index (default 1=Solid).</param>
        public VisioFluentShape Fill(Color color, int pattern = 1) { _s.FillColor = color; _s.FillPattern = pattern; return this; }

        /// <summary>Sets stroke color, weight (inches), and optional pattern.</summary>
        /// <param name="color">Line color.</param>
        /// <param name="weight">Line weight in inches.</param>
        /// <param name="pattern">Line pattern index (default 1=Solid).</param>
        public VisioFluentShape Stroke(Color color, double weight = 0.0138889, int pattern = 1) { _s.LineColor = color; _s.LineWeight = weight; _s.LinePattern = pattern; return this; }
    }
}

