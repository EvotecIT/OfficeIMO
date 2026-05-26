using Color = OfficeIMO.Drawing.OfficeColor;

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

        /// <summary>Applies a reusable shape style.</summary>
        /// <param name="style">Shape style to apply.</param>
        public VisioFluentShape Style(VisioShapeStyle style) { _s.ApplyStyle(style); return this; }

        /// <summary>Sets the whole-shape text color.</summary>
        public VisioFluentShape TextColor(Color color) { EnsureTextStyle().Color = color; return this; }

        /// <summary>Sets the whole-shape text font family.</summary>
        public VisioFluentShape Font(string fontFamily) { EnsureTextStyle().FontFamily = fontFamily; return this; }

        /// <summary>Sets the whole-shape text size in points.</summary>
        public VisioFluentShape FontSize(double size) {
            if (double.IsNaN(size) || double.IsInfinity(size) || size <= 0D) {
                throw new System.ArgumentOutOfRangeException(nameof(size), "Font size must be a finite positive number.");
            }

            EnsureTextStyle().Size = size;
            return this;
        }

        /// <summary>Sets whether the whole-shape text is bold.</summary>
        public VisioFluentShape Bold(bool enabled = true) { EnsureTextStyle().Bold = enabled; return this; }

        /// <summary>Sets whole-shape text alignment.</summary>
        public VisioFluentShape TextAlignment(VisioTextHorizontalAlignment horizontal, VisioTextVerticalAlignment? vertical = null) {
            VisioTextStyle style = EnsureTextStyle();
            style.HorizontalAlignment = horizontal;
            if (vertical.HasValue) {
                style.VerticalAlignment = vertical.Value;
            }

            return this;
        }

        /// <summary>Adds the shape to a page layer.</summary>
        /// <param name="layerName">Layer name.</param>
        public VisioFluentShape Layer(string layerName) { _s.LayerNames.Add(layerName); return this; }

        /// <summary>Sets or replaces a Visio User cell.</summary>
        /// <param name="name">User cell row name.</param>
        /// <param name="value">User cell value.</param>
        /// <param name="unit">Optional Visio unit code.</param>
        /// <param name="formula">Optional ShapeSheet formula.</param>
        /// <param name="prompt">Optional prompt value.</param>
        public VisioFluentShape UserCell(string name, string? value, string? unit = null, string? formula = null, string? prompt = null) { _s.SetUserCell(name, value, unit, formula, prompt); return this; }

        /// <summary>Sets or replaces a Visio Shape Data row.</summary>
        /// <param name="name">Shape Data row name.</param>
        /// <param name="value">Shape Data value.</param>
        /// <param name="label">Optional label shown in Visio's Shape Data window.</param>
        /// <param name="type">Optional Shape Data type.</param>
        /// <param name="prompt">Optional help prompt.</param>
        /// <param name="format">Optional format picture or list values.</param>
        public VisioFluentShape ShapeData(string name, string? value, string? label = null, VisioShapeDataType? type = null, string? prompt = null, string? format = null) { _s.SetShapeData(name, value, label, type, prompt, format); return this; }

        /// <summary>Adds a hyperlink to the shape.</summary>
        /// <param name="address">External hyperlink address.</param>
        /// <param name="description">Optional display description.</param>
        /// <param name="subAddress">Optional internal sub-address.</param>
        public VisioFluentShape Hyperlink(string address, string? description = null, string? subAddress = null) { _s.AddHyperlink(address, description, subAddress); return this; }

        /// <summary>Configures ShapeSheet protection cells.</summary>
        /// <param name="configure">Protection configuration delegate.</param>
        public VisioFluentShape Protect(System.Action<VisioShapeProtection> configure) { _s.Protect(configure); return this; }

        /// <summary>Sets the shape-level Visio placement style.</summary>
        public VisioFluentShape PlacementStyle(VisioPlacementStyle style) { _s.PlacementStyle = style; return this; }

        /// <summary>Sets the shape-level Visio placement flip behavior.</summary>
        public VisioFluentShape PlacementFlip(VisioPlacementFlip flip) { _s.PlacementFlip = flip; return this; }

        /// <summary>Sets the shape-level Visio plow behavior.</summary>
        public VisioFluentShape Plow(VisioShapePlowCode code) { _s.PlowCode = code; return this; }

        /// <summary>Allows or disallows placing other shapes on top of this shape during layout.</summary>
        public VisioFluentShape PlacementOnTop(bool allowed = true) { _s.AllowPlacementOnTop = allowed; return this; }

        /// <summary>Allows or disallows connector routing through this shape.</summary>
        public VisioFluentShape ConnectorPermeability(bool horizontal = true, bool vertical = true) { _s.AllowHorizontalConnectorRoutingThrough = horizontal; _s.AllowVerticalConnectorRoutingThrough = vertical; return this; }

        /// <summary>Allows or disallows this shape splitting other shapes.</summary>
        public VisioFluentShape ShapeSplitting(bool canSplit = true, bool canBeSplit = true) { _s.CanSplitShapes = canSplit; _s.CanBeSplit = canBeSplit; return this; }

        /// <summary>Clears explicit Shape Layout override cells.</summary>
        public VisioFluentShape ClearLayoutPolicy() { _s.ClearLayoutPolicy(); return this; }

        private VisioTextStyle EnsureTextStyle() {
            if (_s.TextStyle == null) {
                _s.TextStyle = new VisioTextStyle();
            }

            return _s.TextStyle;
        }
    }
}
