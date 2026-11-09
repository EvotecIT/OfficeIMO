using OfficeIMO.Drawing;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Visio {
    public partial class VisioShape {
        /// <summary>Sets shape text.</summary>
        /// <param name="text">Text content.</param>
        public VisioShape SetText(string text) { Text = text; return this; }

        /// <summary>Sets fill color and optional pattern.</summary>
        /// <param name="color">Fill color.</param>
        /// <param name="pattern">Fill pattern index (default 1=Solid).</param>
        public VisioShape SetFill(Color color, int pattern = 1) { FillColor = color; FillPattern = pattern; return this; }

        /// <summary>Sets stroke color, weight (inches), and optional pattern.</summary>
        /// <param name="color">Line color.</param>
        /// <param name="weight">Line weight in inches.</param>
        /// <param name="pattern">Line pattern index (default 1=Solid).</param>
        public VisioShape SetStroke(Color color, double weight = 0.0138889, int pattern = 1) { LineColor = color; LineWeight = weight; LinePattern = pattern; return this; }

        /// <summary>Sets the whole-shape text color.</summary>
        public VisioShape SetTextColor(Color color) { EnsureTextStyle().Color = color; return this; }

        /// <summary>Sets the whole-shape text font family.</summary>
        public VisioShape SetFont(string fontFamily) { EnsureTextStyle().FontFamily = fontFamily; return this; }

        /// <summary>Sets the whole-shape text size in points.</summary>
        public VisioShape SetFontSize(double size) {
            if (double.IsNaN(size) || double.IsInfinity(size) || size <= 0D) {
                throw new System.ArgumentOutOfRangeException(nameof(size), "Font size must be a finite positive number.");
            }

            EnsureTextStyle().Size = size;
            return this;
        }

        /// <summary>Sets whether the whole-shape text is bold.</summary>
        public VisioShape SetBold(bool enabled = true) { EnsureTextStyle().Bold = enabled; return this; }

        /// <summary>Sets whether the whole-shape text is italic.</summary>
        public VisioShape SetItalic(bool enabled = true) { EnsureTextStyle().Italic = enabled; return this; }

        /// <summary>Sets the native underline variant for the whole-shape text.</summary>
        public VisioShape SetUnderline(OfficeTextDecorationStyle style = OfficeTextDecorationStyle.Single) { EnsureTextStyle().UnderlineStyle = style; return this; }

        /// <summary>Sets the native strikethrough variant for the whole-shape text.</summary>
        public VisioShape SetStrikethrough(OfficeTextDecorationStyle style = OfficeTextDecorationStyle.Single) { EnsureTextStyle().StrikethroughStyle = style; return this; }

        /// <summary>Sets the native baseline placement for the whole-shape text.</summary>
        public VisioShape SetBaseline(OfficeTextBaseline baseline) { EnsureTextStyle().Baseline = baseline; return this; }

        /// <summary>Sets Visio's native display-time capitalization for the whole-shape text.</summary>
        public VisioShape SetCapitalization(VisioTextCapitalization capitalization) { EnsureTextStyle().Capitalization = capitalization; return this; }

        /// <summary>Enables or disables Visio's native small-capital style.</summary>
        public VisioShape SetSmallCaps(bool enabled = true) { EnsureTextStyle().SmallCaps = enabled; return this; }

        /// <summary>Sets whole-shape text alignment.</summary>
        public VisioShape SetTextAlignment(OfficeAlignment horizontal, VisioTextVerticalAlignment? vertical = null) {
            VisioTextStyle style = EnsureTextStyle();
            style.HorizontalAlignment = horizontal;
            if (vertical.HasValue) {
                style.VerticalAlignment = vertical.Value;
            }

            return this;
        }

        /// <summary>Adds the shape to a page layer.</summary>
        /// <param name="layerName">Layer name.</param>
        public VisioShape SetLayer(string layerName) { LayerNames.Add(layerName); return this; }

        /// <summary>Sets the shape-level Visio placement style.</summary>
        public VisioShape SetPlacementStyle(VisioPlacementStyle style) { PlacementStyle = style; return this; }

        /// <summary>Sets the shape-level Visio placement flip behavior.</summary>
        public VisioShape SetPlacementFlip(VisioPlacementFlip flip) { PlacementFlip = flip; return this; }

        /// <summary>Sets the shape-level Visio plow behavior.</summary>
        public VisioShape SetPlow(VisioShapePlowCode code) { PlowCode = code; return this; }

        /// <summary>Allows or disallows placing other shapes on top of this shape during layout.</summary>
        public VisioShape SetPlacementOnTop(bool allowed = true) { AllowPlacementOnTop = allowed; return this; }

        /// <summary>Allows or disallows connector routing through this shape.</summary>
        public VisioShape SetConnectorPermeability(bool horizontal = true, bool vertical = true) { AllowHorizontalConnectorRoutingThrough = horizontal; AllowVerticalConnectorRoutingThrough = vertical; return this; }

        /// <summary>Allows or disallows this shape splitting other shapes.</summary>
        public VisioShape SetShapeSplitting(bool canSplit = true, bool canBeSplit = true) { CanSplitShapes = canSplit; CanBeSplit = canBeSplit; return this; }

        private VisioTextStyle EnsureTextStyle() {
            if (TextStyle == null) {
                TextStyle = new VisioTextStyle();
            }

            return TextStyle;
        }
    }
}
