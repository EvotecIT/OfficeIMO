using System;
using OfficeIMO.Drawing;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Visio {
    public partial class VisioConnector {
        /// <summary>Sets the native connector routing kind.</summary>
        public VisioConnector SetKind(ConnectorKind kind) { Kind = kind; return this; }

        /// <summary>Sets a begin arrowhead style.</summary>
        /// <param name="arrow">Arrowhead enum value.</param>
        public VisioConnector SetArrowStart(EndArrow arrow) { BeginArrow = arrow; return this; }

        /// <summary>Sets an end arrowhead style.</summary>
        /// <param name="arrow">Arrowhead enum value.</param>
        public VisioConnector SetArrowEnd(EndArrow arrow) { EndArrow = arrow; return this; }

        /// <summary>Sets a connector label.</summary>
        /// <param name="text">Label text.</param>
        public VisioConnector SetLabel(string text) { Label = text; return this; }

        /// <summary>Sets a connector label and places it along the connector path.</summary>
        /// <param name="text">Label text.</param>
        /// <param name="position">Position along the connector path, from 0.0 to 1.0.</param>
        /// <param name="offsetX">Horizontal page-coordinate offset.</param>
        /// <param name="offsetY">Vertical page-coordinate offset.</param>
        public VisioConnector SetLabel(string text, double position, double offsetX = 0D, double offsetY = 0D) { Label = text; this.PlaceLabel(position, offsetX, offsetY); return this; }

        /// <summary>Sets connector line weight (thickness) in inches.</summary>
        /// <param name="weight">Line weight in inches.</param>
        public VisioConnector SetLineWeight(double weight) { LineWeight = weight; return this; }

        /// <summary>Sets connector line pattern (Visio pattern index).</summary>
        /// <param name="pattern">Pattern index (0=None, 1=Solid, ...).</param>
        public VisioConnector SetLinePattern(int pattern) { LinePattern = pattern; return this; }

        /// <summary>Sets connector line color.</summary>
        /// <param name="color">Line color.</param>
        public VisioConnector SetLineColor(Color color) { LineColor = color; return this; }

        /// <summary>Sets the connector-label font family.</summary>
        public VisioConnector SetFont(string fontFamily) { EnsureTextStyle().FontFamily = fontFamily; return this; }

        /// <summary>Sets the connector-label font size in points.</summary>
        public VisioConnector SetFontSize(double size) {
            if (double.IsNaN(size) || double.IsInfinity(size) || size <= 0D) throw new ArgumentOutOfRangeException(nameof(size));
            EnsureTextStyle().Size = size;
            return this;
        }

        /// <summary>Sets the connector-label text color.</summary>
        public VisioConnector SetTextColor(Color color) { EnsureTextStyle().Color = color; return this; }

        /// <summary>Sets whether the connector label is bold.</summary>
        public VisioConnector SetBold(bool enabled = true) { EnsureTextStyle().Bold = enabled; return this; }

        /// <summary>Sets whether the connector label is italic.</summary>
        public VisioConnector SetItalic(bool enabled = true) { EnsureTextStyle().Italic = enabled; return this; }

        /// <summary>Sets the native connector-label underline variant.</summary>
        public VisioConnector SetUnderline(OfficeTextDecorationStyle style = OfficeTextDecorationStyle.Single) { EnsureTextStyle().UnderlineStyle = style; return this; }

        /// <summary>Sets the native connector-label strikethrough variant.</summary>
        public VisioConnector SetStrikethrough(OfficeTextDecorationStyle style = OfficeTextDecorationStyle.Single) { EnsureTextStyle().StrikethroughStyle = style; return this; }

        /// <summary>Sets the native connector-label baseline placement.</summary>
        public VisioConnector SetBaseline(OfficeTextBaseline baseline) { EnsureTextStyle().Baseline = baseline; return this; }

        /// <summary>Sets the native connector-label capitalization mode.</summary>
        public VisioConnector SetCapitalization(VisioTextCapitalization capitalization) { EnsureTextStyle().Capitalization = capitalization; return this; }

        /// <summary>Adds the connector to a page layer.</summary>
        /// <param name="layerName">Layer name.</param>
        public VisioConnector SetLayer(string layerName) { LayerNames.Add(layerName); return this; }

        /// <summary>Sets the connector-level Visio routing style.</summary>
        public VisioConnector SetRouteStyle(VisioPageRouteStyle style) { RouteStyle = style; return this; }

        /// <summary>Sets the connector-level routed connector appearance.</summary>
        public VisioConnector SetRouteAppearance(VisioLineRouteExtension appearance) { RouteAppearance = appearance; return this; }

        /// <summary>Sets connector-level line jump behavior.</summary>
        public VisioConnector SetLineJumps(
            VisioLineJumpStyle style,
            VisioConnectorLineJumpCode code,
            VisioHorizontalLineJumpDirection horizontalDirection = VisioHorizontalLineJumpDirection.Default,
            VisioVerticalLineJumpDirection verticalDirection = VisioVerticalLineJumpDirection.Default) {
            LineJumpStyle = style;
            LineJumpCode = code;
            HorizontalJumpDirection = horizontalDirection;
            VerticalJumpDirection = verticalDirection;
            return this;
        }

        /// <summary>Sets when Visio may reroute this connector.</summary>
        public VisioConnector SetRerouteBehavior(VisioConnectorRerouteBehavior behavior) { RerouteBehavior = behavior; return this; }

        /// <summary>Connects both ends to explicit shape sides.</summary>
        /// <param name="fromSide">Preferred source side.</param>
        /// <param name="toSide">Preferred target side.</param>
        public VisioConnector Sides(VisioSide fromSide, VisioSide toSide) {
            ValidateSide(From, fromSide);
            ValidateSide(To, toSide);
            ClearEndpointPreservation(true);
            ClearEndpointPreservation(false);
            ApplySide(From, fromSide, point => FromConnectionPoint = point);
            ApplySide(To, toSide, point => ToConnectionPoint = point);
            return this;
        }

        /// <summary>Connects the start of the connector to an explicit side.</summary>
        /// <param name="side">Preferred source side.</param>
        public VisioConnector FromSide(VisioSide side) {
            ValidateSide(From, side);
            ClearEndpointPreservation(true);
            ApplySide(From, side, point => FromConnectionPoint = point);
            return this;
        }

        /// <summary>Connects the end of the connector to an explicit side.</summary>
        /// <param name="side">Preferred target side.</param>
        public VisioConnector ToSide(VisioSide side) {
            ValidateSide(To, side);
            ClearEndpointPreservation(false);
            ApplySide(To, side, point => ToConnectionPoint = point);
            return this;
        }

        private static void ValidateSide(VisioShape? shape, VisioSide side) {
            if (side != VisioSide.Auto && shape == null)
                throw new InvalidOperationException("Attach the connector endpoint before choosing a shape side.");
        }

        private static void ApplySide(VisioShape? shape, VisioSide side, Action<VisioConnectionPoint?> assign) {
            if (side == VisioSide.Auto) {
                assign(null);
                return;
            }

            assign(shape!.EnsureSideConnectionPoint(side));
        }

        private VisioTextStyle EnsureTextStyle() {
            if (TextStyle == null) TextStyle = new VisioTextStyle();
            return TextStyle;
        }
    }
}
