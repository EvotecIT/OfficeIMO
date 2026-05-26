using System;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Fluent helper for configuring a connector between two shapes.
    /// </summary>
    public class VisioFluentConnector {
        private readonly VisioConnector _c;

        /// <summary>Initializes a new connector wrapper.</summary>
        /// <param name="connector">Underlying connector model.</param>
        internal VisioFluentConnector(VisioConnector connector) { _c = connector; }

        /// <summary>Sets connector kind to a straight line.</summary>
        public VisioFluentConnector Straight() { _c.Kind = ConnectorKind.Straight; return this; }

        /// <summary>Sets connector kind to right-angle (orthogonal) routing.</summary>
        public VisioFluentConnector RightAngle() { _c.Kind = ConnectorKind.RightAngle; return this; }

        /// <summary>Sets connector kind to curved routing.</summary>
        public VisioFluentConnector Curved() { _c.Kind = ConnectorKind.Curved; return this; }

        /// <summary>Sets connector kind to dynamic routing.</summary>
        public VisioFluentConnector Dynamic() { _c.Kind = ConnectorKind.Dynamic; return this; }

        /// <summary>Sets a begin arrowhead style.</summary>
        /// <param name="arrow">Arrowhead enum value.</param>
        public VisioFluentConnector ArrowStart(EndArrow arrow) { _c.BeginArrow = arrow; return this; }

        /// <summary>Sets an end arrowhead style.</summary>
        /// <param name="arrow">Arrowhead enum value.</param>
        public VisioFluentConnector ArrowEnd(EndArrow arrow) { _c.EndArrow = arrow; return this; }

        /// <summary>Sets a connector label.</summary>
        /// <param name="text">Label text.</param>
        public VisioFluentConnector Label(string text) { _c.Label = text; return this; }

        /// <summary>Sets a connector label and places it along the connector path.</summary>
        /// <param name="text">Label text.</param>
        /// <param name="position">Position along the connector path, from 0.0 to 1.0.</param>
        /// <param name="offsetX">Horizontal page-coordinate offset.</param>
        /// <param name="offsetY">Vertical page-coordinate offset.</param>
        public VisioFluentConnector Label(string text, double position, double offsetX = 0D, double offsetY = 0D) { _c.Label = text; _c.PlaceLabel(position, offsetX, offsetY); return this; }

        /// <summary>Sets connector line weight (thickness) in inches.</summary>
        /// <param name="weight">Line weight in inches.</param>
        public VisioFluentConnector LineWeight(double weight) { _c.LineWeight = weight; return this; }

        /// <summary>Sets connector line pattern (Visio pattern index).</summary>
        /// <param name="pattern">Pattern index (0=None, 1=Solid, ...).</param>
        public VisioFluentConnector LinePattern(int pattern) { _c.LinePattern = pattern; return this; }

        /// <summary>Sets connector line color.</summary>
        /// <param name="color">Line color.</param>
        public VisioFluentConnector LineColor(Color color) { _c.LineColor = color; return this; }

        /// <summary>Routes the connector through explicit page-coordinate waypoints.</summary>
        /// <param name="waypoints">Absolute page coordinates between start and end.</param>
        public VisioFluentConnector RouteThrough(params VisioConnectorWaypoint[] waypoints) { _c.RouteThrough(waypoints); return this; }

        /// <summary>Generates a clean orthogonal route for the connector.</summary>
        /// <param name="style">Orthogonal route orientation.</param>
        /// <param name="offset">Optional offset applied to the center routing lane.</param>
        public VisioFluentConnector RouteOrthogonal(VisioConnectorRouteStyle style = VisioConnectorRouteStyle.Auto, double offset = 0D) { _c.RouteOrthogonal(style, offset); return this; }

        /// <summary>Removes explicit waypoints and returns to dynamic routing.</summary>
        public VisioFluentConnector ClearRoute() { _c.ClearRoute(); return this; }

        /// <summary>Places connector text along the connector path.</summary>
        /// <param name="position">Position along the connector path, from 0.0 to 1.0.</param>
        /// <param name="offsetX">Horizontal page-coordinate offset.</param>
        /// <param name="offsetY">Vertical page-coordinate offset.</param>
        /// <param name="width">Label text box width in page units.</param>
        /// <param name="height">Label text box height in page units.</param>
        public VisioFluentConnector LabelPosition(double position = 0.5D, double offsetX = 0D, double offsetY = 0D, double width = 1.25D, double height = 0.3D) { _c.PlaceLabel(position, offsetX, offsetY, width, height); return this; }

        /// <summary>Places connector text at an absolute page coordinate.</summary>
        /// <param name="pinX">Text pin X coordinate.</param>
        /// <param name="pinY">Text pin Y coordinate.</param>
        /// <param name="width">Label text box width in page units.</param>
        /// <param name="height">Label text box height in page units.</param>
        public VisioFluentConnector LabelAt(double pinX, double pinY, double width = 1.25D, double height = 0.3D) { _c.PlaceLabelAt(pinX, pinY, width, height); return this; }

        /// <summary>Resizes the connector label text box to fit its label text.</summary>
        /// <param name="fontInfo">Font descriptor used for measurement. Uses connector text style, then Office default, when omitted.</param>
        /// <param name="maximumWidth">Optional maximum label width in inches.</param>
        public VisioFluentConnector AutoSizeLabel(OfficeFontInfo? fontInfo = null, double? maximumWidth = null) { _c.ResizeLabelToText(fontInfo, maximumWidth: maximumWidth); return this; }

        /// <summary>Applies a reusable connector style.</summary>
        /// <param name="style">Connector style to apply.</param>
        public VisioFluentConnector Style(VisioConnectorStyle style) { _c.ApplyStyle(style); return this; }

        /// <summary>Applies reusable text formatting to the connector label.</summary>
        /// <param name="style">Text style to apply.</param>
        public VisioFluentConnector TextStyle(VisioTextStyle style) { _c.ApplyTextStyle(style); return this; }

        /// <summary>Adds the connector to a page layer.</summary>
        /// <param name="layerName">Layer name.</param>
        public VisioFluentConnector Layer(string layerName) { _c.LayerNames.Add(layerName); return this; }

        /// <summary>Adds a hyperlink to the connector.</summary>
        /// <param name="address">External hyperlink address.</param>
        /// <param name="description">Optional display description.</param>
        /// <param name="subAddress">Optional internal sub-address.</param>
        public VisioFluentConnector Hyperlink(string address, string? description = null, string? subAddress = null) { _c.AddHyperlink(address, description, subAddress); return this; }

        /// <summary>Configures ShapeSheet protection cells.</summary>
        /// <param name="configure">Protection configuration delegate.</param>
        public VisioFluentConnector Protect(Action<VisioProtection> configure) { _c.Protect(configure); return this; }

        /// <summary>Sets the connector-level Visio routing style.</summary>
        public VisioFluentConnector RouteStyle(VisioPageRouteStyle style) { _c.RouteStyle = style; return this; }

        /// <summary>Sets the connector-level routed connector appearance.</summary>
        public VisioFluentConnector RouteAppearance(VisioLineRouteExtension appearance) { _c.RouteAppearance = appearance; return this; }

        /// <summary>Sets connector-level line jump behavior.</summary>
        public VisioFluentConnector LineJumps(
            VisioLineJumpStyle style,
            VisioConnectorLineJumpCode code,
            VisioHorizontalLineJumpDirection horizontalDirection = VisioHorizontalLineJumpDirection.Default,
            VisioVerticalLineJumpDirection verticalDirection = VisioVerticalLineJumpDirection.Default) {
            _c.LineJumpStyle = style;
            _c.LineJumpCode = code;
            _c.HorizontalJumpDirection = horizontalDirection;
            _c.VerticalJumpDirection = verticalDirection;
            return this;
        }

        /// <summary>Sets when Visio may reroute this connector.</summary>
        public VisioFluentConnector RerouteBehavior(VisioConnectorRerouteBehavior behavior) { _c.RerouteBehavior = behavior; return this; }

        /// <summary>Clears explicit Shape Layout routing override cells.</summary>
        public VisioFluentConnector ClearRoutingPolicy() { _c.ClearRoutingPolicy(); return this; }

        /// <summary>Locks or unlocks connector endpoints.</summary>
        /// <param name="locked">Whether endpoints are locked.</param>
        public VisioFluentConnector LockEndpoints(bool locked = true) { _c.LockEndpoints(locked); return this; }

        /// <summary>Connects both ends to explicit shape sides.</summary>
        /// <param name="fromSide">Preferred source side.</param>
        /// <param name="toSide">Preferred target side.</param>
        public VisioFluentConnector Sides(VisioSide fromSide, VisioSide toSide) {
            ApplySide(_c.From, fromSide, point => _c.FromConnectionPoint = point);
            ApplySide(_c.To, toSide, point => _c.ToConnectionPoint = point);
            return this;
        }

        /// <summary>Connects the start of the connector to an explicit side.</summary>
        /// <param name="side">Preferred source side.</param>
        public VisioFluentConnector FromSide(VisioSide side) {
            ApplySide(_c.From, side, point => _c.FromConnectionPoint = point);
            return this;
        }

        /// <summary>Connects the end of the connector to an explicit side.</summary>
        /// <param name="side">Preferred target side.</param>
        public VisioFluentConnector ToSide(VisioSide side) {
            ApplySide(_c.To, side, point => _c.ToConnectionPoint = point);
            return this;
        }

        private static void ApplySide(VisioShape shape, VisioSide side, Action<VisioConnectionPoint?> assign) {
            if (side == VisioSide.Auto) {
                assign(null);
                return;
            }

            assign(shape.EnsureSideConnectionPoint(side));
        }
    }
}

