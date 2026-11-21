using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a native PowerPoint connection shape.
    /// </summary>
    public class PowerPointConnectionShape : PowerPointShape {
        internal PowerPointConnectionShape(ConnectionShape shape, OpenXmlPartContainer? ownerPart = null) : base(shape, ownerPart) {
        }

        private ConnectionShape Shape => (ConnectionShape)Element;

        /// <summary>
        ///     Gets the preset geometry type of the connection shape.
        /// </summary>
        public OfficePresetShapeType? ShapeType => Shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value.ToOfficeEnum();

        /// <summary>Connector family (straight, elbow or curved), or <see langword="null" /> for other line geometry.</summary>
        public PowerPointConnectorKind? Kind =>
            PowerPointConnectorGeometry.KindOf(Shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value);

        /// <summary>Shape the connector starts at, when its start is glued.</summary>
        public PowerPointShape? StartShape => StartShapeId is uint id ? OwnerSlide?.FindTopLevelShape(id) : null;

        /// <summary>Connection site index the start is glued to.</summary>
        public int? StartSite => (int?)ConnectorDrawingProperties.StartConnection?.Index?.Value;

        /// <summary>Shape the connector ends at, when its end is glued.</summary>
        public PowerPointShape? EndShape => EndShapeId is uint id ? OwnerSlide?.FindTopLevelShape(id) : null;

        /// <summary>Connection site index the end is glued to.</summary>
        public int? EndSite => (int?)ConnectorDrawingProperties.EndConnection?.Index?.Value;

        /// <summary>
        ///     Recomputes the connector's path from its glued ends. Moving or resizing a glued shape through this API
        ///     does this automatically; call it after editing shape geometry another way.
        /// </summary>
        public PowerPointConnectionShape Reroute() {
            RouteGlued(Kind ?? PowerPointConnectorKind.Straight);
            return this;
        }

        internal uint? StartShapeId => ConnectorDrawingProperties.StartConnection?.Id?.Value;

        internal uint? EndShapeId => ConnectorDrawingProperties.EndConnection?.Id?.Value;

        internal NonVisualConnectorShapeDrawingProperties ConnectorDrawingProperties {
            get {
                NonVisualConnectionShapeProperties properties = Shape.NonVisualConnectionShapeProperties
                    ??= new NonVisualConnectionShapeProperties();
                return properties.NonVisualConnectorShapeDrawingProperties ??= new NonVisualConnectorShapeDrawingProperties();
            }
        }

        internal void RouteGlued(PowerPointConnectorKind kind) => RouteGlued(kind, strict: true);

        /// <summary>
        ///     Reroutes after a glued shape changed. A glued end whose sites cannot be resolved (for example a preset
        ///     without known sites in a loaded deck) leaves the connector as stored instead of failing the edit.
        /// </summary>
        internal void RerouteAfterShapeChange() => RouteGlued(Kind ?? PowerPointConnectorKind.Straight, strict: false);

        private void RouteGlued(PowerPointConnectorKind kind, bool strict) {
            PowerPointShape? start = StartShape, end = EndShape;
            if (start == null && end == null) return;

            // An unglued end keeps its current position.
            GetEndPoints(out long startX, out long startY, out long endX, out long endY);
            double? startAngle = null, endAngle = null;
            if (start != null && StartSite is int startIndex) {
                if (!TryResolveSite(start, startIndex, nameof(StartSite), strict, out PowerPointConnectionSite site)) return;
                startX = site.X.ToEmus();
                startY = site.Y.ToEmus();
                startAngle = site.Angle;
            }

            if (end != null && EndSite is int endIndex) {
                if (!TryResolveSite(end, endIndex, nameof(EndSite), strict, out PowerPointConnectionSite site)) return;
                endX = site.X.ToEmus();
                endY = site.Y.ToEmus();
                endAngle = site.Angle;
            }

            PowerPointConnectorGeometry.Route(Shape, kind, startX, startY, endX, endY, startAngle, endAngle);
        }

        private static bool TryResolveSite(PowerPointShape shape, int index, string parameterName, bool strict,
            out PowerPointConnectionSite site) {
            if (strict) {
                site = PowerPointConnectorGeometry.GetSite(shape, index, parameterName);
                return true;
            }

            return PowerPointConnectorGeometry.TryGetSite(shape, index, out site);
        }

        private void GetEndPoints(out long startX, out long startY, out long endX, out long endY) {
            A.Transform2D? transform = Shape.ShapeProperties?.Transform2D;
            long x = transform?.Offset?.X?.Value ?? 0L, y = transform?.Offset?.Y?.Value ?? 0L;
            long cx = transform?.Extents?.Cx?.Value ?? 0L, cy = transform?.Extents?.Cy?.Value ?? 0L;
            bool rotated = Math.Abs((transform?.Rotation?.Value ?? 0) - 5400000) < 2;
            double centerX = x + cx / 2D, centerY = y + cy / 2D;
            double localX = (transform?.HorizontalFlip?.Value == true ? 1D : -1D) * cx / 2D;
            double localY = (transform?.VerticalFlip?.Value == true ? 1D : -1D) * cy / 2D;
            // Local start corner, rotated by 90° when the connector is drawn rotated.
            double sx = rotated ? -localY : localX, sy = rotated ? localX : localY;
            startX = (long)Math.Round(centerX + sx);
            startY = (long)Math.Round(centerY + sy);
            endX = (long)Math.Round(centerX - sx);
            endY = (long)Math.Round(centerY - sy);
        }

        /// <summary>
        ///     Sets the outline color (and optional physical width) and returns the shape for chaining.
        /// </summary>
        public PowerPointConnectionShape Stroke(OfficeColor color, OfficeLength? width = null) {
            if (width.HasValue) ToDrawingLineWidth(width.Value, nameof(width));
            OutlineColor = color;
            if (width.HasValue) {
                OutlineWidth = width;
            }

            return this;
        }
    }
}
