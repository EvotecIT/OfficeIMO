using OfficeIMO.Core.Internal.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a native PowerPoint connection shape.
    /// </summary>
    public class PowerPointConnectionShape : PowerPointShape {
        internal PowerPointConnectionShape(PowerPointSlide slide, OfficeXmlTree tree, int node) : base(slide, tree, node) {
        }

        internal PowerPointConnectionShape(OpenXml.PowerPointPartRef part, OfficeXmlTree tree, int node) : base(part, tree, node) {
        }

        internal PowerPointConnectionShape(ConnectionShape shape, OpenXmlPartContainer? ownerPart = null) : base(shape, ownerPart) {
        }

        private const int Presentation = OfficeXmlNamespaces.PresentationML;
        private const int Drawing = OfficeXmlNamespaces.DrawingML;

        private OfficeXmlNode ShapeProperties => XmlElement.Child(Presentation, "spPr");
        private string? PresetToken => ShapeProperties.Child(Drawing, "prstGeom").Attribute("prst");

        /// <summary>
        ///     Gets the preset geometry type of the connection shape.
        /// </summary>
        public OfficePresetShapeType? ShapeType =>
            PresetToken is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out OfficePresetShapeType type)
                ? type : null;

        /// <summary>Connector family (straight, elbow or curved), or <see langword="null" /> for other line geometry.</summary>
        public PowerPointConnectorKind? Kind =>
            PowerPointConnectorGeometry.KindOf(ShapeType);

        /// <summary>Shape the connector starts at, when its start is glued.</summary>
        public PowerPointShape? StartShape => StartShapeId is uint id ? OwnerSlide?.FindTopLevelShape(id) : null;

        /// <summary>Connection site index the start is glued to.</summary>
        public int? StartSite => (int?)ConnectionEndpoint("stCxn").UInt32Attribute("idx");

        /// <summary>Shape the connector ends at, when its end is glued.</summary>
        public PowerPointShape? EndShape => EndShapeId is uint id ? OwnerSlide?.FindTopLevelShape(id) : null;

        /// <summary>Connection site index the end is glued to.</summary>
        public int? EndSite => (int?)ConnectionEndpoint("endCxn").UInt32Attribute("idx");

        /// <summary>
        ///     Recomputes the connector's path from its glued ends. Moving or resizing a glued shape through this API
        ///     does this automatically; call it after editing shape geometry another way.
        /// </summary>
        public PowerPointConnectionShape Reroute() {
            RouteGlued(Kind ?? PowerPointConnectorKind.Straight);
            return this;
        }

        internal uint? StartShapeId => ConnectionEndpoint("stCxn").UInt32Attribute("id");

        internal uint? EndShapeId => ConnectionEndpoint("endCxn").UInt32Attribute("id");

        private OfficeXmlNode ConnectionEndpoint(string name) => XmlElement.Child(Presentation, "nvCxnSpPr")
            .Child(Presentation, "cNvCxnSpPr").Child(Drawing, name);

        internal void GlueEnds(uint start, int fromSite, uint end, int toSite) {
            OfficeXmlNode drawing = XmlElement.Child(Presentation, "nvCxnSpPr").Child(Presentation, "cNvCxnSpPr");
            OfficeXmlNode startNode = drawing.Append(drawing.CreateElement(Drawing, "stCxn"));
            startNode.SetUInt32Attribute("id", start);
            startNode.SetUInt32Attribute("idx", (uint)fromSite);
            OfficeXmlNode endNode = drawing.Append(drawing.CreateElement(Drawing, "endCxn"));
            endNode.SetUInt32Attribute("id", end);
            endNode.SetUInt32Attribute("idx", (uint)toSite);
            OwnerPartRef.MarkEdited();
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

            PowerPointConnectorGeometry.Route(this, kind, startX, startY, endX, endY, startAngle, endAngle);
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
            var transform = new ATransform2DView(ShapeProperties.Child(Drawing, "xfrm"));
            long x = transform.Off?.Element.Int64Attribute("x") ?? 0L, y = transform.Off?.Element.Int64Attribute("y") ?? 0L;
            long cx = transform.Ext?.Cx ?? 0L, cy = transform.Ext?.Cy ?? 0L;
            bool rotated = Math.Abs((transform.Rot ?? 0) - 5400000) < 2;
            double centerX = x + cx / 2D, centerY = y + cy / 2D;
            double localX = (transform.FlipH == true ? 1D : -1D) * cx / 2D;
            double localY = (transform.FlipV == true ? 1D : -1D) * cy / 2D;
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
