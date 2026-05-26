using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Options for creating an OfficeIMO callout shape and leader connector.
    /// </summary>
    public sealed class VisioCalloutOptions {
        /// <summary>
        /// Width of the callout shape in page units.
        /// </summary>
        public double Width { get; set; } = 2.2D;

        /// <summary>
        /// Height of the callout shape in page units.
        /// </summary>
        public double Height { get; set; } = 0.75D;

        /// <summary>
        /// Optional page layer for the callout shape and leader connector. Set to null or empty to skip layer assignment.
        /// </summary>
        public string? LayerName { get; set; } = "Annotations";

        /// <summary>
        /// Preferred side on the target shape. Auto chooses the side nearest the callout.
        /// </summary>
        public VisioSide TargetSide { get; set; } = VisioSide.Auto;

        /// <summary>
        /// Preferred side on the callout shape. Auto chooses the side facing the target.
        /// </summary>
        public VisioSide CalloutSide { get; set; } = VisioSide.Auto;

        /// <summary>
        /// Connector kind used for the callout leader.
        /// </summary>
        public ConnectorKind LeaderKind { get; set; } = ConnectorKind.RightAngle;

        /// <summary>
        /// Whether to generate explicit orthogonal waypoints for right-angle leaders.
        /// </summary>
        public bool RouteLeader { get; set; } = true;

        /// <summary>
        /// Optional offset for the generated leader routing lane.
        /// </summary>
        public double RouteOffset { get; set; }

        /// <summary>
        /// Optional style for the callout shape. When null, a warm annotation style is used.
        /// </summary>
        public VisioShapeStyle? ShapeStyle { get; set; }

        /// <summary>
        /// Optional style for the leader connector. When null, a dashed annotation leader is used.
        /// </summary>
        public VisioConnectorStyle? LeaderStyle { get; set; }

        internal VisioShapeStyle GetShapeStyle() {
            return ShapeStyle ?? new VisioShapeStyle(Color.FromRgb(255, 248, 225), Color.FromRgb(166, 124, 0), 0.014D);
        }

        internal VisioConnectorStyle GetLeaderStyle() {
            return LeaderStyle ?? new VisioConnectorStyle(Color.FromRgb(166, 124, 0), 0.012D, 2, EndArrow.None) {
                Kind = LeaderKind
            };
        }
    }
}
