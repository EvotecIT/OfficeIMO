namespace OfficeIMO.Visio {
    /// <summary>
    /// Options controlling swimlane activity move and relayout operations.
    /// </summary>
    public sealed class VisioSwimlaneRelayoutOptions {
        /// <summary>Vertical gap between activities stacked in the same lane/phase cell.</summary>
        public double ActivityGap { get; set; } = 0.22D;

        /// <summary>Whether connectors attached to swimlane activities should be rerouted after movement.</summary>
        public bool RerouteConnectors { get; set; } = true;

        /// <summary>Whether rerouting should try to avoid other page shapes.</summary>
        public bool AvoidShapes { get; set; } = true;

        /// <summary>Padding used when obstacle-aware routing is enabled.</summary>
        public double RoutingPadding { get; set; } = 0.12D;

        /// <summary>Maximum routing lane search depth used when obstacle-aware routing is enabled.</summary>
        public int MaxRoutingLanes { get; set; } = 8;
    }
}
