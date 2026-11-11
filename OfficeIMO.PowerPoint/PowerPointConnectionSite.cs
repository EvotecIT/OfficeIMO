namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     A point on a shape where a connector can be glued, taken from the shape's geometry.
    /// </summary>
    public readonly struct PowerPointConnectionSite {
        internal PowerPointConnectionSite(int index, OfficeLength x, OfficeLength y, double angle) {
            Index = index;
            X = x;
            Y = y;
            Angle = angle;
        }

        /// <summary>Site index stored in the connector's glue.</summary>
        public int Index { get; }

        /// <summary>Horizontal position on the slide.</summary>
        public OfficeLength X { get; }

        /// <summary>Vertical position on the slide.</summary>
        public OfficeLength Y { get; }

        /// <summary>Direction a connector leaves the site, in degrees clockwise from the positive x-axis (0 right, 90 down).</summary>
        public double Angle { get; }
    }

    /// <summary>
    ///     Connector geometry family.
    /// </summary>
    public enum PowerPointConnectorKind {
        /// <summary>A straight line between the two ends.</summary>
        Straight,
        /// <summary>Horizontal and vertical segments with right-angle bends.</summary>
        Elbow,
        /// <summary>A smooth curve.</summary>
        Curved
    }
}
