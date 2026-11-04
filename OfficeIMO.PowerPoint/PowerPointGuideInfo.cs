namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Orientation of a slide guide.
    /// </summary>
    public enum PowerPointGuideOrientation {
        /// <summary>
        ///     Horizontal guide.
        /// </summary>
        Horizontal,
        /// <summary>
        ///     Vertical guide.
        /// </summary>
        Vertical
    }

    /// <summary>
    ///     Represents a slide guide with orientation and position.
    /// </summary>
    public readonly struct PowerPointGuideInfo {
        /// <summary>
        ///     Creates a guide info entry.
        /// </summary>
        public PowerPointGuideInfo(PowerPointGuideOrientation orientation, OfficeLength position) {
            Orientation = orientation;
            Position = position;
        }

        /// <summary>
        ///     Guide orientation.
        /// </summary>
        public PowerPointGuideOrientation Orientation { get; }

        /// <summary>
        ///     Guide position.
        /// </summary>
        public OfficeLength Position { get; }

        /// <summary>
        ///     Returns a display-friendly string.
        /// </summary>
        public override string ToString() {
            return $"{Orientation} @ {Position.ToEmus()}";
        }
    }
}
