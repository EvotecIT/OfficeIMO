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
    ///     Represents a slide guide with orientation and position (EMUs).
    /// </summary>
    public readonly struct PowerPointGuideInfo {
        /// <summary>
        ///     Creates a guide info entry.
        /// </summary>
        public PowerPointGuideInfo(PowerPointGuideOrientation orientation, long positionEmus) {
            Orientation = orientation;
            PositionEmus = positionEmus;
        }

        /// <summary>
        ///     Guide orientation.
        /// </summary>
        public PowerPointGuideOrientation Orientation { get; }

        /// <summary>
        ///     Guide position in EMUs.
        /// </summary>
        public long PositionEmus { get; }

        /// <summary>
        ///     Returns a display-friendly string.
        /// </summary>
        public override string ToString() {
            return $"{Orientation} @ {PositionEmus}";
        }
    }
}
