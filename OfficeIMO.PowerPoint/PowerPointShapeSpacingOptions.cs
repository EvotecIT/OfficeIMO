namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Options for fixed spacing distribution.
    /// </summary>
    public sealed class PowerPointShapeSpacingOptions {
        /// <summary>
        ///     Spacing between shapes in EMUs.
        /// </summary>
        public long SpacingEmus { get; set; }

        /// <summary>
        ///     Main-axis alignment (Left/Center/Right for horizontal; Top/Middle/Bottom for vertical).
        /// </summary>
        public PowerPointShapeAlignment? Alignment { get; set; }

        /// <summary>
        ///     Cross-axis alignment (Top/Middle/Bottom for horizontal; Left/Center/Right for vertical).
        /// </summary>
        public PowerPointShapeAlignment? CrossAxisAlignment { get; set; }

        /// <summary>
        ///     When true, reduces spacing to keep the block within the bounds.
        /// </summary>
        public bool ClampSpacingToBounds { get; set; }

        /// <summary>
        ///     When true, scales shapes along the distribution axis to fit within bounds.
        /// </summary>
        public bool ScaleToFitBounds { get; set; }

        /// <summary>
        ///     When scaling, preserves aspect ratio by scaling both width and height.
        /// </summary>
        public bool PreserveAspect { get; set; }
    }
}
