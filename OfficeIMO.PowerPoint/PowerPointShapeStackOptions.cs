namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Options for stacking shapes.
    /// </summary>
    public sealed class PowerPointShapeStackOptions {
        /// <summary>
        ///     Spacing between shapes in EMUs.
        /// </summary>
        public long SpacingEmus { get; set; }

        /// <summary>
        ///     Cross-axis alignment. Defaults to Top for horizontal stacks and Left for vertical stacks.
        /// </summary>
        public PowerPointShapeAlignment? Alignment { get; set; }

        /// <summary>
        ///     Justification along the stack axis.
        /// </summary>
        public PowerPointShapeStackJustify Justify { get; set; } = PowerPointShapeStackJustify.Start;

        /// <summary>
        ///     When true, reduces spacing to keep the stack within the bounds.
        /// </summary>
        public bool ClampSpacingToBounds { get; set; }

        /// <summary>
        ///     When true, scales shapes along the stack axis to fit within bounds.
        /// </summary>
        public bool ScaleToFitBounds { get; set; }

        /// <summary>
        ///     When scaling, preserves aspect ratio by scaling both width and height.
        /// </summary>
        public bool PreserveAspect { get; set; }
    }
}
