namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Options for automatic grid layout.
    /// </summary>
    public sealed class PowerPointShapeGridOptions {
        /// <summary>Minimum number of columns to use.</summary>
        public int? MinColumns { get; set; }
        /// <summary>Maximum number of columns to use.</summary>
        public int? MaxColumns { get; set; }
        /// <summary>Target cell aspect ratio (width / height).</summary>
        public double? TargetCellAspect { get; set; }
        /// <summary>Horizontal gutter between cells (EMUs).</summary>
        public long GutterX { get; set; }
        /// <summary>Vertical gutter between cells (EMUs).</summary>
        public long GutterY { get; set; }
        /// <summary>Whether shapes should be resized to the grid cell.</summary>
        public bool ResizeToCell { get; set; } = true;
        /// <summary>Flow direction for filling cells.</summary>
        public PowerPointShapeGridFlow Flow { get; set; } = PowerPointShapeGridFlow.RowMajor;
    }
}
