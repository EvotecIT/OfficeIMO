namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Grid dimensions used by <see cref="VisioBlockDiagramBuilder"/>. Visual styling is owned by
    /// <see cref="VisioStyleTheme"/>.
    /// </summary>
    public sealed class VisioBlockDiagramLayoutOptions {
        /// <summary>Default block width in page units.</summary>
        public double BlockWidth { get; set; } = 2.35;

        /// <summary>Default block height in page units.</summary>
        public double BlockHeight { get; set; } = 1.0;

        /// <summary>Default column gap in page units.</summary>
        public double ColumnGap { get; set; } = 1.1;

        /// <summary>Default row gap in page units.</summary>
        public double RowGap { get; set; } = 0.8;

        /// <summary>Creates a detached copy of these options.</summary>
        public VisioBlockDiagramLayoutOptions Clone() => new VisioBlockDiagramLayoutOptions {
            BlockWidth = BlockWidth,
            BlockHeight = BlockHeight,
            ColumnGap = ColumnGap,
            RowGap = RowGap
        };
    }
}
