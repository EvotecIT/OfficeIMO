namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded Chart3d options preserved from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChart3DOptions {
        internal LegacyXlsChart3DOptions(
            short rotationDegrees,
            short elevationDegrees,
            short fieldOfViewDegrees,
            ushort heightOrThicknessPercent,
            short depthPercent,
            ushort gapWidthPercent,
            ushort flags) {
            RotationDegrees = rotationDegrees;
            ElevationDegrees = elevationDegrees;
            FieldOfViewDegrees = fieldOfViewDegrees;
            HeightOrThicknessPercent = heightOrThicknessPercent;
            DepthPercent = depthPercent;
            GapWidthPercent = gapWidthPercent;
            Flags = flags;
        }

        /// <summary>Gets the clockwise rotation, in degrees, around the vertical center line.</summary>
        public short RotationDegrees { get; }

        /// <summary>Gets the elevation, in degrees, around the horizontal center line.</summary>
        public short ElevationDegrees { get; }

        /// <summary>Gets the 3-D plot area field-of-view angle.</summary>
        public short FieldOfViewDegrees { get; }

        /// <summary>Gets the raw height or pie-thickness percentage value.</summary>
        public ushort HeightOrThicknessPercent { get; }

        /// <summary>Gets the non-pie 3-D plot area height percentage interpretation.</summary>
        public short HeightPercent => unchecked((short)HeightOrThicknessPercent);

        /// <summary>Gets the 3-D plot area depth percentage.</summary>
        public short DepthPercent { get; }

        /// <summary>Gets the gap width percentage.</summary>
        public ushort GapWidthPercent { get; }

        /// <summary>Gets the raw Chart3d flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the 3-D plot area is rendered with perspective.</summary>
        public bool UsesPerspective => (Flags & 0x0001) != 0;

        /// <summary>Gets whether bar-chart data points are clustered.</summary>
        public bool IsClustered => (Flags & 0x0002) != 0;

        /// <summary>Gets whether the 3-D plot area height is automatically determined.</summary>
        public bool UsesAutomaticScaling => (Flags & 0x0004) != 0;

        /// <summary>Gets whether the chart group is not a pie chart.</summary>
        public bool IsNotPieChart => (Flags & 0x0010) != 0;

        /// <summary>Gets whether chart walls are rendered in 2-D.</summary>
        public bool UsesTwoDimensionalWalls => (Flags & 0x0020) != 0;

        /// <summary>Gets whether the reserved Chart3d flag bits are zero.</summary>
        public bool HasZeroReservedBits => (Flags & 0xffc8) == 0;

        /// <summary>Gets the decoded chart group shape state.</summary>
        public string ChartGroupShapeName => IsNotPieChart ? "NotPie" : "Pie";
    }
}
