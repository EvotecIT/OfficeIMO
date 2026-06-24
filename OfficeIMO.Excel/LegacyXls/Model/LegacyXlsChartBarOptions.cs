namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded Bar chart group options preserved from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartBarOptions {
        internal LegacyXlsChartBarOptions(short overlapPercentage, ushort gapWidthPercentage, ushort flags) {
            OverlapPercentage = overlapPercentage;
            GapWidthPercentage = gapWidthPercentage;
            Flags = flags;
        }

        /// <summary>Gets the overlap between data points in the same category as a percentage of bar width.</summary>
        public short OverlapPercentage { get; }

        /// <summary>Gets the gap width between adjacent categories as a percentage of bar width.</summary>
        public ushort GapWidthPercentage { get; }

        /// <summary>Gets the raw Bar option flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the chart group swaps category and value axis orientation.</summary>
        public bool IsTransposed => (Flags & 0x0001) != 0;

        /// <summary>Gets whether data points in the same chart group are stacked.</summary>
        public bool IsStacked => (Flags & 0x0002) != 0;

        /// <summary>Gets whether stacked values are displayed as percentages.</summary>
        public bool IsPercentStacked => (Flags & 0x0004) != 0;

        /// <summary>Gets whether data points in the chart group have shadows.</summary>
        public bool HasShadow => (Flags & 0x0008) != 0;
    }
}
