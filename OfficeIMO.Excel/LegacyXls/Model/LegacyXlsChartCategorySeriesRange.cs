namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes category, date, or series axis interval metadata decoded from a legacy XLS CatSerRange record.
    /// </summary>
    public sealed class LegacyXlsChartCategorySeriesRange {
        internal LegacyXlsChartCategorySeriesRange(short crossingCategory, short labelInterval, short tickInterval, ushort flags) {
            CrossingCategory = crossingCategory;
            LabelInterval = labelInterval;
            TickInterval = tickInterval;
            Flags = flags;
        }

        /// <summary>Gets the raw crossing category, series, or date-derived crossing value.</summary>
        public short CrossingCategory { get; }

        /// <summary>Gets the interval between axis labels.</summary>
        public short LabelInterval { get; }

        /// <summary>Gets the interval between visible major or minor tick marks.</summary>
        public short TickInterval { get; }

        /// <summary>Gets the raw CatSerRange flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the value axis crosses this axis between major tick marks.</summary>
        public bool CrossesBetweenTickMarks => (Flags & 0x0001) != 0;

        /// <summary>Gets whether the value axis crosses this axis at the last category, series, or maximum date.</summary>
        public bool CrossesAtMaximum => (Flags & 0x0002) != 0;

        /// <summary>Gets whether this axis is displayed in reverse order.</summary>
        public bool Reversed => (Flags & 0x0004) != 0;
    }
}
