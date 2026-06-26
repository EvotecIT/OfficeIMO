namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded ValueRange chart-axis scaling metadata.
    /// </summary>
    public sealed class LegacyXlsChartValueRange {
        internal LegacyXlsChartValueRange(
            double minimum,
            double maximum,
            double majorUnit,
            double minorUnit,
            double crossingValue,
            ushort flags) {
            Minimum = minimum;
            Maximum = maximum;
            MajorUnit = majorUnit;
            MinorUnit = minorUnit;
            CrossingValue = crossingValue;
            Flags = flags;
        }

        /// <summary>Gets the raw minimum axis value.</summary>
        public double Minimum { get; }

        /// <summary>Gets the raw maximum axis value.</summary>
        public double Maximum { get; }

        /// <summary>Gets the raw major tick interval.</summary>
        public double MajorUnit { get; }

        /// <summary>Gets the raw minor tick interval.</summary>
        public double MinorUnit { get; }

        /// <summary>Gets the raw axis crossing value.</summary>
        public double CrossingValue { get; }

        /// <summary>Gets the raw ValueRange flags bitfield.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the minimum value is calculated automatically.</summary>
        public bool AutoMinimum => (Flags & 0x0001) != 0;

        /// <summary>Gets whether the maximum value is calculated automatically.</summary>
        public bool AutoMaximum => (Flags & 0x0002) != 0;

        /// <summary>Gets whether the major interval is calculated automatically.</summary>
        public bool AutoMajorUnit => (Flags & 0x0004) != 0;

        /// <summary>Gets whether the minor interval is calculated automatically.</summary>
        public bool AutoMinorUnit => (Flags & 0x0008) != 0;

        /// <summary>Gets whether the crossing value is calculated automatically.</summary>
        public bool AutoCrossingValue => (Flags & 0x0010) != 0;

        /// <summary>Gets whether the value axis uses a logarithmic scale.</summary>
        public bool LogarithmicScale => (Flags & 0x0020) != 0;

        /// <summary>Gets whether axis values are displayed in reverse order.</summary>
        public bool Reversed => (Flags & 0x0040) != 0;

        /// <summary>Gets whether the other axes cross this value axis at its maximum value.</summary>
        public bool MaximumCrossing => (Flags & 0x0080) != 0;
    }
}
