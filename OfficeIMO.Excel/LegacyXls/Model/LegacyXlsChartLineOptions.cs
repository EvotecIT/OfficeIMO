namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded Line chart group options preserved from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartLineOptions {
        internal LegacyXlsChartLineOptions(ushort flags) {
            Flags = flags;
        }

        /// <summary>Gets the raw Line chart group flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether data points sharing the same category are stacked.</summary>
        public bool IsStacked => (Flags & 0x0001) != 0;

        /// <summary>Gets whether stacked values are displayed as percentages.</summary>
        public bool IsPercentStacked => (Flags & 0x0002) != 0;

        /// <summary>Gets whether one or more data markers in the chart group have shadows.</summary>
        public bool HasShadow => (Flags & 0x0004) != 0;

        /// <summary>Gets whether the percentage-stacked bit is valid for the decoded stacked state.</summary>
        public bool HasValidPercentStackedState => IsStacked || !IsPercentStacked;

        /// <summary>Gets whether the reserved Line flag bits are zero.</summary>
        public bool HasZeroReservedBits => (Flags & 0xfff8) == 0;
    }
}
