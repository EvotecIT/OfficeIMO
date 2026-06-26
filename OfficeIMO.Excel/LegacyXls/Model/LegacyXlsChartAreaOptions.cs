namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes area chart group options decoded from a legacy XLS Area record.
    /// </summary>
    public sealed class LegacyXlsChartAreaOptions {
        internal LegacyXlsChartAreaOptions(ushort flags) {
            Flags = flags;
        }

        /// <summary>Gets the raw option flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the area chart group is stacked.</summary>
        public bool IsStacked => (Flags & 0x0001) != 0;

        /// <summary>Gets whether the area chart group is 100 percent stacked.</summary>
        public bool IsPercentStacked => (Flags & 0x0002) != 0;

        /// <summary>Gets whether the area chart group has a shadow.</summary>
        public bool HasShadow => (Flags & 0x0004) != 0;

        /// <summary>Gets whether the percentage-stacked flag is valid for the decoded stacking state.</summary>
        public bool HasValidPercentStackedState => IsStacked || !IsPercentStacked;

        /// <summary>Gets whether reserved option bits are zero.</summary>
        public bool HasZeroReservedBits => (Flags & 0xfff8) == 0;
    }
}
