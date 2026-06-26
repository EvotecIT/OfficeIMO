namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes decoded Dat chart data-table display options.
    /// </summary>
    public sealed class LegacyXlsChartDataTableOptions {
        internal LegacyXlsChartDataTableOptions(ushort flags) {
            Flags = flags;
        }

        /// <summary>Gets the raw Dat option flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether horizontal borders are displayed within the data table.</summary>
        public bool HasHorizontalBorders => (Flags & 0x0001) != 0;

        /// <summary>Gets whether vertical borders are displayed within the data table.</summary>
        public bool HasVerticalBorders => (Flags & 0x0002) != 0;

        /// <summary>Gets whether an outside outline is displayed around the data table.</summary>
        public bool HasOutlineBorder => (Flags & 0x0004) != 0;

        /// <summary>Gets whether legend key symbols are displayed next to series names.</summary>
        public bool ShowSeriesKeys => (Flags & 0x0008) != 0;

        /// <summary>Gets Dat option bits outside the currently decoded display flags.</summary>
        public ushort ReservedFlags => (ushort)(Flags & 0xFFF0);

        /// <summary>Gets whether Dat option bits outside the currently decoded display flags are set.</summary>
        public bool HasReservedFlags => ReservedFlags != 0;

        /// <summary>Gets a compact reserved-bit state for corpus diagnostics.</summary>
        public string ReservedState => HasReservedFlags ? $"Reserved:0x{ReservedFlags:X4}" : "ReservedClear";
    }
}
