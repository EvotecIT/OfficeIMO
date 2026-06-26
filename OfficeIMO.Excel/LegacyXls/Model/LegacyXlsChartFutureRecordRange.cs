namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a ChartFrtInfo future-record identifier range.
    /// </summary>
    public sealed class LegacyXlsChartFutureRecordRange {
        internal LegacyXlsChartFutureRecordRange(ushort firstRecordType, ushort lastRecordType) {
            FirstRecordType = firstRecordType;
            LastRecordType = lastRecordType;
        }

        /// <summary>Gets the first future-record type in the range.</summary>
        public ushort FirstRecordType { get; }

        /// <summary>Gets the last future-record type in the range.</summary>
        public ushort LastRecordType { get; }

        /// <summary>Gets a stable range key for reports.</summary>
        public string RangeKey => $"0x{FirstRecordType:X4}-0x{LastRecordType:X4}";
    }
}
