namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from a chart ObjectLink record.
    /// </summary>
    public sealed class LegacyXlsChartObjectLink {
        internal LegacyXlsChartObjectLink(ushort linkedObject, string linkedObjectName, ushort seriesIndex, ushort dataPointIndex) {
            LinkedObject = linkedObject;
            LinkedObjectName = linkedObjectName ?? throw new ArgumentNullException(nameof(linkedObjectName));
            SeriesIndex = seriesIndex;
            DataPointIndex = dataPointIndex;
        }

        /// <summary>Gets the raw linked chart object identifier.</summary>
        public ushort LinkedObject { get; }

        /// <summary>Gets the decoded linked chart object name.</summary>
        public string LinkedObjectName { get; }

        /// <summary>Gets the zero-based series index for series/data-point links.</summary>
        public ushort SeriesIndex { get; }

        /// <summary>Gets the zero-based data-point index, or 0xFFFF when the link targets a whole series.</summary>
        public ushort DataPointIndex { get; }
    }
}
