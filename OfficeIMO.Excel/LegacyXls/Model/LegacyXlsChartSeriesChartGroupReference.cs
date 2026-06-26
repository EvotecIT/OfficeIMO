namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a SerToCrt link from a chart series to its containing chart group.
    /// </summary>
    public sealed class LegacyXlsChartSeriesChartGroupReference {
        internal LegacyXlsChartSeriesChartGroupReference(ushort chartGroupIndex) {
            ChartGroupIndex = chartGroupIndex;
        }

        /// <summary>Gets the zero-based index of the referenced ChartFormat record.</summary>
        public ushort ChartGroupIndex { get; }
    }
}
