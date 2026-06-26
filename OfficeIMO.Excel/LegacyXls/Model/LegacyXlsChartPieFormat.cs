namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes pie or doughnut chart explosion metadata decoded from a legacy XLS PieFormat record.
    /// </summary>
    public sealed class LegacyXlsChartPieFormat {
        internal LegacyXlsChartPieFormat(short explosionPercentage) {
            ExplosionPercentage = explosionPercentage;
        }

        /// <summary>
        /// Gets the distance of the data point or series from the chart center as a percentage.
        /// </summary>
        public short ExplosionPercentage { get; }
    }
}
