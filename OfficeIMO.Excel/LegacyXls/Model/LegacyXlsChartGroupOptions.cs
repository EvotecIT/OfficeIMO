namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only ChartFormat chart-group metadata from a BIFF chart stream.
    /// </summary>
    public sealed class LegacyXlsChartGroupOptions {
        internal LegacyXlsChartGroupOptions(bool variedDataPointColors, ushort drawingOrder) {
            VariedDataPointColors = variedDataPointColors;
            DrawingOrder = drawingOrder;
        }

        /// <summary>Gets whether data-point colors or marker formatting vary inside a single-series chart group.</summary>
        public bool VariedDataPointColors { get; }

        /// <summary>Gets the chart group's zero-based drawing order relative to other chart groups.</summary>
        public ushort DrawingOrder { get; }
    }
}
