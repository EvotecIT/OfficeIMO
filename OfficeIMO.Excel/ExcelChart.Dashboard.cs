using System;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        /// <summary>Applies the dashboard preset's style, legend and data labels to this chart.</summary>
        /// <remarks>The chart's existing data, family and placement are retained. Explicit style overrides
        /// use the same values as <see cref="ExcelSheet.AddDashboardChart"/>.</remarks>
        /// <param name="preset">The dashboard appearance preset.</param>
        /// <param name="styleId">Optional chart style override.</param>
        /// <param name="colorStyleId">Optional chart color-style override.</param>
        /// <returns>This native chart for further editing.</returns>
        public ExcelChart ApplyDashboardPreset(ExcelDashboardChartPreset preset = ExcelDashboardChartPreset.Comparison,
            int? styleId = null, int? colorStyleId = null) {
            if (!Enum.IsDefined(typeof(ExcelDashboardChartPreset), preset))
                throw new ArgumentOutOfRangeException(nameof(preset));
            ExcelSheet.ApplyDashboardAppearance(this, preset, styleId, colorStyleId);
            return this;
        }
    }
}
