namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Creates a sequential chart layout helper for dashboard-style worksheets that use the default Excel row and column grid.
        /// </summary>
        /// <param name="row">One-based worksheet row for the first chart.</param>
        /// <param name="column">One-based worksheet column for the first chart.</param>
        /// <param name="width">Default chart physical width.</param>
        /// <param name="height">Default chart physical height.</param>
        /// <param name="chartsPerRow">Number of charts to place before wrapping to the next row.</param>
        /// <param name="horizontalGap">Minimum horizontal physical gap between chart slots, calculated against default-width Excel columns.</param>
        /// <param name="verticalGapRows">Minimum vertical gap between chart rows, in worksheet rows, calculated against default-height Excel rows.</param>
        /// <returns>A layout helper that produces chart placement slots.</returns>
        public ExcelChartGridLayout ChartLayout(
            int row,
            int column,
            OfficeLength? width = null,
            OfficeLength? height = null,
            int chartsPerRow = 2,
            OfficeLength? horizontalGap = null,
            int verticalGapRows = 2) {
            width ??= OfficeLength.Pixels(520);
            height ??= OfficeLength.Pixels(320);
            horizontalGap ??= OfficeLength.Pixels(48);
            return new ExcelChartGridLayout(row, column, width.Value, height.Value, chartsPerRow, horizontalGap.Value, verticalGapRows);
        }
    }
}
