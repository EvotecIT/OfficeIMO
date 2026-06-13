namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Creates a sequential chart layout helper for dashboard-style worksheets that use the default Excel row and column grid.
        /// </summary>
        /// <param name="row">One-based worksheet row for the first chart.</param>
        /// <param name="column">One-based worksheet column for the first chart.</param>
        /// <param name="widthPixels">Default chart width in pixels.</param>
        /// <param name="heightPixels">Default chart height in pixels.</param>
        /// <param name="chartsPerRow">Number of charts to place before wrapping to the next row.</param>
        /// <param name="horizontalGapPixels">Minimum horizontal gap between chart slots, in pixels, calculated against default-width Excel columns.</param>
        /// <param name="verticalGapRows">Minimum vertical gap between chart rows, in worksheet rows, calculated against default-height Excel rows.</param>
        /// <returns>A layout helper that produces chart placement slots.</returns>
        public ExcelChartGridLayout ChartLayout(
            int row,
            int column,
            int widthPixels = 520,
            int heightPixels = 320,
            int chartsPerRow = 2,
            int horizontalGapPixels = 48,
            int verticalGapRows = 2) {
            return new ExcelChartGridLayout(row, column, widthPixels, heightPixels, chartsPerRow, horizontalGapPixels, verticalGapRows);
        }
    }
}
