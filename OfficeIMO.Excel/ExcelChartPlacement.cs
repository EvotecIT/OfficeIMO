namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart placement slot in an Excel worksheet.
    /// </summary>
    public sealed class ExcelChartPlacement {
        /// <summary>
        /// Initializes a new chart placement.
        /// </summary>
        /// <param name="row">One-based worksheet row where the chart should be anchored.</param>
        /// <param name="column">One-based worksheet column where the chart should be anchored.</param>
        /// <param name="widthPixels">Chart width in pixels.</param>
        /// <param name="heightPixels">Chart height in pixels.</param>
        public ExcelChartPlacement(int row, int column, int widthPixels, int heightPixels) {
            Row = row;
            Column = column;
            WidthPixels = widthPixels;
            HeightPixels = heightPixels;
        }

        /// <summary>
        /// Gets the one-based worksheet row where the chart should be anchored.
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// Gets the one-based worksheet column where the chart should be anchored.
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// Gets the chart width in pixels.
        /// </summary>
        public int WidthPixels { get; }

        /// <summary>
        /// Gets the chart height in pixels.
        /// </summary>
        public int HeightPixels { get; }
    }
}
