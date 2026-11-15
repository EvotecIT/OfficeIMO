namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes the worksheet ranges of one chart series: categories (X) and values (Y) for category and scatter charts, plus bubble sizes for bubble charts.
    /// </summary>
    public sealed class ExcelChartSeriesRange {
        /// <summary>
        /// Creates a series range definition from a category or X range and a values or Y range.
        /// </summary>
        public ExcelChartSeriesRange(string xRangeA1, string yRangeA1)
            : this(string.Empty, xRangeA1, yRangeA1, null) {
        }

        /// <summary>
        /// Creates a named series range definition; bubble charts also take a bubble size range.
        /// </summary>
        public ExcelChartSeriesRange(string name, string xRangeA1, string yRangeA1, string? bubbleSizeRangeA1 = null) {
            ArgumentException.ThrowIfNullOrWhiteSpace(xRangeA1);
            ArgumentException.ThrowIfNullOrWhiteSpace(yRangeA1);
            if (bubbleSizeRangeA1 != null && string.IsNullOrWhiteSpace(bubbleSizeRangeA1)) {
                throw new ArgumentException("Bubble size range cannot be empty.", nameof(bubbleSizeRangeA1));
            }

            Name = name ?? string.Empty;
            XRangeA1 = xRangeA1;
            YRangeA1 = yRangeA1;
            BubbleSizeRangeA1 = bubbleSizeRangeA1;
        }

        /// <summary>
        /// Gets the series display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the categories range of a category chart, or the X values range of a scatter or bubble chart (A1 notation).
        /// </summary>
        public string XRangeA1 { get; }

        /// <summary>
        /// Gets the values range of a category chart, or the Y values range of a scatter or bubble chart (A1 notation).
        /// </summary>
        public string YRangeA1 { get; }

        /// <summary>
        /// Gets the bubble size range (A1 notation), when applicable.
        /// </summary>
        public string? BubbleSizeRangeA1 { get; }
    }
}
