using System;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes chart series ranges for scatter/bubble charts.
    /// </summary>
    public sealed class ExcelChartSeriesRange {
        /// <summary>
        /// Creates a series range definition for scatter charts.
        /// </summary>
        public ExcelChartSeriesRange(string xRangeA1, string yRangeA1)
            : this(string.Empty, xRangeA1, yRangeA1, null) {
        }

        /// <summary>
        /// Creates a series range definition for scatter/bubble charts.
        /// </summary>
        public ExcelChartSeriesRange(string name, string xRangeA1, string yRangeA1, string? bubbleSizeRangeA1 = null) {
            if (string.IsNullOrWhiteSpace(xRangeA1)) {
                throw new ArgumentException("X range cannot be null or empty.", nameof(xRangeA1));
            }
            if (string.IsNullOrWhiteSpace(yRangeA1)) {
                throw new ArgumentException("Y range cannot be null or empty.", nameof(yRangeA1));
            }
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
        /// Gets the X values range (A1 notation).
        /// </summary>
        public string XRangeA1 { get; }

        /// <summary>
        /// Gets the Y values range (A1 notation).
        /// </summary>
        public string YRangeA1 { get; }

        /// <summary>
        /// Gets the bubble size range (A1 notation), when applicable.
        /// </summary>
        public string? BubbleSizeRangeA1 { get; }
    }
}
