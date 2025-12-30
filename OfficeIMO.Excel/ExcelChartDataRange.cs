using System;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents the workbook range used for chart data (categories + series values).
    /// </summary>
    public sealed class ExcelChartDataRange {
        /// <summary>
        /// Creates a chart data range descriptor for categories and series values.
        /// </summary>
        public ExcelChartDataRange(string sheetName, int startRow, int startColumn, int categoryCount, int seriesCount, bool hasHeaderRow = true) {
            if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentNullException(nameof(sheetName));
            if (startRow <= 0 || startColumn <= 0) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (categoryCount <= 0) throw new ArgumentOutOfRangeException(nameof(categoryCount));
            if (seriesCount <= 0) throw new ArgumentOutOfRangeException(nameof(seriesCount));

            SheetName = sheetName;
            StartRow = startRow;
            StartColumn = startColumn;
            CategoryCount = categoryCount;
            SeriesCount = seriesCount;
            HasHeaderRow = hasHeaderRow;
        }

        /// <summary>
        /// Gets the worksheet name that owns the data range.
        /// </summary>
        public string SheetName { get; }

        /// <summary>
        /// Gets the 1-based starting row of the data range.
        /// </summary>
        public int StartRow { get; }

        /// <summary>
        /// Gets the 1-based starting column of the data range.
        /// </summary>
        public int StartColumn { get; }

        /// <summary>
        /// Gets the number of category rows.
        /// </summary>
        public int CategoryCount { get; }

        /// <summary>
        /// Gets the number of series columns.
        /// </summary>
        public int SeriesCount { get; }

        /// <summary>
        /// Gets whether the data range includes a header row.
        /// </summary>
        public bool HasHeaderRow { get; }

        /// <summary>
        /// Gets the first row index that contains category values.
        /// </summary>
        public int CategoryStartRow => StartRow + (HasHeaderRow ? 1 : 0);

        /// <summary>
        /// Gets the last row index that contains category values.
        /// </summary>
        public int CategoryEndRow => CategoryStartRow + CategoryCount - 1;

        /// <summary>
        /// Gets the first column index that contains series values.
        /// </summary>
        public int SeriesStartColumn => StartColumn + 1;

        /// <summary>
        /// Gets the last column index that contains series values.
        /// </summary>
        public int SeriesEndColumn => SeriesStartColumn + SeriesCount - 1;

        /// <summary>
        /// Returns a new range with updated category and series counts.
        /// </summary>
        public ExcelChartDataRange WithSize(int categoryCount, int seriesCount) {
            return new ExcelChartDataRange(SheetName, StartRow, StartColumn, categoryCount, seriesCount, HasHeaderRow);
        }

        /// <summary>
        /// Gets the data range (headers + categories + values) as an A1 reference.
        /// </summary>
        public string DataRangeA1 {
            get {
                int endRow = CategoryEndRow;
                int endCol = SeriesEndColumn;
                return ExcelChartUtils.BuildRangeA1(StartRow, StartColumn, endRow, endCol);
            }
        }

        /// <summary>
        /// Gets the categories range as an A1 reference.
        /// </summary>
        public string CategoriesRangeA1 {
            get {
                return ExcelChartUtils.BuildRangeA1(CategoryStartRow, StartColumn, CategoryEndRow, StartColumn);
            }
        }

        /// <summary>
        /// Gets the header cell for the specified series (A1), or empty when no header row is used.
        /// </summary>
        public string SeriesNameCellA1(int seriesIndex) {
            if (!HasHeaderRow) return string.Empty;
            int col = SeriesStartColumn + seriesIndex;
            return ExcelChartUtils.BuildCellA1(StartRow, col);
        }

        /// <summary>
        /// Gets the series values range as an A1 reference for the specified series.
        /// </summary>
        public string SeriesValuesRangeA1(int seriesIndex) {
            int col = SeriesStartColumn + seriesIndex;
            return ExcelChartUtils.BuildRangeA1(CategoryStartRow, col, CategoryEndRow, col);
        }
    }
}
