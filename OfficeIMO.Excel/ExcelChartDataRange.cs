using System;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes how categories and series values are arranged in a worksheet chart data range.
    /// </summary>
    public enum ExcelChartDataOrientation {
        /// <summary>
        /// Categories are stored in the first column and each series is stored in a following column.
        /// </summary>
        Vertical,

        /// <summary>
        /// Categories are stored in the first row and each series is stored in a following row.
        /// </summary>
        Horizontal
    }

    /// <summary>
    /// Represents the workbook range used for chart data (categories + series values).
    /// </summary>
    public sealed class ExcelChartDataRange {
        private const int MaxExcelRows = 1_048_576;
        private const int MaxExcelColumns = 16_384;

        /// <summary>
        /// Creates a chart data range descriptor for categories and series values.
        /// </summary>
        public ExcelChartDataRange(string sheetName, int startRow, int startColumn, int categoryCount, int seriesCount, bool hasHeaderRow = true)
            : this(sheetName, startRow, startColumn, categoryCount, seriesCount, hasHeaderRow, ExcelChartDataOrientation.Vertical) {
        }

        /// <summary>
        /// Creates a chart data range descriptor for categories and series values.
        /// </summary>
        public ExcelChartDataRange(string sheetName, int startRow, int startColumn, int categoryCount, int seriesCount, bool hasHeaderRow, ExcelChartDataOrientation orientation) {
            if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentNullException(nameof(sheetName));
            if (startRow <= 0 || startColumn <= 0) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (categoryCount <= 0) throw new ArgumentOutOfRangeException(nameof(categoryCount));
            if (seriesCount <= 0) throw new ArgumentOutOfRangeException(nameof(seriesCount));
            if (startRow > MaxExcelRows) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn >= MaxExcelColumns) throw new ArgumentOutOfRangeException(nameof(startColumn));
            if (!Enum.IsDefined(typeof(ExcelChartDataOrientation), orientation)) throw new ArgumentOutOfRangeException(nameof(orientation));

            if (orientation == ExcelChartDataOrientation.Vertical) {
                int categoryStartRow = startRow + (hasHeaderRow ? 1 : 0);
                int categoryEndRow = categoryStartRow + categoryCount - 1;
                int seriesStartColumn = startColumn + 1;
                int seriesEndColumn = seriesStartColumn + seriesCount - 1;
                if (categoryEndRow > MaxExcelRows) throw new ArgumentOutOfRangeException(nameof(categoryCount));
                if (seriesEndColumn > MaxExcelColumns) throw new ArgumentOutOfRangeException(nameof(seriesCount));
            } else {
                int categoryStartColumn = startColumn + (hasHeaderRow ? 1 : 0);
                int categoryEndColumn = categoryStartColumn + categoryCount - 1;
                int seriesStartRow = startRow + 1;
                int seriesEndRow = seriesStartRow + seriesCount - 1;
                if (categoryEndColumn > MaxExcelColumns) throw new ArgumentOutOfRangeException(nameof(categoryCount));
                if (seriesEndRow > MaxExcelRows) throw new ArgumentOutOfRangeException(nameof(seriesCount));
            }

            SheetName = sheetName;
            StartRow = startRow;
            StartColumn = startColumn;
            CategoryCount = categoryCount;
            SeriesCount = seriesCount;
            HasHeaderRow = hasHeaderRow;
            Orientation = orientation;
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
        /// Gets whether the data range includes series-name headers.
        /// </summary>
        public bool HasHeaderRow { get; }

        /// <summary>
        /// Gets the orientation of the chart data range.
        /// </summary>
        public ExcelChartDataOrientation Orientation { get; }

        /// <summary>
        /// Gets the first row index that contains category values.
        /// </summary>
        public int CategoryStartRow => Orientation == ExcelChartDataOrientation.Vertical ? StartRow + (HasHeaderRow ? 1 : 0) : StartRow;

        /// <summary>
        /// Gets the last row index that contains category values.
        /// </summary>
        public int CategoryEndRow => Orientation == ExcelChartDataOrientation.Vertical ? CategoryStartRow + CategoryCount - 1 : StartRow;

        /// <summary>
        /// Gets the first column index that contains category values.
        /// </summary>
        public int CategoryStartColumn => Orientation == ExcelChartDataOrientation.Vertical ? StartColumn : StartColumn + (HasHeaderRow ? 1 : 0);

        /// <summary>
        /// Gets the last column index that contains category values.
        /// </summary>
        public int CategoryEndColumn => Orientation == ExcelChartDataOrientation.Vertical ? StartColumn : CategoryStartColumn + CategoryCount - 1;

        /// <summary>
        /// Gets the first row index that contains series values.
        /// </summary>
        public int SeriesStartRow => Orientation == ExcelChartDataOrientation.Vertical ? CategoryStartRow : StartRow + 1;

        /// <summary>
        /// Gets the last row index that contains series values.
        /// </summary>
        public int SeriesEndRow => Orientation == ExcelChartDataOrientation.Vertical ? CategoryEndRow : SeriesStartRow + SeriesCount - 1;

        /// <summary>
        /// Gets the first column index that contains series values.
        /// </summary>
        public int SeriesStartColumn => Orientation == ExcelChartDataOrientation.Vertical ? StartColumn + 1 : CategoryStartColumn;

        /// <summary>
        /// Gets the last column index that contains series values.
        /// </summary>
        public int SeriesEndColumn => Orientation == ExcelChartDataOrientation.Vertical ? SeriesStartColumn + SeriesCount - 1 : CategoryEndColumn;

        /// <summary>
        /// Returns a new range with updated category and series counts.
        /// </summary>
        public ExcelChartDataRange WithSize(int categoryCount, int seriesCount) {
            return new ExcelChartDataRange(SheetName, StartRow, StartColumn, categoryCount, seriesCount, HasHeaderRow, Orientation);
        }

        /// <summary>
        /// Gets the data range (headers + categories + values) as an A1 reference.
        /// </summary>
        public string DataRangeA1 {
            get {
                int endRow = Orientation == ExcelChartDataOrientation.Vertical ? CategoryEndRow : SeriesEndRow;
                int endCol = Orientation == ExcelChartDataOrientation.Vertical ? SeriesEndColumn : CategoryEndColumn;
                return ExcelChartUtils.BuildRangeA1(StartRow, StartColumn, endRow, endCol);
            }
        }

        /// <summary>
        /// Gets the categories range as an A1 reference.
        /// </summary>
        public string CategoriesRangeA1 {
            get {
                return ExcelChartUtils.BuildRangeA1(CategoryStartRow, CategoryStartColumn, CategoryEndRow, CategoryEndColumn);
            }
        }

        /// <summary>
        /// Gets the header cell for the specified series (A1), or empty when no series-name header is used.
        /// </summary>
        public string SeriesNameCellA1(int seriesIndex) {
            if (!HasHeaderRow) return string.Empty;
            if (Orientation == ExcelChartDataOrientation.Vertical) {
                int col = SeriesStartColumn + seriesIndex;
                return ExcelChartUtils.BuildCellA1(StartRow, col);
            } else {
                int row = SeriesStartRow + seriesIndex;
                return ExcelChartUtils.BuildCellA1(row, StartColumn);
            }
        }

        /// <summary>
        /// Gets the series values range as an A1 reference for the specified series.
        /// </summary>
        public string SeriesValuesRangeA1(int seriesIndex) {
            if (Orientation == ExcelChartDataOrientation.Vertical) {
                int col = SeriesStartColumn + seriesIndex;
                return ExcelChartUtils.BuildRangeA1(CategoryStartRow, col, CategoryEndRow, col);
            } else {
                int row = SeriesStartRow + seriesIndex;
                return ExcelChartUtils.BuildRangeA1(row, CategoryStartColumn, row, CategoryEndColumn);
            }
        }
    }
}
