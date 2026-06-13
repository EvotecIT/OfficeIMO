using System;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Produces sequential chart placements for worksheet dashboards that use the default Excel row and column grid.
    /// </summary>
    public sealed class ExcelChartGridLayout {
        private const int DefaultColumnPixels = 64;
        private const int DefaultRowPixels = 20;

        private readonly int _startRow;
        private readonly int _startColumn;
        private readonly int _widthPixels;
        private readonly int _heightPixels;
        private readonly int _chartsPerRow;
        private readonly int _horizontalAdvanceColumns;
        private readonly int _verticalAdvanceRows;
        private int _index;

        /// <summary>
        /// Initializes a chart grid layout.
        /// </summary>
        /// <param name="row">One-based worksheet row for the first chart.</param>
        /// <param name="column">One-based worksheet column for the first chart.</param>
        /// <param name="widthPixels">Chart width in pixels.</param>
        /// <param name="heightPixels">Chart height in pixels.</param>
        /// <param name="chartsPerRow">Number of charts to place before wrapping to the next row.</param>
        /// <param name="horizontalGapPixels">Minimum horizontal gap between chart slots, in pixels, calculated against default-width Excel columns.</param>
        /// <param name="verticalGapRows">Minimum vertical gap between chart rows, in worksheet rows, calculated against default-height Excel rows.</param>
        public ExcelChartGridLayout(
            int row,
            int column,
            int widthPixels = 520,
            int heightPixels = 320,
            int chartsPerRow = 2,
            int horizontalGapPixels = 48,
            int verticalGapRows = 2) {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column));
            if (widthPixels <= 0) throw new ArgumentOutOfRangeException(nameof(widthPixels));
            if (heightPixels <= 0) throw new ArgumentOutOfRangeException(nameof(heightPixels));
            if (chartsPerRow <= 0) throw new ArgumentOutOfRangeException(nameof(chartsPerRow));
            if (horizontalGapPixels < 0) throw new ArgumentOutOfRangeException(nameof(horizontalGapPixels));
            if (verticalGapRows < 0) throw new ArgumentOutOfRangeException(nameof(verticalGapRows));

            _startRow = row;
            _startColumn = column;
            _widthPixels = widthPixels;
            _heightPixels = heightPixels;
            _chartsPerRow = chartsPerRow;
            _horizontalAdvanceColumns = Math.Max(1, CeilDiv(widthPixels + horizontalGapPixels, DefaultColumnPixels));
            _verticalAdvanceRows = Math.Max(1, CeilDiv(heightPixels, DefaultRowPixels) + verticalGapRows);
        }

        /// <summary>
        /// Gets the next chart placement and advances the layout cursor.
        /// </summary>
        /// <returns>Chart placement for the next dashboard slot.</returns>
        public ExcelChartPlacement Next() {
            int rowOffset = _index / _chartsPerRow;
            int columnOffset = _index % _chartsPerRow;
            _index++;

            return new ExcelChartPlacement(
                _startRow + (rowOffset * _verticalAdvanceRows),
                _startColumn + (columnOffset * _horizontalAdvanceColumns),
                _widthPixels,
                _heightPixels);
        }

        private static int CeilDiv(int value, int divisor) {
            return (value + divisor - 1) / divisor;
        }
    }
}
