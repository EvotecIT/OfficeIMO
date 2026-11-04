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
        private readonly OfficeLength _width;
        private readonly OfficeLength _height;
        private readonly int _chartsPerRow;
        private readonly int _horizontalAdvanceColumns;
        private readonly int _verticalAdvanceRows;
        private int _index;

        /// <summary>
        /// Initializes a chart grid layout.
        /// </summary>
        /// <param name="row">One-based worksheet row for the first chart.</param>
        /// <param name="column">One-based worksheet column for the first chart.</param>
        /// <param name="width">Chart physical width.</param>
        /// <param name="height">Chart physical height.</param>
        /// <param name="chartsPerRow">Number of charts to place before wrapping to the next row.</param>
        /// <param name="horizontalGap">Minimum horizontal physical gap between chart slots, calculated against default-width Excel columns.</param>
        /// <param name="verticalGapRows">Minimum vertical gap between chart rows, in worksheet rows, calculated against default-height Excel rows.</param>
        public ExcelChartGridLayout(
            int row,
            int column,
            OfficeLength? width = null,
            OfficeLength? height = null,
            int chartsPerRow = 2,
            OfficeLength? horizontalGap = null,
            int verticalGapRows = 2) {
            OfficeLength chartWidth = width ?? OfficeLength.Pixels(520);
            OfficeLength chartHeight = height ?? OfficeLength.Pixels(320);
            OfficeLength gap = horizontalGap ?? OfficeLength.Pixels(48);
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column));
            if (chartWidth <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
            if (chartHeight <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(height));
            if (chartsPerRow <= 0) throw new ArgumentOutOfRangeException(nameof(chartsPerRow));
            if (gap < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(horizontalGap));
            if (verticalGapRows < 0) throw new ArgumentOutOfRangeException(nameof(verticalGapRows));

            _startRow = row;
            _startColumn = column;
            _width = chartWidth;
            _height = chartHeight;
            _chartsPerRow = chartsPerRow;
            _horizontalAdvanceColumns = Math.Max(1, CeilDiv((chartWidth + gap).ToEmus(), OfficeLength.Pixels(DefaultColumnPixels).ToEmus()));
            _verticalAdvanceRows = Math.Max(1, checked(CeilDiv(chartHeight.ToEmus(), OfficeLength.Pixels(DefaultRowPixels).ToEmus()) + verticalGapRows));
        }

        /// <summary>
        /// Gets the next chart placement and advances the layout cursor.
        /// </summary>
        /// <returns>Chart placement for the next dashboard slot.</returns>
        public ExcelChartPlacement Next() {
            int rowOffset = _index / _chartsPerRow;
            int columnOffset = _index % _chartsPerRow;
            var placement = new ExcelChartPlacement(
                checked(_startRow + checked(rowOffset * _verticalAdvanceRows)),
                checked(_startColumn + checked(columnOffset * _horizontalAdvanceColumns)),
                _width,
                _height);
            _index = checked(_index + 1);
            return placement;
        }

        private static int CeilDiv(long value, long divisor) {
            return checked((int)((value / divisor) + (value % divisor == 0 ? 0 : 1)));
        }
    }
}
