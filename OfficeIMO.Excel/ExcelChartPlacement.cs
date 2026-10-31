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
        /// <param name="width">Chart physical width.</param>
        /// <param name="height">Chart physical height.</param>
        public ExcelChartPlacement(int row, int column, OfficeLength width, OfficeLength height) {
            if (row <= 0 || row > ExcelA1.MaxRows) throw new System.ArgumentOutOfRangeException(nameof(row));
            if (column <= 0 || column > ExcelA1.MaxColumns) throw new System.ArgumentOutOfRangeException(nameof(column));
            if (width <= OfficeLength.Zero) throw new System.ArgumentOutOfRangeException(nameof(width));
            if (height <= OfficeLength.Zero) throw new System.ArgumentOutOfRangeException(nameof(height));
            Row = row;
            Column = column;
            Width = width;
            Height = height;
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
        /// Gets the chart physical width.
        /// </summary>
        public OfficeLength Width { get; }

        /// <summary>
        /// Gets the chart physical height.
        /// </summary>
        public OfficeLength Height { get; }
    }
}
