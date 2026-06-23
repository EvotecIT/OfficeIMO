namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents the used-range bounds declared by a legacy BIFF worksheet DIMENSIONS record.
    /// </summary>
    public sealed class LegacyXlsWorksheetDimension {
        private LegacyXlsWorksheetDimension(int? firstRow, int? firstColumn, int? lastRow, int? lastColumn) {
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
        }

        /// <summary>
        /// Gets a dimension instance representing an explicitly empty worksheet.
        /// </summary>
        public static LegacyXlsWorksheetDimension Empty { get; } = new LegacyXlsWorksheetDimension(null, null, null, null);

        /// <summary>
        /// Gets whether the DIMENSIONS record declared no used cells.
        /// </summary>
        public bool IsEmpty => !FirstRow.HasValue;

        /// <summary>
        /// Gets the one-based first used row, when the worksheet is not empty.
        /// </summary>
        public int? FirstRow { get; }

        /// <summary>
        /// Gets the one-based first used column, when the worksheet is not empty.
        /// </summary>
        public int? FirstColumn { get; }

        /// <summary>
        /// Gets the one-based last used row, when the worksheet is not empty.
        /// </summary>
        public int? LastRow { get; }

        /// <summary>
        /// Gets the one-based last used column, when the worksheet is not empty.
        /// </summary>
        public int? LastColumn { get; }

        /// <summary>
        /// Gets the declared range in A1 notation. Empty sheets report <c>A1:A1</c>.
        /// </summary>
        public string UsedRangeA1 {
            get {
                if (IsEmpty) {
                    return "A1:A1";
                }

                string start = A1.CellReference(FirstRow!.Value, FirstColumn!.Value);
                string end = A1.CellReference(LastRow!.Value, LastColumn!.Value);
                return start == end ? start + ":" + start : start + ":" + end;
            }
        }

        internal static LegacyXlsWorksheetDimension FromOneBasedBounds(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            if (firstRow <= 0) throw new ArgumentOutOfRangeException(nameof(firstRow));
            if (firstColumn <= 0) throw new ArgumentOutOfRangeException(nameof(firstColumn));
            if (lastRow < firstRow) throw new ArgumentOutOfRangeException(nameof(lastRow));
            if (lastColumn < firstColumn) throw new ArgumentOutOfRangeException(nameof(lastColumn));

            return new LegacyXlsWorksheetDimension(firstRow, firstColumn, lastRow, lastColumn);
        }
    }
}
