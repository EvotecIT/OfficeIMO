namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents row-block lookup metadata decoded from a BIFF Index record.
    /// </summary>
    public sealed class LegacyXlsWorksheetIndex {
        /// <summary>
        /// Creates row-block lookup metadata.
        /// </summary>
        /// <param name="firstRowIndex">One-based first row covered by the indexed row block.</param>
        /// <param name="rowAfterLastIndex">One-based row after the last covered row, matching the BIFF record semantics.</param>
        /// <param name="reservedRecordOffset">Reserved legacy stream offset field retained for diagnostics and corpus comparison.</param>
        /// <param name="dbCellBlockCount">Number of DBCell block offsets stored by the record.</param>
        public LegacyXlsWorksheetIndex(int firstRowIndex, int rowAfterLastIndex, uint reservedRecordOffset, int dbCellBlockCount) {
            FirstRowIndex = firstRowIndex;
            RowAfterLastIndex = rowAfterLastIndex;
            ReservedRecordOffset = reservedRecordOffset;
            DbCellBlockCount = dbCellBlockCount;
        }

        /// <summary>Gets the one-based first row covered by the indexed row block.</summary>
        public int FirstRowIndex { get; }

        /// <summary>Gets the one-based row after the last covered row, matching the BIFF record semantics.</summary>
        public int RowAfterLastIndex { get; }

        /// <summary>Gets the reserved legacy stream offset field retained for diagnostics and corpus comparison.</summary>
        public uint ReservedRecordOffset { get; }

        /// <summary>Gets the number of DBCell block offsets stored by the record.</summary>
        public int DbCellBlockCount { get; }
    }
}
