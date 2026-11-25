namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Writes an optional header row and row-major selector results on an empty sheet straight into the pending
        /// cell buffer, the route <see cref="TryBufferFreshSheetCellValues"/> takes for boxed cells. The values stay
        /// typed until save, and the per-cell admission work runs once for the block instead of once per cell.
        /// Missing values stay blank, and date and time values (offset and date-only values included) take the tabular insertion format, as they do when the
        /// block is the workbook's only sheet. Returns false, with nothing written, when the sheet cannot take the block.
        /// </summary>
        internal bool TryBufferFreshSheetTabularValues(IReadOnlyList<string>? headers, ExcelTabularCellValues values, int startRow, CancellationToken ct) {
            int columnCount = values.ColumnCount;
            int headerRows = headers != null ? 1 : 0;
            if (columnCount <= 0 || startRow <= 0
                || (long)startRow + values.RowCount + headerRows - 1 > ExcelA1.MaxRows
                || columnCount > ExcelA1.MaxColumns
                || (long)(values.RowCount + headerRows) * columnCount < PendingDirectCellValueMinimumCellCount) {
                return false;
            }

            if (!_excelDocument.IsMaterializingDeferredDataSetImport && _excelDocument.HasUnmaterializedDirectDataSetRows) {
                _excelDocument.MaterializeDeferredDataSetImportKeepingCellBuffers(ct);
            }

            if (HasCellValueWriteAccess) {
                return TryBufferFreshSheetTabularValuesLocked(headers, values, startRow, ct);
            }

            return Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => TryBufferFreshSheetTabularValuesLocked(headers, values, startRow, ct));
        }

        private bool TryBufferFreshSheetTabularValuesLocked(IReadOnlyList<string>? headers, ExcelTabularCellValues values, int startRow, CancellationToken ct) {
            if (_pendingCellValueDirectSaveBuffer != null || !CanAcceptPendingDirectCell()) return false;
            EnsureDynamicArrayWriteIndex();
            if (SpillOwnership.WriteRanges.Count != 0) return false;

            int columnCount = values.ColumnCount;
            if (headers != null) {
                for (int column = 0; column < columnCount; column++) {
                    if (!CanBufferDirectCell(ExcelBufferedCell.Text(headers[column]))) return false;
                }
            }

            int cellCount = values.ColumnCount * values.RowCount;
            for (int index = 0; index < cellCount; index++) {
                ExcelBufferedCell cell = GetBufferedCell(values, index);
                if (cell.Kind == ExcelBufferedCellKind.Blank || cell.Reference is DateTimeOffset) continue;
                if (!CanBufferDirectCell(cell)) return false;
            }

            uint dateStyle = 0U;
            for (int index = 0; index < cellCount; index++) {
                ExcelBufferedCellKind kind = values.GetKind(index);
                if (kind == ExcelBufferedCellKind.DateTime
                    || (kind == ExcelBufferedCellKind.Value && IsTabularDateValue(values.GetReference(index)))) {
                    dateStyle = GetOrCreateDirectTabularDateStyle(useCellValueNumberFormats: false);
                    break;
                }
            }

            int row = startRow;
            if (headers != null) {
                for (int column = 0; column < columnCount; column++) {
                    if (!TryAppendFreshSheetCell(row, column + 1, ExcelBufferedCell.Text(headers[column]), dateStyle, ct)) return false;
                }
                row++;
            }

            int valueIndex = 0;
            for (int r = 0; r < values.RowCount; r++, row++) {
                ct.ThrowIfCancellationRequested();
                for (int column = 1; column <= columnCount; column++, valueIndex++) {
                    ExcelBufferedCell cell = GetBufferedCell(values, valueIndex);
                    if (cell.Kind == ExcelBufferedCellKind.Blank) continue;
                    if (!TryAppendFreshSheetCell(row, column, cell, dateStyle, ct)) return false;
                }
            }

            ClearHeaderCache();
            if (_pendingCellValueDirectSaveBuffer != null) {
                _excelDocument.MarkFormulaInputMutation();
                if (!_excelDocument.IsPackageDirty) _excelDocument.MarkPackageDirty();
            }
            return true;
        }

        /// <summary>
        /// Appends one cell. The cells the worksheet mirrors, and values that need conversion first, take the
        /// single-cell route; later plain cells go straight into the buffer that route has created.
        /// </summary>
        private bool TryAppendFreshSheetCell(int row, int column, in ExcelBufferedCell cell, uint dateStyle, CancellationToken ct) {
            ExcelCellValueBuffer? buffer = _pendingCellValueDirectSaveBuffer;
            // Offset and date-only values arrive as values and are converted to date cells by the single-cell route.
            bool date = dateStyle != 0U && (cell.Kind == ExcelBufferedCellKind.DateTime || IsTabularDateValue(cell.Reference));
            if (buffer != null
                && buffer.CellCount >= PendingDirectCellValueMinimumCellCount
                && cell.Kind != ExcelBufferedCellKind.Value
                && buffer.CanAdd(row, column)) {
                buffer.Add(row, column, cell);
                if (date) {
                    buffer.SetStyle(buffer.CellCount - 1, dateStyle);
                } else if (cell.HasLineBreak) {
                    SetBufferedAutomaticStyle(buffer, buffer.CellCount - 1, 0U, cell);
                }
                return true;
            }

            if (TrySetPendingDirectCellLocked(row, column, cell)) {
                if (date && _pendingCellValueDirectSaveBuffer is { } created) {
                    int cellIndex = created.CellCount - 1;
                    created.SetStyle(cellIndex, dateStyle);
                    if (IsMirroredBufferedCell(cellIndex)) GetCell(row, column).StyleIndex = dateStyle;
                }
                return true;
            }
            // Cells already written are rewritten with the same values by the DOM route.
            MaterializePendingDirectCellValues(ct);
            return false;
        }

        /// <summary>Values other than <see cref="DateTime"/> that tabular writes format as dates.</summary>
        private static bool IsTabularDateValue(object? value) =>
            value is DateTimeOffset
#if NET6_0_OR_GREATER
            || value is DateOnly
#endif
            ;

        private static ExcelBufferedCell GetBufferedCell(ExcelTabularCellValues values, int index) {
            int offset = values.GetChunk(index, out byte[] kinds, out long[] bits);
            var kind = (ExcelBufferedCellKind)kinds[offset];
            switch (kind) {
                case ExcelBufferedCellKind.Blank:
                    return default;
                case ExcelBufferedCellKind.Text:
                    return ExcelBufferedCell.Text((string)values.GetReferenceById(bits[offset]));
                case ExcelBufferedCellKind.Value:
                    return ExcelBufferedCell.FromValue(values.GetReferenceById(bits[offset]));
                default:
                    return ExcelBufferedCell.Restore(kind, bits[offset], null);
            }
        }
    }
}
