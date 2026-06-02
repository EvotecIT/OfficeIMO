using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        /// <summary>
        /// Sets the specified value into a cell, inferring the data type.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">The value to assign.</param>
        public void CellValue(int row, int column, object? value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                CellValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellValueCoreNoMaterialize(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <summary>
        /// Sets the value of a cell using a nullable struct.
        /// </summary>
        /// <typeparam name="T">The value type.</typeparam>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">The nullable value to assign.</param>
        public void CellValue<T>(int row, int column, T? value) where T : struct {
            if (TrySetPendingDirectCellValue(row, column, value.HasValue ? value.Value : null)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellValueCore(row, column, value.HasValue ? value.Value : null);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellValueCoreNoMaterialize(row, column, value.HasValue ? value.Value : null);
            } finally {
                lck.ExitWriteLock();
            }
        }
    }
}
