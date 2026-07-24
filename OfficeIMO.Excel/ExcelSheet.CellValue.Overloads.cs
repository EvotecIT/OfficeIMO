using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, string value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellStringValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellStringValueCore(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, double value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellDoubleValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellDoubleValueCore(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, float value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellDoubleValueCore(row, column, (double)value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellDoubleValueCore(row, column, (double)value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, decimal value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellDecimalValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellDecimalValueCore(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, int value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, long value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, short value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTime value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellDateTimeValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellDateTimeValueCore(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTimeOffset value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellDateTimeOffsetValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellDateTimeOffsetValueCore(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
#if NET6_0_OR_GREATER
        public void CellValue(int row, int column, DateOnly value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellDateOnlyValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellDateOnlyValueCore(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, TimeOnly value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellTimeOnlyValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellTimeOnlyValueCore(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
#endif
        public void CellValue(int row, int column, TimeSpan value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellTimeSpanValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellTimeSpanValueCore(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, uint value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, ulong value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, ushort value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, byte value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, sbyte value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellNumberTextValueCore(row, column, InvariantNumberText.Get(value));
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, bool value) {
            if (TrySetPendingDirectCellValue(row, column, value)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellBooleanValueCore(row, column, value);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();
            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellBooleanValueCore(row, column, value);
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <summary>
        /// Sets a formula in the specified cell while preserving any existing value as its cached result.
        /// </summary>
        /// <remarks>
        /// Cached formula values are part of the workbook contract and are intentionally retained. This method is
        /// not a secure-erasure or redaction API; callers removing sensitive historic content should rebuild the
        /// workbook from approved data rather than relying on an in-place formula mutation.
        /// </remarks>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="formula">The formula expression.</param>
        public void CellFormula(int row, int column, string formula) {
            if (TrySetPendingDirectCellFormula(row, column, formula)) {
                return;
            }

            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                CellFormulaCore(row, column, formula);
                return;
            }

            MaterializeDeferredDataSetImportIfNeeded();

            var lck = _excelDocument.EnsureLock();
            lck.EnterWriteLock();
            try {
                CellFormulaCore(row, column, formula);
            } finally {
                lck.ExitWriteLock();
            }
        }
    }
}
