using System;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Batch ownership is thread-affine because ReaderWriterLockSlim write ownership is
        // thread-affine. A workbook mutation on another thread must never inherit this fast path.
        private int _batchOwnerThreadId;
        private bool _isBatchOperation {
            get => Volatile.Read(ref _batchOwnerThreadId) == Environment.CurrentManagedThreadId;
            set => Volatile.Write(ref _batchOwnerThreadId, value ? Environment.CurrentManagedThreadId : 0);
        }

        private bool _batchHasCellMutations;

        internal NoLockContext BeginNoLock() => new NoLockContext();

        /// <summary>
        /// Executes multiple worksheet mutations under a single workbook write lock. Cell styling inside the batch
        /// (bold, fills, fonts, wrapping, number formats and similar) looks formats up through an index and saves the
        /// stylesheet once when the batch ends, so styling cells row by row stays proportional to the cells styled.
        /// </summary>
        /// <param name="action">The worksheet updates to execute.</param>
        public void Batch(Action<ExcelSheet> action) {
            if (action == null) throw new ArgumentNullException(nameof(action));

            if (Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                try {
                    using (TryBeginStyleBatch()) {
                        action(this);
                    }
                } finally {
                    CommitCommentVmlChanges();
                }
                return;
            }

            ReaderWriterLockSlim lck = _excelDocument.EnsureLock();
            if (_isBatchOperation && lck.IsWriteLockHeld) {
                MaterializeDeferredDataSetImportIfNeeded();
                action(this);
                return;
            }

            Locking.ExecuteWrite(lck, () => {
                bool wasBatchOperation = _isBatchOperation;
                bool hadBatchCellMutations = _batchHasCellMutations;
                try {
                    MaterializeDeferredDataSetImportIfNeeded();
                    _isBatchOperation = true;
                    _batchHasCellMutations = false;
                    using (TryBeginStyleBatch()) {
                        action(this);
                    }
                } finally {
                    // Batch is not a transaction: edits made before a callback throws still need to reach Save.
                    try {
                        CommitCommentVmlChanges();
                        if (_batchHasCellMutations) _excelDocument.MarkPackageDirty();
                    } finally {
                        _isBatchOperation = wasBatchOperation;
                        _batchHasCellMutations = hadBatchCellMutations;
                    }
                }
            });
        }

        internal sealed class NoLockContext : IDisposable {
            private readonly IDisposable _scope;

            internal NoLockContext() => _scope = Locking.EnterNoLockScope();

            public void Dispose() => _scope.Dispose();
        }

        private void WriteLock(Action action) {
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                action();
                MarkRequiresSavePreparation();
            });
        }

        private void WriteLockWorksheetPreparationOnly(Action action) {
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                action();
                MarkRequiresWorksheetPreparation();
            });
        }

        private void WriteLockWorksheetPreparationOnly(Func<bool> action) {
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                if (action()) MarkRequiresWorksheetPreparation();
            });
        }

        /// <summary>Joins a workbook write lock held by any sheet and honors an explicit no-lock scope.</summary>
        private bool HasCellValueWriteAccess => _isBatchOperation || Locking.IsNoLock
            || _excelDocument.EnsureLock().IsWriteLockHeld;

        private void WriteLockConditional(Action action) {
            if (_isBatchOperation || Locking.IsNoLock) {
                MaterializeDeferredDataSetImportIfNeeded();
                action();
                MarkRequiresSavePreparation();
                return;
            }

            WriteLock(() => {
                MaterializeDeferredDataSetImportIfNeeded();
                action();
            });
        }

        private void MaterializeDeferredDataSetImportIfNeeded() {
            if (_excelDocument.IsPreservingDirectDataSetExternalCellMutation
                && _excelDocument.HasDeferredDirectDataSetImport
                && !_excelDocument.HasPendingDirectCellValues) {
                return;
            }

            if (_excelDocument.HasUnmaterializedDirectDataSetRows
                || _excelDocument.HasPendingDirectCellValues) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }
        }
    }
}
