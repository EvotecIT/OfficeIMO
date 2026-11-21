namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private bool RequiresDeferredMaterialization
            => HasPendingDirectCellValues || HasUnmaterializedDirectDataSetRows;

        /// <summary>Reads workbook metadata under the read lock without writing deferred worksheet rows first.</summary>
        internal T ExecuteReadWithoutMaterializing<T>(Func<T> read) {
            ArgumentNullException.ThrowIfNull(read);
            ReaderWriterLockSlim lck = EnsureLock();
            if (Locking.IsNoLock || lck.IsWriteLockHeld || lck.IsReadLockHeld || lck.IsUpgradeableReadLockHeld) {
                return read();
            }

            lck.EnterReadLock();
            try {
                return read();
            } finally {
                lck.ExitReadLock();
            }
        }

        internal T ExecuteReadAfterMaterializing<T>(Func<T> read, CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(read);
            cancellationToken.ThrowIfCancellationRequested();
            ReaderWriterLockSlim lck = EnsureLock();
            if (Locking.IsNoLock || lck.IsWriteLockHeld) {
                MaterializeDeferredDataSetImportLocked(cancellationToken, preserveFastSaveModel: true);
                cancellationToken.ThrowIfCancellationRequested();
                return read();
            }

            if (lck.IsReadLockHeld) {
                if (RequiresDeferredMaterialization) {
                    throw new InvalidOperationException(
                        "Deferred worksheet data must be materialized before entering a workbook read scope.");
                }
                cancellationToken.ThrowIfCancellationRequested();
                return read();
            }

            lck.EnterReadLock();
            try {
                if (!RequiresDeferredMaterialization) {
                    cancellationToken.ThrowIfCancellationRequested();
                    return read();
                }
            } finally {
                lck.ExitReadLock();
            }

            lck.EnterUpgradeableReadLock();
            try {
                if (RequiresDeferredMaterialization) {
                    lck.EnterWriteLock();
                    try {
                        MaterializeDeferredDataSetImportLocked(cancellationToken, preserveFastSaveModel: true);
                    } finally {
                        lck.ExitWriteLock();
                    }
                }
                cancellationToken.ThrowIfCancellationRequested();
                return read();
            } finally {
                lck.ExitUpgradeableReadLock();
            }
        }
    }
}
