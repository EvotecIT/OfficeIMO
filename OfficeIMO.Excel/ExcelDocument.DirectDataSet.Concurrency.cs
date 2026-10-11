using System;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private bool RequiresDeferredMaterialization
            => HasPendingDirectCellValues || HasUnmaterializedDirectDataSetRows;

        internal T ExecuteReadAfterMaterializing<T>(Func<T> read, CancellationToken cancellationToken = default) {
            if (read == null) throw new ArgumentNullException(nameof(read));
            cancellationToken.ThrowIfCancellationRequested();
            ReaderWriterLockSlim lck = EnsureLock();
            if (Locking.IsNoLock || lck.IsWriteLockHeld) {
                MaterializeDeferredDataSetImportLocked(cancellationToken);
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
                        MaterializeDeferredDataSetImportLocked(cancellationToken);
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
