using System;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private bool RequiresDeferredMaterialization
            => HasPendingDirectCellValues || HasUnmaterializedDirectDataSetRows;

        internal T ExecuteReadAfterMaterializing<T>(Func<T> read) {
            if (read == null) throw new ArgumentNullException(nameof(read));
            ReaderWriterLockSlim lck = EnsureLock();
            if (Locking.IsNoLock || lck.IsWriteLockHeld) {
                MaterializeDeferredDataSetImportLocked(CancellationToken.None);
                return read();
            }

            if (lck.IsReadLockHeld) {
                if (RequiresDeferredMaterialization) {
                    throw new InvalidOperationException(
                        "Deferred worksheet data must be materialized before entering a workbook read scope.");
                }
                return read();
            }

            lck.EnterReadLock();
            try {
                if (!RequiresDeferredMaterialization) return read();
            } finally {
                lck.ExitReadLock();
            }

            lck.EnterUpgradeableReadLock();
            try {
                if (RequiresDeferredMaterialization) {
                    lck.EnterWriteLock();
                    try {
                        MaterializeDeferredDataSetImportLocked(CancellationToken.None);
                    } finally {
                        lck.ExitWriteLock();
                    }
                }
                return read();
            } finally {
                lck.ExitUpgradeableReadLock();
            }
        }
    }
}
