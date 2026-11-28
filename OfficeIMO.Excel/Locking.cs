namespace OfficeIMO.Excel {
    internal static class Locking {
        private static readonly AsyncLocal<bool> _noLockScope = new();

        public static IDisposable EnterNoLockScope() {
            var prev = _noLockScope.Value;
            _noLockScope.Value = true;
            return new Scope(() => _noLockScope.Value = prev);
        }

        public static bool IsNoLock => _noLockScope.Value;

        private sealed class Scope : IDisposable {
            private readonly Action _onDispose;
            public Scope(Action onDispose) => _onDispose = onDispose;
            public void Dispose() => _onDispose();
        }

        /// <summary>Serialize the short apply-to-DOM stage only.</summary>
        public static void ExecuteWrite(ReaderWriterLockSlim? lck, Action apply) {
            if (IsNoLock || lck is null || lck.IsWriteLockHeld) {
                apply();
                return;
            }

            lck.EnterWriteLock();
            try {
                apply();
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <summary>Serialize the short apply-to-DOM stage only with return value.</summary>
        public static T ExecuteWrite<T>(ReaderWriterLockSlim? lck, Func<T> apply) {
            if (IsNoLock || lck is null || lck.IsWriteLockHeld) {
                return apply();
            }

            lck.EnterWriteLock();
            try {
                return apply();
            } finally {
                lck.ExitWriteLock();
            }
        }

        /// <summary>
        /// Execute a read of the package model with optional locking. Reading a part the package has not loaded yet
        /// reads the package stream and fills the part, which two readers can't do at once, so these reads take the
        /// upgradeable lock: one at a time, beside plain readers, and still able to upgrade to a write.
        /// </summary>
        public static T ExecuteRead<T>(ReaderWriterLockSlim? lck, Func<T> read) {
            if (IsNoLock || lck is null || lck.IsWriteLockHeld || lck.IsReadLockHeld || lck.IsUpgradeableReadLockHeld) {
                return read();
            }

            lck.EnterUpgradeableReadLock();
            try {
                return read();
            } finally {
                lck.ExitUpgradeableReadLock();
            }
        }

        /// <summary>Execute a read of the package model with optional locking, one reader at a time.</summary>
        public static void ExecuteRead(ReaderWriterLockSlim? lck, Action read) {
            if (IsNoLock || lck is null || lck.IsWriteLockHeld || lck.IsReadLockHeld || lck.IsUpgradeableReadLockHeld) {
                read();
                return;
            }

            lck.EnterUpgradeableReadLock();
            try {
                read();
            } finally {
                lck.ExitUpgradeableReadLock();
            }
        }

    }
}
