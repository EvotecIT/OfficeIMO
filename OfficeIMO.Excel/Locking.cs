using System;
using System.Threading;

namespace OfficeIMO.Excel
{
    internal static class Locking
    {
        private static readonly AsyncLocal<bool> _noLockScope = new();

        public static IDisposable EnterNoLockScope()
        {
            var prev = _noLockScope.Value;
            _noLockScope.Value = true;
            return new Scope(() => _noLockScope.Value = prev);
        }

        public static bool IsNoLock => _noLockScope.Value;

        private sealed class Scope : IDisposable
        {
            private readonly Action _onDispose;
            public Scope(Action onDispose) => _onDispose = onDispose;
            public void Dispose() => _onDispose();
        }

        /// <summary>Serialize the short apply-to-DOM stage only.</summary>
        public static void ExecuteWrite(ReaderWriterLockSlim? lck, Action apply)
        {
            if (IsNoLock || lck is null) 
            { 
                apply(); 
                return; 
            }
            
            lck.EnterWriteLock();
            try 
            { 
                apply(); 
            }
            finally 
            { 
                lck.ExitWriteLock(); 
            }
        }

        /// <summary>Serialize the short apply-to-DOM stage only with return value.</summary>
        public static T ExecuteWrite<T>(ReaderWriterLockSlim? lck, Func<T> apply)
        {
            if (IsNoLock || lck is null) 
            { 
                return apply();
            }
            
            lck.EnterWriteLock();
            try 
            { 
                return apply(); 
            }
            finally 
            { 
                lck.ExitWriteLock(); 
            }
        }

        /// <summary>Execute read operation with optional locking.</summary>
        public static T ExecuteRead<T>(ReaderWriterLockSlim? lck, Func<T> read)
        {
            if (IsNoLock || lck is null)
            {
                return read();
            }

            lck.EnterReadLock();
            try
            {
                return read();
            }
            finally
            {
                lck.ExitReadLock();
            }
        }

        /// <summary>Execute read operation with optional locking.</summary>
        public static void ExecuteRead(ReaderWriterLockSlim? lck, Action read)
        {
            if (IsNoLock || lck is null)
            {
                read();
                return;
            }

            lck.EnterReadLock();
            try
            {
                read();
            }
            finally
            {
                lck.ExitReadLock();
            }
        }

    }
}