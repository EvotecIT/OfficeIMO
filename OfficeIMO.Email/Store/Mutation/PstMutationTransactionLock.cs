using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Email.Store;

/// <summary>Coordinates OfficeIMO mutation transactions for one physical PST across processes.</summary>
internal sealed class PstMutationTransactionLock : IDisposable {
    private const long LockOffset = long.MaxValue - 1;
    private const int LockExclusive = 2;
    private const int LockNonBlocking = 4;
    private const int LockUnlock = 8;
    private static readonly ConcurrentDictionary<string, byte> ProcessLocks =
        new ConcurrentDictionary<string, byte>(StringComparer.Ordinal);
    private readonly FileStream _lockStream;
    private readonly string _identity;
    private bool _disposed;

    private PstMutationTransactionLock(FileStream lockStream, string identity) {
        _lockStream = lockStream;
        _identity = identity;
    }

    internal string Identity => _identity;

    internal static PstMutationTransactionLock Acquire(string sourcePath, SafeFileHandle sourceHandle) {
        string identity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath, sourceHandle);
        if (!ProcessLocks.TryAdd(identity, 0)) {
            throw new IOException("Another OfficeIMO mutation transaction already owns this physical PST.");
        }

        FileStream? lockStream = null;
        try {
            lockStream = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? new FileStream(sourcePath, FileMode.Open, FileAccess.ReadWrite,
                    FileShare.Read | FileShare.Delete, 1, FileOptions.RandomAccess)
                : DuplicateUnixStream(sourceHandle);
            string lockIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath,
                lockStream.SafeFileHandle);
            if (!string.Equals(identity, lockIdentity, StringComparison.Ordinal)) {
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }
            AcquirePhysicalLock(lockStream);
            string currentIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath);
            if (!string.Equals(identity, currentIdentity, StringComparison.Ordinal)) {
                ReleasePhysicalLock(lockStream);
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }
            return new PstMutationTransactionLock(lockStream, identity);
        } catch (UnauthorizedAccessException exception) {
            lockStream?.Dispose();
            ProcessLocks.TryRemove(identity, out _);
            throw new IOException(
                "The physical PST mutation lock could not be acquired with write access.", exception);
        } catch (IOException exception) {
            lockStream?.Dispose();
            ProcessLocks.TryRemove(identity, out _);
            throw new IOException("Another OfficeIMO mutation transaction already owns this physical PST.", exception);
        } catch {
            lockStream?.Dispose();
            ProcessLocks.TryRemove(identity, out _);
            throw;
        }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        try {
            try {
                ReleasePhysicalLock(_lockStream);
            } finally {
                _lockStream.Dispose();
            }
        } finally {
            ProcessLocks.TryRemove(_identity, out _);
        }
    }

    private static void AcquirePhysicalLock(FileStream stream) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            stream.Lock(LockOffset, 1);
            return;
        }
        int descriptor = stream.SafeFileHandle.DangerousGetHandle().ToInt32();
        if (Flock(descriptor, LockExclusive | LockNonBlocking) != 0) {
            throw new IOException("The physical PST lock could not be acquired " +
                "(OS error " + Marshal.GetLastWin32Error() + ").");
        }
    }

    private static FileStream DuplicateUnixStream(SafeFileHandle sourceHandle) {
        int descriptor = Dup(sourceHandle.DangerousGetHandle().ToInt32());
        if (descriptor < 0) {
            throw new IOException("The source PST handle could not be duplicated for physical locking " +
                "(OS error " + Marshal.GetLastWin32Error() + ").");
        }
        var duplicate = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        try {
            return new FileStream(duplicate, FileAccess.Read, 1, isAsync: false);
        } catch {
            duplicate.Dispose();
            throw;
        }
    }

    private static void ReleasePhysicalLock(FileStream stream) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            stream.Unlock(LockOffset, 1);
            return;
        }
        int descriptor = stream.SafeFileHandle.DangerousGetHandle().ToInt32();
        if (Flock(descriptor, LockUnlock) != 0) {
            throw new IOException("The physical PST lock could not be released " +
                "(OS error " + Marshal.GetLastWin32Error() + ").");
        }
    }

    [DllImport("libc", EntryPoint = "flock", SetLastError = true)]
    private static extern int Flock(int descriptor, int operation);

    [DllImport("libc", EntryPoint = "dup", SetLastError = true)]
    private static extern int Dup(int descriptor);
}
