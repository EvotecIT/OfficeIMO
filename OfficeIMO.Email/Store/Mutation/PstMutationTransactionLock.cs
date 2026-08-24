using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Email.Store;

/// <summary>Coordinates OfficeIMO mutation transactions for one physical PST across processes.</summary>
internal sealed class PstMutationTransactionLock : IDisposable {
    private const long LockOffset = long.MaxValue - 1;
    private const int OpenReadOnly = 0;
    private const int OpenReadWrite = 2;
    private const int LinuxOpenCloseOnExec = 0x00080000;
    private const int MacOpenCloseOnExec = 0x01000000;
    private const int LockExclusive = 2;
    private const int LockNonBlocking = 4;
    private const int LockUnlock = 8;
    private const int ErrorBadFileDescriptor = 9;
    private const uint GenericRead = 0x80000000;
    private const uint ShareRead = 0x00000001;
    private const uint ShareWrite = 0x00000002;
    private const uint ShareDelete = 0x00000004;
    private const uint OpenExisting = 3;
    private const uint FileAttributeNormal = 0x00000080;
    private static readonly ConcurrentDictionary<string, byte> ProcessLocks =
        new ConcurrentDictionary<string, byte>(StringComparer.Ordinal);
    private readonly SafeFileHandle _lockHandle;
    private readonly string _identity;
    private bool _disposed;

    private PstMutationTransactionLock(SafeFileHandle lockHandle, string identity) {
        _lockHandle = lockHandle;
        _identity = identity;
    }

    internal string Identity => _identity;

    internal static PstMutationTransactionLock Acquire(string sourcePath, SafeFileHandle sourceHandle) {
        string identity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath, sourceHandle);
        if (!ProcessLocks.TryAdd(identity, 0)) {
            throw new IOException("Another OfficeIMO mutation transaction already owns this physical PST.");
        }

        SafeFileHandle? lockHandle = null;
        try {
            lockHandle = OpenAndAcquirePhysicalLock(sourcePath);
            string lockIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath,
                lockHandle);
            if (!string.Equals(identity, lockIdentity, StringComparison.Ordinal)) {
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }
            string currentIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath);
            if (!string.Equals(identity, currentIdentity, StringComparison.Ordinal)) {
                ReleasePhysicalLock(lockHandle);
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }
            return new PstMutationTransactionLock(lockHandle, identity);
        } catch (UnauthorizedAccessException exception) {
            lockHandle?.Dispose();
            ProcessLocks.TryRemove(identity, out _);
            throw new IOException(
                "The physical PST mutation lock could not be acquired with the required access.", exception);
        } catch (IOException exception) {
            lockHandle?.Dispose();
            ProcessLocks.TryRemove(identity, out _);
            throw new IOException("Another OfficeIMO mutation transaction already owns this physical PST.", exception);
        } catch {
            lockHandle?.Dispose();
            ProcessLocks.TryRemove(identity, out _);
            throw;
        }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        try {
            try {
                ReleasePhysicalLock(_lockHandle);
            } finally {
                _lockHandle.Dispose();
            }
        } finally {
            ProcessLocks.TryRemove(_identity, out _);
        }
    }

    private static SafeFileHandle OpenAndAcquirePhysicalLock(string sourcePath) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            SafeFileHandle handle = CreateFileWindows(sourcePath, GenericRead,
                ShareRead | ShareWrite | ShareDelete, IntPtr.Zero, OpenExisting,
                FileAttributeNormal, IntPtr.Zero);
            if (!handle.IsInvalid) {
                try {
                    AcquirePhysicalLock(handle);
                    return handle;
                } catch {
                    handle.Dispose();
                    throw;
                }
            }
            int error = Marshal.GetLastWin32Error();
            handle.Dispose();
            if (error == 5) {
                throw new UnauthorizedAccessException(
                    "The physical PST lock handle requires read access.");
            }
            throw new IOException("The physical PST lock handle could not be opened " +
                "(OS error " + error + ").");
        }
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux) &&
            !RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            throw new PlatformNotSupportedException(
                "Physical PST mutation locking supports Windows, Linux, and macOS.");
        }
        SafeFileHandle readHandle = OpenUnixHandle(sourcePath, GetUnixOpenFlags());
        int result = FlockUnix(readHandle, LockExclusive | LockNonBlocking);
        if (result == 0) return readHandle;
        int lockError = Marshal.GetLastWin32Error();
        readHandle.Dispose();
        if (lockError != ErrorBadFileDescriptor) {
            throw new IOException("The physical PST lock could not be acquired " +
                "(OS error " + lockError + ").");
        }

        // Local Unix flock accepts a read-only descriptor. Some network filesystems emulate
        // it with POSIX write locks and return EBADF, so retry with write access only there.
        SafeFileHandle writeHandle = OpenUnixHandle(sourcePath, GetUnixReadWriteOpenFlags());
        try {
            AcquirePhysicalLock(writeHandle);
            return writeHandle;
        } catch {
            writeHandle.Dispose();
            throw;
        }
    }

    private static SafeFileHandle OpenUnixHandle(string sourcePath, int flags) {
        int descriptor = OpenUnix(sourcePath, flags);
        if (descriptor >= 0) return new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        int error = Marshal.GetLastWin32Error();
        if (error == 13) {
            throw new UnauthorizedAccessException(
                "The physical PST lock handle could not be opened with the access required by this filesystem.");
        }
        throw new IOException("The physical PST lock handle could not be opened " +
            "(OS error " + error + ").");
    }

    private static void AcquirePhysicalLock(SafeFileHandle handle) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            if (LockFileWindows(handle, unchecked((uint)LockOffset),
                    unchecked((uint)((ulong)LockOffset >> 32)), 1, 0)) return;
            throw new IOException("The physical PST lock could not be acquired " +
                "(OS error " + Marshal.GetLastWin32Error() + ").");
        }
        int result = FlockUnix(handle, LockExclusive | LockNonBlocking);
        if (result == 0) return;
        throw new IOException("The physical PST lock could not be acquired " +
            "(OS error " + Marshal.GetLastWin32Error() + ").");
    }

    private static void ReleasePhysicalLock(SafeFileHandle handle) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            if (!UnlockFileWindows(handle, unchecked((uint)LockOffset),
                    unchecked((uint)((ulong)LockOffset >> 32)), 1, 0)) {
                throw new IOException("The physical PST lock could not be released " +
                    "(OS error " + Marshal.GetLastWin32Error() + ").");
            }
            return;
        }
        int result = FlockUnix(handle, LockUnlock);
        if (result == 0) return;
        throw new IOException("The physical PST lock could not be released " +
            "(OS error " + Marshal.GetLastWin32Error() + ").");
    }

    private static int FlockUnix(SafeFileHandle handle, int operation) =>
        FlockUnix(handle.DangerousGetHandle().ToInt32(), operation);

    internal static int GetUnixOpenFlags() => OpenReadOnly |
        (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? MacOpenCloseOnExec
            : LinuxOpenCloseOnExec);

    private static int GetUnixReadWriteOpenFlags() => OpenReadWrite |
        (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? MacOpenCloseOnExec
            : LinuxOpenCloseOnExec);

    [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode,
        SetLastError = true)]
    private static extern SafeFileHandle CreateFileWindows(string fileName, uint desiredAccess,
        uint shareMode, IntPtr securityAttributes, uint creationDisposition,
        uint flagsAndAttributes, IntPtr templateFile);

    [DllImport("kernel32.dll", EntryPoint = "LockFile", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool LockFileWindows(SafeFileHandle file, uint offsetLow,
        uint offsetHigh, uint bytesLow, uint bytesHigh);

    [DllImport("kernel32.dll", EntryPoint = "UnlockFile", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnlockFileWindows(SafeFileHandle file, uint offsetLow,
        uint offsetHigh, uint bytesLow, uint bytesHigh);

    [DllImport("libc", EntryPoint = "open", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern int OpenUnix(string path, int flags);

    [DllImport("libc", EntryPoint = "flock", SetLastError = true)]
    private static extern int FlockUnix(int descriptor, int operation);
}
