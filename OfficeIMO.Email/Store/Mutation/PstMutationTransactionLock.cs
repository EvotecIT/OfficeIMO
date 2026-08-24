using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Email.Store;

/// <summary>Coordinates OfficeIMO mutation transactions for one physical PST across processes.</summary>
internal sealed class PstMutationTransactionLock : IDisposable {
    private const long LockOffset = long.MaxValue - 1;
    private const int OpenReadWrite = 2;
    private const int LinuxOpenCloseOnExec = 0x00080000;
    private const int MacOpenCloseOnExec = 0x01000000;
    private const int LinuxOpenFileDescriptionSetLock = 37;
    private const int MacOpenFileDescriptionSetLock = 90;
    private const short LinuxWriteLock = 1;
    private const short LinuxUnlock = 2;
    private const short MacWriteLock = 3;
    private const short MacUnlock = 2;
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
            lockHandle = OpenPhysicalLockHandle(sourcePath);
            string lockIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath,
                lockHandle);
            if (!string.Equals(identity, lockIdentity, StringComparison.Ordinal)) {
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }
            AcquirePhysicalLock(lockHandle);
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

    private static SafeFileHandle OpenPhysicalLockHandle(string sourcePath) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            SafeFileHandle handle = CreateFileWindows(sourcePath, GenericRead,
                ShareRead | ShareWrite | ShareDelete, IntPtr.Zero, OpenExisting,
                FileAttributeNormal, IntPtr.Zero);
            if (!handle.IsInvalid) return handle;
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
        int descriptor = OpenUnix(sourcePath, GetUnixOpenFlags());
        if (descriptor >= 0) {
            return new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        }
        int unixError = Marshal.GetLastWin32Error();
        if (unixError == 13) {
            throw new UnauthorizedAccessException(
                "The physical PST lock handle requires write access.");
        }
        throw new IOException("The physical PST lock handle could not be opened " +
            "(OS error " + unixError + ").");
    }

    private static void AcquirePhysicalLock(SafeFileHandle handle) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            if (LockFileWindows(handle, unchecked((uint)LockOffset),
                    unchecked((uint)((ulong)LockOffset >> 32)), 1, 0)) return;
            throw new IOException("The physical PST lock could not be acquired " +
                "(OS error " + Marshal.GetLastWin32Error() + ").");
        }
        int descriptor = handle.DangerousGetHandle().ToInt32();
        int result;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            var fileLock = new MacFileLock {
                Start = LockOffset,
                Length = 1,
                Type = MacWriteLock,
                Whence = 0
            };
            result = SetMacFileLock(descriptor, MacOpenFileDescriptionSetLock, ref fileLock);
        } else {
            var fileLock = new LinuxFileLock {
                Type = LinuxWriteLock,
                Whence = 0,
                Start = LockOffset,
                Length = 1
            };
            result = SetLinuxFileLock(descriptor, LinuxOpenFileDescriptionSetLock, ref fileLock);
        }
        if (result == 0) return;
        int error = Marshal.GetLastWin32Error();
        ThrowIfOpenFileDescriptionLockIsUnsupported(error);
        throw new IOException("The physical PST lock could not be acquired " +
            "(OS error " + error + ").");
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
        int descriptor = handle.DangerousGetHandle().ToInt32();
        int result;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            var fileLock = new MacFileLock {
                Start = LockOffset,
                Length = 1,
                Type = MacUnlock,
                Whence = 0
            };
            result = SetMacFileLock(descriptor, MacOpenFileDescriptionSetLock, ref fileLock);
        } else {
            var fileLock = new LinuxFileLock {
                Type = LinuxUnlock,
                Whence = 0,
                Start = LockOffset,
                Length = 1
            };
            result = SetLinuxFileLock(descriptor, LinuxOpenFileDescriptionSetLock, ref fileLock);
        }
        if (result == 0) return;
        throw new IOException("The physical PST lock could not be released " +
            "(OS error " + Marshal.GetLastWin32Error() + ").");
    }

    private static void ThrowIfOpenFileDescriptionLockIsUnsupported(int error) {
        if (error == 22 || error == 45 || error == 95) {
            throw new PlatformNotSupportedException(
                "The source filesystem does not support open-file-description locks.");
        }
    }

    internal static int GetUnixOpenFlags() => OpenReadWrite |
        (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? MacOpenCloseOnExec
            : LinuxOpenCloseOnExec);

    [StructLayout(LayoutKind.Sequential)]
    private struct LinuxFileLock {
        internal short Type;
        internal short Whence;
        internal long Start;
        internal long Length;
        internal int ProcessId;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MacFileLock {
        internal long Start;
        internal long Length;
        internal int ProcessId;
        internal short Type;
        internal short Whence;
    }

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

    [DllImport("libc", EntryPoint = "fcntl", SetLastError = true)]
    private static extern int SetLinuxFileLock(int descriptor, int command,
        ref LinuxFileLock fileLock);

    [DllImport("libc", EntryPoint = "fcntl", SetLastError = true)]
    private static extern int SetMacFileLock(int descriptor, int command,
        ref MacFileLock fileLock);
}
