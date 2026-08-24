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
    private const int LinuxDuplicateCloseOnExec = 1030;
    private const int MacDuplicateCloseOnExec = 67;
    private const int LinuxOpenFileDescriptionSetLock = 37;
    private const int MacOpenFileDescriptionSetLock = 90;
    private const short LinuxWriteLock = 1;
    private const short LinuxUnlock = 2;
    private const short MacWriteLock = 3;
    private const short MacUnlock = 2;
    private const int LockExclusive = 2;
    private const int LockNonBlocking = 4;
    private const int LockUnlock = 8;
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
    private readonly PhysicalLockKind _lockKind;
    private bool _disposed;

    private PstMutationTransactionLock(SafeFileHandle lockHandle, string identity,
        PhysicalLockKind lockKind) {
        _lockHandle = lockHandle;
        _identity = identity;
        _lockKind = lockKind;
    }

    internal string Identity => _identity;

    internal static PstMutationTransactionLock OpenUnixSource(string sourcePath,
        out FileStream source) {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux) &&
            !RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            throw new PlatformNotSupportedException(
                "Unix PST mutation locking supports Linux and macOS.");
        }

        try {
            try {
                return OpenUnixSource(sourcePath, GetUnixReadWriteOpenFlags(),
                    preferOpenFileDescriptionLock: true, out source);
            } catch (UnauthorizedAccessException) {
                return OpenUnixSource(sourcePath, GetUnixOpenFlags(),
                    preferOpenFileDescriptionLock: false, out source);
            }
        } catch (UnauthorizedAccessException exception) {
            throw new IOException(
                "The physical PST mutation lock could not be acquired with the required access.", exception);
        } catch (IOException exception) {
            throw new IOException("Another OfficeIMO mutation transaction already owns this physical PST.", exception);
        }
    }

    internal static PstMutationTransactionLock Acquire(string sourcePath,
        SafeFileHandle sourceHandle) {
        string identity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath, sourceHandle);
        if (!ProcessLocks.TryAdd(identity, 0)) {
            throw new IOException("Another OfficeIMO mutation transaction already owns this physical PST.");
        }

        SafeFileHandle? lockHandle = null;
        try {
            lockHandle = OpenAndAcquireWindowsLock(sourcePath);
            string lockIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath,
                lockHandle);
            if (!string.Equals(identity, lockIdentity, StringComparison.Ordinal)) {
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }
            string currentIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath);
            if (!string.Equals(identity, currentIdentity, StringComparison.Ordinal)) {
                ReleasePhysicalLock(lockHandle, PhysicalLockKind.Windows);
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }
            return new PstMutationTransactionLock(lockHandle, identity,
                PhysicalLockKind.Windows);
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
                ReleasePhysicalLock(_lockHandle, _lockKind);
            } finally {
                _lockHandle.Dispose();
            }
        } finally {
            ProcessLocks.TryRemove(_identity, out _);
        }
    }

    private static PstMutationTransactionLock OpenUnixSource(string sourcePath, int flags,
        bool preferOpenFileDescriptionLock, out FileStream source) {
        SafeFileHandle? lockHandle = null;
        SafeFileHandle? sourceHandle = null;
        FileStream? input = null;
        string? identity = null;
        bool processLockHeld = false;
        bool physicalLockHeld = false;
        PhysicalLockKind lockKind = PhysicalLockKind.UnixFlock;
        try {
            lockHandle = OpenUnixHandle(sourcePath, flags);
            identity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath, lockHandle);
            if (!ProcessLocks.TryAdd(identity, 0)) {
                throw new IOException(
                    "Another OfficeIMO mutation transaction already owns this physical PST.");
            }
            processLockHeld = true;

            if (preferOpenFileDescriptionLock && TryAcquireUnixOpenFileDescriptionLock(lockHandle)) {
                lockKind = PhysicalLockKind.UnixOpenFileDescription;
                physicalLockHeld = true;
                input = new FileStream(sourcePath, FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete, 64 * 1024, FileOptions.RandomAccess);
            } else {
                sourceHandle = DuplicateUnixHandle(lockHandle);
                input = new FileStream(sourceHandle, FileAccess.Read, 64 * 1024, isAsync: false);
                sourceHandle = null;
                AcquireUnixFlock(lockHandle);
                lockKind = PhysicalLockKind.UnixFlock;
                physicalLockHeld = true;
            }

            string sourceIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath,
                input.SafeFileHandle);
            string currentIdentity = EmailStorePathIdentity.GetPhysicalIdentityKey(sourcePath);
            if (!string.Equals(identity, sourceIdentity, StringComparison.Ordinal) ||
                !string.Equals(identity, currentIdentity, StringComparison.Ordinal)) {
                throw new IOException(
                    "The source PST path changed while its mutation lock was being acquired.");
            }

            source = input;
            input = null;
            var transactionLock = new PstMutationTransactionLock(lockHandle, identity, lockKind);
            lockHandle = null;
            processLockHeld = false;
            physicalLockHeld = false;
            return transactionLock;
        } finally {
            if (physicalLockHeld && lockHandle != null) {
                try { ReleasePhysicalLock(lockHandle, lockKind); } catch { }
            }
            input?.Dispose();
            sourceHandle?.Dispose();
            lockHandle?.Dispose();
            if (processLockHeld && identity != null) ProcessLocks.TryRemove(identity, out _);
        }
    }

    private static SafeFileHandle OpenAndAcquireWindowsLock(string sourcePath) {
        SafeFileHandle handle = CreateFileWindows(sourcePath, GenericRead,
            ShareRead | ShareWrite | ShareDelete, IntPtr.Zero, OpenExisting,
            FileAttributeNormal, IntPtr.Zero);
        if (!handle.IsInvalid) {
            try {
                AcquireWindowsLock(handle);
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

    private static SafeFileHandle OpenUnixHandle(string sourcePath, int flags) {
        int descriptor = OpenUnix(sourcePath, flags);
        if (descriptor >= 0) return new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        int error = Marshal.GetLastWin32Error();
        bool writeRequested = (flags & 3) == OpenReadWrite;
        if (error == 13 || writeRequested && (error == 1 || error == 30)) {
            throw new UnauthorizedAccessException(
                "The physical PST lock handle could not be opened with the requested access.");
        }
        throw new IOException("The physical PST lock handle could not be opened " +
            "(OS error " + error + ").");
    }

    private static SafeFileHandle DuplicateUnixHandle(SafeFileHandle handle) {
        int command = RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? MacDuplicateCloseOnExec
            : LinuxDuplicateCloseOnExec;
        int descriptor = FcntlUnix(handle.DangerousGetHandle().ToInt32(), command, 0);
        if (descriptor >= 0) {
            return new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        }
        throw new IOException("The physical PST source handle could not be duplicated " +
            "(OS error " + Marshal.GetLastWin32Error() + ").");
    }

    private static bool TryAcquireUnixOpenFileDescriptionLock(SafeFileHandle handle) {
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
        if (result == 0) return true;
        int error = Marshal.GetLastWin32Error();
        if (error == 22 || error == 45 || error == 95) return false;
        throw new IOException("The physical PST lock could not be acquired " +
            "(OS error " + error + ").");
    }

    private static void AcquireWindowsLock(SafeFileHandle handle) {
        if (LockFileWindows(handle, unchecked((uint)LockOffset),
                unchecked((uint)((ulong)LockOffset >> 32)), 1, 0)) return;
        throw new IOException("The physical PST lock could not be acquired " +
            "(OS error " + Marshal.GetLastWin32Error() + ").");
    }

    private static void AcquireUnixFlock(SafeFileHandle handle) {
        if (FlockUnix(handle.DangerousGetHandle().ToInt32(),
                LockExclusive | LockNonBlocking) == 0) return;
        throw new IOException("The physical PST lock could not be acquired " +
            "(OS error " + Marshal.GetLastWin32Error() + ").");
    }

    private static void ReleasePhysicalLock(SafeFileHandle handle, PhysicalLockKind lockKind) {
        if (lockKind == PhysicalLockKind.Windows) {
            if (!UnlockFileWindows(handle, unchecked((uint)LockOffset),
                    unchecked((uint)((ulong)LockOffset >> 32)), 1, 0)) {
                throw new IOException("The physical PST lock could not be released " +
                    "(OS error " + Marshal.GetLastWin32Error() + ").");
            }
            return;
        }
        if (lockKind == PhysicalLockKind.UnixFlock) {
            if (FlockUnix(handle.DangerousGetHandle().ToInt32(), LockUnlock) == 0) return;
            throw new IOException("The physical PST lock could not be released " +
                "(OS error " + Marshal.GetLastWin32Error() + ").");
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

    internal static int GetUnixOpenFlags() => OpenReadOnly |
        (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? MacOpenCloseOnExec
            : LinuxOpenCloseOnExec);

    private static int GetUnixReadWriteOpenFlags() => OpenReadWrite |
        (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? MacOpenCloseOnExec
            : LinuxOpenCloseOnExec);

    private enum PhysicalLockKind {
        Windows,
        UnixOpenFileDescription,
        UnixFlock
    }

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

    [DllImport("libc", EntryPoint = "flock", SetLastError = true)]
    private static extern int FlockUnix(int descriptor, int operation);

    [DllImport("libc", EntryPoint = "fcntl", SetLastError = true)]
    private static extern int FcntlUnix(int descriptor, int command, int argument);

    [DllImport("libc", EntryPoint = "fcntl", SetLastError = true)]
    private static extern int SetLinuxFileLock(int descriptor, int command,
        ref LinuxFileLock fileLock);

    [DllImport("libc", EntryPoint = "fcntl", SetLastError = true)]
    private static extern int SetMacFileLock(int descriptor, int command,
        ref MacFileLock fileLock);
}
