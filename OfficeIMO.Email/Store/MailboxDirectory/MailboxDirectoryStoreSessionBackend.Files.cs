using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.Email.Store;

internal sealed partial class MailboxDirectoryStoreSessionBackend {
    private bool IsRegularMailboxFile(string path) {
        if (!TryOpenRegularMailboxFile(path, 4 * 1024, out FileStream stream)) return false;
        stream.Dispose();
        return true;
    }

    private FileStream OpenRegularMailboxFile(string path) {
        if (TryOpenRegularMailboxFile(path, 64 * 1024, out FileStream stream)) return stream;
        throw new IOException("The mailbox item is no longer a regular readable file.");
    }

    private bool TryOpenRegularMailboxFile(
        string path,
        int bufferSize,
        out FileStream stream) {
        stream = null!;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            try {
                stream = new FileStream(
                    path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize,
                    FileOptions.SequentialScan);
                if (!OpenedPathRemainsInsideRoot(stream.SafeFileHandle, path)) {
                    stream.Dispose();
                    stream = null!;
                    return false;
                }
                return true;
            } catch (Exception exception) when (
                exception is IOException || exception is UnauthorizedAccessException) {
                return false;
            }
        }

        int nonBlocking = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? 0x0004 : 0x0800;
        int closeOnExec = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? 0x01000000 : 0x00080000;
        int noFollow = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? 0x00000100 : 0x00020000;
        int descriptor = OpenUnixPathWithoutLinks(path, nonBlocking | closeOnExec | noFollow);
        if (descriptor < 0) return false;
        if (!IsRegularUnixDescriptor(descriptor)) {
            CloseUnix(descriptor);
            return false;
        }
        if (SeekUnix(descriptor, 0L, 1) < 0L) {
            CloseUnix(descriptor);
            return false;
        }

        var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        try {
            stream = new FileStream(handle, FileAccess.Read, bufferSize, isAsync: false);
            return true;
        } catch {
            handle.Dispose();
            throw;
        }
    }

    private static bool IsRegularUnixDescriptor(int descriptor) =>
        GetUnixFileStatus(new IntPtr(descriptor), out UnixFileStatus status) == 0
        && (status.Mode & 0xF000) == 0x8000;

    private bool OpenedPathRemainsInsideRoot(SafeFileHandle handle, string requestedPath) {
        var buffer = new StringBuilder(1024);
        uint length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
        if (length == 0) return false;
        if (length >= buffer.Capacity) {
            buffer = new StringBuilder(checked((int)length + 1));
            length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
            if (length == 0 || length >= buffer.Capacity) return false;
        }
        return IsResolvedPathInsideRoot(EmailStorePathIdentity.NormalizeWindowsFinalPath(buffer.ToString()), requestedPath);
    }

    private bool IsResolvedPathInsideRoot(string resolvedPath, string requestedPath) {
        string normalized;
        try {
            normalized = Path.GetFullPath(resolvedPath);
        } catch (Exception exception) when (
            exception is ArgumentException || exception is NotSupportedException || exception is PathTooLongException) {
            return false;
        }
        string requested = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? EmailStorePathIdentity.ResolvePhysicalPath(requestedPath)
            : Path.GetFullPath(requestedPath);
        return normalized.StartsWith(_windowsOpenRoot, _rootComparison)
            && string.Equals(normalized, requested, _rootComparison);
    }

    private int OpenUnixPathWithoutLinks(string path, int fileFlags) {
        string normalized = Path.GetFullPath(path);
        if (!normalized.StartsWith(_root, _rootComparison)) return -1;
        string relative = normalized.Substring(_root.Length);
        string[] segments = relative.Split(
            new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
            StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0 || segments.Any(segment => segment == "." || segment == "..")) return -1;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            const int noFollow = 0x00000100;
            const int noFollowAny = 0x20000000;
            string canonicalPath = Path.Combine(_unixOpenRoot, string.Join(Path.DirectorySeparatorChar.ToString(), segments));
            return OpenUnix(canonicalPath, (fileFlags & ~noFollow) | noFollowAny);
        }

        const int linuxCloseOnExec = 0x00080000;
        const int linuxNoFollow = 0x00020000;
        const int linuxDirectory = 0x00010000;
        int directory = OpenUnix(
            TrimTrailingDirectorySeparators(_unixOpenRoot),
            linuxCloseOnExec | linuxNoFollow | linuxDirectory);
        if (directory < 0) return -1;
        try {
            for (int index = 0; index < segments.Length - 1; index++) {
                int child = OpenAtUnix(
                    directory,
                    segments[index],
                    linuxCloseOnExec | linuxNoFollow | linuxDirectory);
                if (child < 0) return -1;
                CloseUnix(directory);
                directory = child;
            }
            return OpenAtUnix(directory, segments[segments.Length - 1], fileFlags);
        } finally {
            CloseUnix(directory);
        }
    }

    internal static string TrimTrailingDirectorySeparators(string path) {
        string root = Path.GetPathRoot(path) ?? string.Empty;
        int length = path.Length;
        while (length > root.Length
               && (path[length - 1] == Path.DirectorySeparatorChar
                   || path[length - 1] == Path.AltDirectorySeparatorChar)) {
            length--;
        }
        return length == path.Length ? path : path.Substring(0, length);
    }

    private static string? ResolveUnixRealPath(string path) {
        IntPtr resolved = RealPathUnix(path, IntPtr.Zero);
        if (resolved == IntPtr.Zero) return null;
        try {
            return Marshal.PtrToStringAnsi(resolved);
        } finally {
            FreeUnix(resolved);
        }
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandle(
        SafeFileHandle file,
        StringBuilder filePath,
        uint filePathLength,
        uint flags);

    [DllImport("libc", EntryPoint = "open", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int OpenUnix(string path, int flags);

    [DllImport("libc", EntryPoint = "openat", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int OpenAtUnix(int directoryDescriptor, string path, int flags);

    [DllImport("libc", EntryPoint = "lseek", SetLastError = true)]
    private static extern long SeekUnix(int descriptor, long offset, int origin);

    [DllImport("System.Native", EntryPoint = "SystemNative_FStat", SetLastError = true)]
    private static extern int GetUnixFileStatus(IntPtr descriptor, out UnixFileStatus status);

    [DllImport("libc", EntryPoint = "close", SetLastError = true)]
    private static extern int CloseUnix(int descriptor);

    [StructLayout(LayoutKind.Sequential)]
    private struct UnixFileStatus {
        internal int Flags;
        internal int Mode;
        internal uint Uid;
        internal uint Gid;
        internal long Size;
        internal long AccessTime;
        internal long AccessTimeNanoseconds;
        internal long ModificationTime;
        internal long ModificationTimeNanoseconds;
        internal long ChangeTime;
        internal long ChangeTimeNanoseconds;
        internal long BirthTime;
        internal long BirthTimeNanoseconds;
        internal long Device;
        internal long RawDevice;
        internal long Inode;
        internal uint UserFlags;
        internal int HardLinkCount;
    }

    [DllImport("libc", EntryPoint = "realpath", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern IntPtr RealPathUnix(string path, IntPtr resolvedPath);

    [DllImport("libc", EntryPoint = "free")]
    private static extern void FreeUnix(IntPtr pointer);

}
