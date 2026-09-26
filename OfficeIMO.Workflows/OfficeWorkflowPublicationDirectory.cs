using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;

namespace OfficeIMO.Workflows;

/// <summary>Per-directory publication operations anchored to opened directory handles.</summary>
/// <remarks>
/// Unix operations use directory-relative syscalls. Windows holds every path component without
/// delete sharing while using ordinary file APIs, so an ancestor cannot be replaced mid-transaction.
/// Callers retain this scope through rollback and temporary-file cleanup.
/// </remarks>
internal sealed class OfficeWorkflowPublicationDirectory : IDisposable {
    private readonly string _path;
    private readonly SafeFileHandle? _unixDirectory;
    private readonly List<SafeFileHandle>? _windowsLocks;

    private OfficeWorkflowPublicationDirectory(string path, SafeFileHandle? unixDirectory,
        List<SafeFileHandle>? windowsLocks) {
        _path = path;
        _unixDirectory = unixDirectory;
        _windowsLocks = windowsLocks;
    }

    internal static OfficeWorkflowPublicationDirectory Open(string destinationPath) {
        string directory = Path.GetDirectoryName(Path.GetFullPath(destinationPath))
            ?? throw new ArgumentException("Output path requires a directory.", nameof(destinationPath));
        if (OperatingSystem.IsWindows()) return OpenWindows(directory);
        if (OperatingSystem.IsLinux() || OperatingSystem.IsMacOS()) return OpenUnix(directory);
        throw new PlatformNotSupportedException("Secure redaction publication requires Windows, Linux, or macOS.");
    }

    internal string PathFor(string name) => Path.Combine(_path, name);

    internal void WriteNew(string name, byte[] bytes) {
        ValidateName(name);
        if (_unixDirectory is null) {
            using var stream = new FileStream(PathFor(name), FileMode.CreateNew, FileAccess.Write,
                FileShare.None, 81920, FileOptions.SequentialScan);
            stream.Write(bytes);
            stream.Flush(flushToDisk: true);
            return;
        }
        int descriptor = UnixOpenAt(UnixFd, name, UnixWriteOnly | UnixCreate | UnixExclusive |
            UnixNoFollow | UnixCloseOnExec, 384); // 0600
        if (descriptor < 0) throw UnixError("create", name);
        using var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        using var output = new FileStream(handle, FileAccess.Write, 81920, isAsync: false);
        output.Write(bytes);
        output.Flush(flushToDisk: true);
    }

    internal void MoveNoReplace(string source, string destination) {
        ValidateName(source);
        ValidateName(destination);
        if (_unixDirectory is null) {
            File.Move(PathFor(source), PathFor(destination));
            return;
        }
        int result = UnixRenameNoReplace(source, destination);
        if (result != 0) throw UnixError("publish without replacing another file", destination);
    }

    internal bool TryBackup(string source, string backup) {
        ValidateName(source);
        ValidateName(backup);
        if (_unixDirectory is null) {
            if (!WindowsEntryExists(source)) return false;
            File.Move(PathFor(source), PathFor(backup));
            return true;
        }
        if (UnixRenameNoReplace(source, backup) == 0) return true;
        int error = Marshal.GetLastWin32Error();
        if (error == ErrorNoEntry) return false;
        throw UnixError("stage rollback", source, error);
    }

    internal bool DeleteIfExists(string name) {
        ValidateName(name);
        if (_unixDirectory is null) {
            if (!WindowsEntryExists(name)) return false;
            File.Delete(PathFor(name));
            return true;
        }
        if (UnixUnlinkAt(UnixFd, name, 0) == 0) return true;
        int error = Marshal.GetLastWin32Error();
        if (error == ErrorNoEntry) return false;
        throw UnixError("remove", name, error);
    }

    private bool WindowsEntryExists(string name) {
        try {
            _ = File.GetAttributes(PathFor(name));
            return true;
        } catch (FileNotFoundException) {
            return false;
        } catch (DirectoryNotFoundException) {
            return false;
        }
    }

    private int UnixFd => checked((int)_unixDirectory!.DangerousGetHandle().ToInt64());

    private int UnixRenameNoReplace(string source, string destination) {
        try {
            return OperatingSystem.IsMacOS()
                ? MacRenameAtExclusive(UnixFd, source, UnixFd, destination, 4)
                : LinuxRenameAtNoReplace(UnixFd, source, UnixFd, destination, 1);
        } catch (EntryPointNotFoundException exception) {
            throw new PlatformNotSupportedException("Atomic no-replace publication is unavailable on this Unix runtime.", exception);
        }
    }

    public void Dispose() {
        _unixDirectory?.Dispose();
        if (_windowsLocks is null) return;
        for (int index = _windowsLocks.Count - 1; index >= 0; index--)
            _windowsLocks[index].Dispose();
    }

    private static OfficeWorkflowPublicationDirectory OpenUnix(string directory) {
        int flags = UnixDirectory | UnixNoFollow | UnixCloseOnExec;
        int descriptor = UnixOpen(Path.DirectorySeparatorChar.ToString(), flags, 0);
        if (descriptor < 0) throw UnixError("open root directory", directory);
        var current = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
        try {
            string relative = Path.GetRelativePath(Path.DirectorySeparatorChar.ToString(), directory);
            if (relative != ".") {
                foreach (string segment in relative.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries)) {
                    ValidateName(segment);
                    int parent = checked((int)current.DangerousGetHandle().ToInt64());
                    int child = UnixOpenAt(parent, segment, flags, 0);
                    if (child < 0 && Marshal.GetLastWin32Error() == ErrorNoEntry) {
                        if (UnixMkdirAt(parent, segment, 493) != 0 &&
                            Marshal.GetLastWin32Error() != ErrorAlreadyExists)
                            throw UnixError("create directory", segment);
                        child = UnixOpenAt(parent, segment, flags, 0);
                    }
                    if (child < 0) throw UnixError("open directory", segment);
                    current.Dispose();
                    current = new SafeFileHandle(new IntPtr(child), ownsHandle: true);
                }
            }
            return new OfficeWorkflowPublicationDirectory(directory, current, null);
        } catch {
            current.Dispose();
            throw;
        }
    }

    private static OfficeWorkflowPublicationDirectory OpenWindows(string directory) {
        string root = Path.GetPathRoot(directory)
            ?? throw new ArgumentException("Output path requires a rooted directory.", nameof(directory));
        var locks = new List<SafeFileHandle>();
        try {
            locks.Add(OpenLockedWindowsDirectory(root));
            string current = root;
            string relative = Path.GetRelativePath(root, directory);
            if (relative != ".") {
                foreach (string segment in relative.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries)) {
                    ValidateName(segment);
                    current = Path.Combine(current, segment);
                    Directory.CreateDirectory(current);
                    locks.Add(OpenLockedWindowsDirectory(current));
                }
            }
            return new OfficeWorkflowPublicationDirectory(directory, null, locks);
        } catch {
            foreach (SafeFileHandle handle in locks) handle.Dispose();
            throw;
        }
    }

    private static SafeFileHandle OpenLockedWindowsDirectory(string path) {
        SafeFileHandle handle = WindowsCreateFile(WindowsApiPath(path), 0, FileShare.Read | FileShare.Write,
            IntPtr.Zero, 3, 0x02000000 | 0x00200000, IntPtr.Zero);
        if (handle.IsInvalid) {
            int error = Marshal.GetLastWin32Error();
            handle.Dispose();
            throw new IOException($"Unable to lock publication directory (OS error {error}).");
        }
        if (!WindowsGetFileInformationByHandle(handle, out WindowsFileInformation info) ||
            (info.Attributes & 0x10) == 0 || (info.Attributes & 0x400) != 0) {
            handle.Dispose();
            throw new IOException("Publication directory must be a regular directory without a reparse point.");
        }
        return handle;
    }

    private static string WindowsApiPath(string path) {
        string full = Path.GetFullPath(path);
        if (full.Length < 248 || full.StartsWith(@"\\?\", StringComparison.Ordinal) ||
            full.StartsWith(@"\\.\", StringComparison.Ordinal)) return full;
        return full.StartsWith(@"\\", StringComparison.Ordinal)
            ? @"\\?\UNC\" + full[2..] : @"\\?\" + full;
    }

    private static void ValidateName(string name) {
        if (string.IsNullOrEmpty(name) || name is "." or ".." ||
            name.IndexOfAny(new[] { '/', '\\' }) >= 0)
            throw new ArgumentException("Publication operation requires one file or directory name.", nameof(name));
    }

    private static int UnixDirectory => OperatingSystem.IsMacOS() ? 0x00100000 : 0x00010000;
    private static int UnixNoFollow => OperatingSystem.IsMacOS() ? 0x00000100 : 0x00020000;
    private static int UnixCloseOnExec => OperatingSystem.IsMacOS() ? 0x01000000 : 0x00080000;
    private static int UnixCreate => OperatingSystem.IsMacOS() ? 0x00000200 : 0x00000040;
    private static int UnixExclusive => OperatingSystem.IsMacOS() ? 0x00000800 : 0x00000080;
    private const int UnixWriteOnly = 1;
    private const int ErrorNoEntry = 2;
    private const int ErrorAlreadyExists = 17;

    private static IOException UnixError(string operation, string name) =>
        UnixError(operation, name, Marshal.GetLastWin32Error());
    private static IOException UnixError(string operation, string name, int error) =>
        new($"Unable to {operation} publication entry '{name}' (OS error {error}).");

    [DllImport("libc", EntryPoint = "open", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int UnixOpen(string path, int flags, uint mode);
    [DllImport("libc", EntryPoint = "openat", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int UnixOpenAt(int directory, string path, int flags, uint mode);
    [DllImport("libc", EntryPoint = "mkdirat", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int UnixMkdirAt(int directory, string path, uint mode);
    [DllImport("libc", EntryPoint = "renameat2", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int LinuxRenameAtNoReplace(int sourceDirectory, string source, int destinationDirectory, string destination, uint flags);
    [DllImport("libc", EntryPoint = "renameatx_np", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int MacRenameAtExclusive(int sourceDirectory, string source, int destinationDirectory, string destination, uint flags);
    [DllImport("libc", EntryPoint = "unlinkat", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int UnixUnlinkAt(int directory, string path, int flags);

    [StructLayout(LayoutKind.Sequential)]
    private struct WindowsFileTime { internal uint Low; internal uint High; }
    [StructLayout(LayoutKind.Sequential)]
    private struct WindowsFileInformation {
        internal uint Attributes;
        internal WindowsFileTime CreationTime;
        internal WindowsFileTime LastAccessTime;
        internal WindowsFileTime LastWriteTime;
        internal uint VolumeSerialNumber;
        internal uint FileSizeHigh;
        internal uint FileSizeLow;
        internal uint NumberOfLinks;
        internal uint FileIndexHigh;
        internal uint FileIndexLow;
    }
    [DllImport("kernel32.dll", EntryPoint = "CreateFileW", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern SafeFileHandle WindowsCreateFile(string path, uint desiredAccess, FileShare share,
        IntPtr securityAttributes, uint creationDisposition, uint flags, IntPtr templateFile);
    [DllImport("kernel32.dll", EntryPoint = "GetFileInformationByHandle", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool WindowsGetFileInformationByHandle(SafeFileHandle handle, out WindowsFileInformation information);
}
