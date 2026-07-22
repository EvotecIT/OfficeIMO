using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;

namespace OfficeIMO.Email.Store;

internal static partial class EmailStorePathIdentity {
    private const int FileCaseSensitiveInfo = 23;
    private const uint FileCaseSensitiveDirectory = 0x00000001;
    private const uint OpenExisting = 3;
    private const uint FileFlagBackupSemantics = 0x02000000;

    private static bool TryGetWindowsDirectoryCaseInsensitive(
        string directory, out bool caseInsensitive) {
        caseInsensitive = true;
        try {
            using SafeFileHandle handle = OpenPathHandle(directory);
            if (handle.IsInvalid) return false;
            if (!GetFileInformationByHandleEx(handle, FileCaseSensitiveInfo,
                    out FileCaseSensitiveInformation information,
                    (uint)Marshal.SizeOf<FileCaseSensitiveInformation>())) {
                return false;
            }

            caseInsensitive = (information.Flags & FileCaseSensitiveDirectory) == 0;
            return true;
        } catch (DllNotFoundException) {
            return false;
        } catch (EntryPointNotFoundException) {
            return false;
        }
    }

    private static bool TryResolvePhysicalPath(string path, out string resolvedPath) {
        string fullPath = Path.GetFullPath(path);
        string existingPath = TrimEndingDirectorySeparators(fullPath);
        var missingSegments = new Stack<string>();
        while (!File.Exists(existingPath) && !Directory.Exists(existingPath)) {
            string? parent = Path.GetDirectoryName(existingPath);
            if (string.IsNullOrEmpty(parent) || string.Equals(parent, existingPath, StringComparison.Ordinal)) {
                resolvedPath = fullPath;
                return false;
            }

            string name = Path.GetFileName(existingPath);
            if (!string.IsNullOrEmpty(name)) missingSegments.Push(name);
            existingPath = parent;
        }

        if (!TryResolveExistingPhysicalPath(existingPath, out string resolvedExisting)) {
            resolvedPath = fullPath;
            return false;
        }

        resolvedPath = resolvedExisting;
        foreach (string segment in missingSegments) resolvedPath = Path.Combine(resolvedPath, segment);
        resolvedPath = TrimEndingDirectorySeparators(Path.GetFullPath(resolvedPath));
        return true;
    }

    private static bool TryResolveExistingPhysicalPath(string path, out string resolvedPath) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            return TryResolveWindowsPhysicalPath(path, out resolvedPath);
        }
        return TryResolvePosixPhysicalPath(path, out resolvedPath);
    }

    private static bool TryResolveWindowsPhysicalPath(string path, out string resolvedPath) {
        resolvedPath = path;
        try {
            using SafeFileHandle handle = OpenPathHandle(path);
            if (handle.IsInvalid) return false;
            var buffer = new StringBuilder(32768);
            uint length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
            if (length == 0 || length >= buffer.Capacity) return false;
            resolvedPath = NormalizeWindowsFinalPath(buffer.ToString());
            return true;
        } catch (DllNotFoundException) {
            return false;
        } catch (EntryPointNotFoundException) {
            return false;
        }
    }

    private static bool TryResolvePosixPhysicalPath(string path, out string resolvedPath) {
        resolvedPath = path;
        IntPtr pointer = IntPtr.Zero;
        try {
            pointer = RealPath(path, IntPtr.Zero);
            if (pointer == IntPtr.Zero) return false;
            string? value = Marshal.PtrToStringAnsi(pointer);
            if (string.IsNullOrEmpty(value)) return false;
            resolvedPath = value;
            return true;
        } catch (DllNotFoundException) {
            return false;
        } catch (EntryPointNotFoundException) {
            return false;
        } finally {
            if (pointer != IntPtr.Zero) Free(pointer);
        }
    }

    private static SafeFileHandle OpenPathHandle(string path) => CreateFile(
        path, 0, FileShare.Read | FileShare.Write | FileShare.Delete,
        IntPtr.Zero, OpenExisting, FileFlagBackupSemantics, IntPtr.Zero);

    private static string NormalizeWindowsFinalPath(string path) {
        const string uncPrefix = @"\\?\UNC\";
        const string devicePrefix = @"\\?\";
        if (path.StartsWith(uncPrefix, StringComparison.OrdinalIgnoreCase)) {
            return string.Concat(@"\\", path.Substring(uncPrefix.Length));
        }
        return path.StartsWith(devicePrefix, StringComparison.OrdinalIgnoreCase)
            ? path.Substring(devicePrefix.Length)
            : path;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FileCaseSensitiveInformation {
        internal uint Flags;
    }

    [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode,
        SetLastError = true)]
    private static extern SafeFileHandle CreateFile(string fileName, uint desiredAccess,
        FileShare shareMode, IntPtr securityAttributes, uint creationDisposition,
        uint flagsAndAttributes, IntPtr templateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetFileInformationByHandleEx(SafeFileHandle file,
        int fileInformationClass, out FileCaseSensitiveInformation fileInformation,
        uint bufferSize);

    [DllImport("kernel32.dll", EntryPoint = "GetFinalPathNameByHandleW",
        CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetFinalPathNameByHandle(SafeFileHandle file,
        StringBuilder filePath, uint filePathLength, uint flags);

    [DllImport("libc", EntryPoint = "realpath", CharSet = CharSet.Ansi,
        SetLastError = true)]
    private static extern IntPtr RealPath(string path, IntPtr resolvedPath);

    [DllImport("libc", EntryPoint = "free")]
    private static extern void Free(IntPtr pointer);
}
