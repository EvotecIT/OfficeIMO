using Microsoft.Win32.SafeHandles;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.Internal {
    internal static partial class OfficePathIdentity {
        private static FileStream OpenWindowsRegularFileForRead(string path, int bufferSize) {
            SafeFileHandle handle = CreateFile(path, GenericRead, FileShare.Read,
                IntPtr.Zero, OpenExisting, FileFlagOpenReparsePoint | FileFlagSequentialScan, IntPtr.Zero);
            if (handle.IsInvalid) {
                handle.Dispose();
                throw WindowsIdentityError(path, "open");
            }
            try {
                if (!GetFileInformationByHandle(handle, out ByHandleFileInformation information)) {
                    throw WindowsIdentityError(path, "inspect");
                }
                if ((information.FileAttributes & (FileAttributeDirectory | FileAttributeReparsePoint)) != 0
                    || GetFileType(handle) != FileTypeDisk) {
                    throw new InvalidDataException("The filesystem entry is not a regular file.");
                }
                _ = GetWindowsMetadata(path, handle);
                return new FileStream(handle, FileAccess.Read, bufferSize, isAsync: false);
            } catch {
                handle.Dispose();
                throw;
            }
        }

        private const int FileIdInfo = 18;
        private const int FileCaseSensitiveInfo = 23;
        private const uint FileCaseSensitiveDirectory = 0x00000001;
        private const uint OpenExisting = 3;
        private const uint GenericRead = 0x80000000;
        private const uint FileFlagBackupSemantics = 0x02000000;
        private const uint FileFlagOpenReparsePoint = 0x00200000;
        private const uint FileFlagSequentialScan = 0x08000000;
        private const uint FileTypeDisk = 0x0001;
        private const uint InvalidFileAttributes = 0xffffffff;
        private const uint FileAttributeReparsePoint = 0x00000400;
        private const uint FileAttributeDirectory = 0x00000010;
        private const int ErrorFileNotFound = 2;
        private const int ErrorPathNotFound = 3;
        private const int ErrorInvalidFunction = 1;
        private const int ErrorNotSupported = 50;
        private const int ErrorInvalidParameter = 87;
        private const int ErrorCallNotImplemented = 120;

        private static string ResolveWindowsExistingPath(string path) {
            using (SafeFileHandle handle = OpenWindowsPathHandle(path)) {
                if (handle.IsInvalid) throw WindowsIdentityError(path, "open");
                return GetWindowsFinalPath(handle);
            }
        }

        private static bool TryGetWindowsMetadata(string path, out OfficeFileMetadata metadata) {
            using (SafeFileHandle handle = OpenWindowsPathHandle(path)) {
                if (!handle.IsInvalid) {
                    metadata = GetWindowsMetadata(path, handle);
                    return true;
                }
                int error = Marshal.GetLastWin32Error();
                if (error == ErrorFileNotFound || error == ErrorPathNotFound) {
                    metadata = default(OfficeFileMetadata);
                    return false;
                }
                throw WindowsIdentityError(path, "open", error);
            }
        }

        private static OfficeFileMetadata GetWindowsMetadata(string path, SafeFileHandle handle) {
            if (!GetFileInformationByHandle(handle, out ByHandleFileInformation legacy)) {
                throw WindowsIdentityError(path, "read its link count");
            }
            string authority = GetWindowsAuthority(GetWindowsFinalPath(handle));
            OfficePhysicalFileIdentity identity;
            ulong legacyFileIndex = ((ulong)legacy.FileIndexHigh << 32) | legacy.FileIndexLow;
            try {
                if (GetFileInformationByHandleEx(handle, FileIdInfo, out FileIdInformation fileId,
                        (uint)Marshal.SizeOf<FileIdInformation>())) {
                    identity = OfficePhysicalFileIdentity.CreateWindowsExtended(authority,
                        fileId.VolumeSerialNumber, fileId.FileId.LowPart, fileId.FileId.HighPart,
                        legacyFileIndex);
                } else {
                    int error = Marshal.GetLastWin32Error();
                    if (!IsUnsupportedExtendedFileId(error)) {
                        throw WindowsIdentityError(path, "read its 128-bit file identity", error);
                    }
                    identity = CreateLegacyWindowsIdentity(authority, legacy.VolumeSerialNumber,
                        legacyFileIndex);
                }
            } catch (EntryPointNotFoundException) {
                identity = CreateLegacyWindowsIdentity(authority, legacy.VolumeSerialNumber,
                    legacyFileIndex);
            }
            return new OfficeFileMetadata(identity, legacy.NumberOfLinks, 0,
                (legacy.FileAttributes & FileAttributeDirectory) != 0);
        }

        private static OfficePhysicalFileIdentity CreateLegacyWindowsIdentity(string authority,
            ulong volume, ulong fileIndex) =>
            OfficePhysicalFileIdentity.CreateWindowsLegacy(authority, volume, fileIndex);

        private static bool IsUnsupportedExtendedFileId(int error) =>
            error == ErrorInvalidFunction || error == ErrorNotSupported ||
            error == ErrorInvalidParameter || error == ErrorCallNotImplemented;

        private static bool TryGetWindowsDirectoryCaseInsensitive(string directory, out bool caseInsensitive) {
            caseInsensitive = true;
            try {
                using (SafeFileHandle handle = OpenWindowsPathHandle(directory)) {
                    if (handle.IsInvalid || !GetFileInformationByHandleEx(handle, FileCaseSensitiveInfo,
                            out FileCaseSensitiveInformation information,
                            (uint)Marshal.SizeOf<FileCaseSensitiveInformation>())) return false;
                    caseInsensitive = (information.Flags & FileCaseSensitiveDirectory) == 0;
                    return true;
                }
            } catch (DllNotFoundException) {
                return false;
            } catch (EntryPointNotFoundException) {
                return false;
            }
        }

        private static bool TryReadWindowsLinkTarget(string path, out string? target) {
#if NET8_0_OR_GREATER
            try {
                target = new FileInfo(path).LinkTarget ?? new DirectoryInfo(path).LinkTarget;
                return target != null;
            } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException) {
                throw new IOException("Could not inspect linked path '" + path + "'.", exception);
            }
#else
            target = null;
            return false;
#endif
        }

        private static bool HasWindowsReparsePoint(string path) {
            uint attributes = GetFileAttributes(path);
            if (attributes != InvalidFileAttributes) return (attributes & FileAttributeReparsePoint) != 0;
            int error = Marshal.GetLastWin32Error();
            if (error == ErrorFileNotFound || error == ErrorPathNotFound) return false;
            throw new IOException("Unable to inspect reparse-point metadata for '" + path +
                "' (OS error " + error + ").");
        }

        private static SafeFileHandle OpenWindowsPathHandle(string path) => CreateFile(
            path, 0, FileShare.Read | FileShare.Write | FileShare.Delete,
            IntPtr.Zero, OpenExisting, FileFlagBackupSemantics, IntPtr.Zero);

        private static string GetWindowsFinalPath(SafeFileHandle handle) {
            var buffer = new StringBuilder(1024);
            uint length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
            if (length >= buffer.Capacity) {
                buffer = new StringBuilder(checked((int)length + 1));
                length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
            }
            if (length == 0 || length >= buffer.Capacity) {
                throw new IOException("The physical Windows path could not be resolved (OS error " +
                    Marshal.GetLastWin32Error() + ").");
            }
            return NormalizeWindowsFinalPath(buffer.ToString());
        }

        private static string GetWindowsAuthority(string path) {
            if (path.StartsWith(@"\\?\", StringComparison.OrdinalIgnoreCase)) return string.Empty;
            if (!path.StartsWith(@"\\", StringComparison.Ordinal)) return string.Empty;
            int separator = path.IndexOf('\u005c', 2);
            if (separator < 0) return path.Substring(2).ToUpperInvariant();
            string server = path.Substring(2, separator - 2);
            if (IsLocalWindowsAuthority(server)) return string.Empty;
            return server.ToUpperInvariant();
        }

        private static bool IsLocalWindowsAuthority(string server) =>
            string.Equals(server, ".", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(server, "localhost", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(server, "127.0.0.1", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(server, "[::1]", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(server, Environment.MachineName, StringComparison.OrdinalIgnoreCase);

        internal static string NormalizeWindowsFinalPath(string path) {
            const string uncPrefix = @"\\?\UNC\";
            const string devicePrefix = @"\\?\";
            if (path.StartsWith(uncPrefix, StringComparison.OrdinalIgnoreCase)) {
                return @"\\" + path.Substring(uncPrefix.Length);
            }
            if (path.StartsWith(devicePrefix, StringComparison.OrdinalIgnoreCase) &&
                path.Length >= devicePrefix.Length + 3 &&
                path[devicePrefix.Length + 1] == ':' &&
                path[devicePrefix.Length + 2] == '\\') {
                return path.Substring(devicePrefix.Length);
            }
            return path;
        }

        private static IOException WindowsIdentityError(string path, string operation) =>
            WindowsIdentityError(path, operation, Marshal.GetLastWin32Error());

        private static IOException WindowsIdentityError(string path, string operation, int error) =>
            new IOException("Unable to " + operation + " for '" + path + "' (OS error " + error + ").");

        [StructLayout(LayoutKind.Sequential)]
        private struct FileCaseSensitiveInformation { internal uint Flags; }

        [StructLayout(LayoutKind.Sequential)]
        private struct FileId128 {
            internal ulong LowPart;
            internal ulong HighPart;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FileIdInformation {
            internal ulong VolumeSerialNumber;
            internal FileId128 FileId;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ByHandleFileInformation {
            internal uint FileAttributes;
            internal System.Runtime.InteropServices.ComTypes.FILETIME CreationTime;
            internal System.Runtime.InteropServices.ComTypes.FILETIME LastAccessTime;
            internal System.Runtime.InteropServices.ComTypes.FILETIME LastWriteTime;
            internal uint VolumeSerialNumber;
            internal uint FileSizeHigh;
            internal uint FileSizeLow;
            internal uint NumberOfLinks;
            internal uint FileIndexHigh;
            internal uint FileIndexLow;
        }

        [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string fileName, uint desiredAccess,
            FileShare shareMode, IntPtr securityAttributes, uint creationDisposition,
            uint flagsAndAttributes, IntPtr templateFile);

        [DllImport("kernel32.dll", EntryPoint = "GetFileAttributesW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFileAttributes(string fileName);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandleEx(SafeFileHandle file,
            int fileInformationClass, out FileCaseSensitiveInformation fileInformation, uint bufferSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandleEx(SafeFileHandle file,
            int fileInformationClass, out FileIdInformation fileInformation, uint bufferSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(SafeFileHandle file,
            out ByHandleFileInformation fileInformation);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint GetFileType(SafeFileHandle file);

        [DllImport("kernel32.dll", EntryPoint = "GetFinalPathNameByHandleW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFinalPathNameByHandle(SafeFileHandle file,
            StringBuilder filePath, uint filePathLength, uint flags);
    }
}
