using Microsoft.Win32.SafeHandles;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.GoogleWorkspace.Drive {
    internal static class GoogleDriveDownloadFileGuard {
        internal static FileStream CreateNew(string path, int bufferSize) =>
            Open(path, createNew: true, bufferSize);

        internal static FileStream OpenExisting(string path, int bufferSize) =>
            Open(path, createNew: false, bufferSize);

        internal static void EnsurePathReferencesHandle(string path, FileStream stream) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                if (!GetFileInformationByHandle(stream.SafeFileHandle, out ByHandleFileInformation information) ||
                    information.NumberOfLinks != 1) {
                    throw new IOException("The guarded download destination must have exactly one filesystem link.");
                }
                string actual = GetWindowsFinalPath(stream.SafeFileHandle);
                string expected = Path.GetFullPath(path);
                if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase)) {
                    throw new IOException("The guarded download destination path no longer references the opened file.");
                }
                return;
            }

            if (GetUnixFileStatus(stream.SafeFileHandle.DangerousGetHandle(), out UnixFileStatus opened) != 0 ||
                opened.HardLinkCount != 1) {
                throw new IOException("The guarded download destination identity could not be verified.");
            }
            int descriptor = OpenUnixNoFollow(path, createNew: false);
            if (descriptor < 0) {
                throw new IOException("The guarded download destination path no longer references a regular file.");
            }
            try {
                if (GetUnixFileStatus(new IntPtr(descriptor), out UnixFileStatus current) != 0 ||
                    current.HardLinkCount != 1 ||
                    opened.Device != current.Device || opened.Inode != current.Inode) {
                    throw new IOException("The guarded download destination was replaced during transfer.");
                }
            } finally {
                CloseUnix(descriptor);
            }
        }

        private static FileStream Open(string path, bool createNew, int bufferSize) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                var stream = new FileStream(path, createNew ? FileMode.CreateNew : FileMode.Open,
                    FileAccess.ReadWrite, FileShare.Read, bufferSize, FileOptions.SequentialScan);
                try {
                    EnsurePathReferencesHandle(path, stream);
                    return stream;
                } catch {
                    stream.Dispose();
                    throw;
                }
            }

            int descriptor = OpenUnixNoFollow(path, createNew);
            if (descriptor < 0) {
                throw new IOException(createNew
                    ? "A new guarded download will not overwrite an existing or linked destination."
                    : "The checkpointed download destination is not a regular unlinked file.");
            }
            if (GetUnixFileStatus(new IntPtr(descriptor), out UnixFileStatus status) != 0 ||
                (status.Mode & 0xF000) != 0x8000 || status.HardLinkCount != 1) {
                CloseUnix(descriptor);
                throw new IOException("The guarded download destination is not a regular file.");
            }

            var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
            try {
                var stream = new FileStream(handle, FileAccess.ReadWrite, bufferSize, isAsync: false);
                EnsurePathReferencesHandle(path, stream);
                return stream;
            } catch {
                handle.Dispose();
                throw;
            }
        }

        private static int OpenUnixNoFollow(string path, bool createNew) {
            int flags;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                const int readWrite = 0x0002;
                const int nonBlocking = 0x0004;
                const int create = 0x0200;
                const int exclusive = 0x0800;
                const int noFollow = 0x0100;
                const int closeOnExec = 0x01000000;
                flags = readWrite | nonBlocking | noFollow | closeOnExec;
                if (createNew) flags |= create | exclusive;
            } else {
                const int readWrite = 0x0002;
                const int create = 0x0040;
                const int exclusive = 0x0080;
                const int nonBlocking = 0x0800;
                const int noFollow = 0x00020000;
                const int closeOnExec = 0x00080000;
                flags = readWrite | nonBlocking | noFollow | closeOnExec;
                if (createNew) flags |= create | exclusive;
            }
            return OpenUnix(path, flags, 384); // 0600 for a newly created destination.
        }

        private static string GetWindowsFinalPath(SafeFileHandle handle) {
            var buffer = new StringBuilder(1024);
            uint length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
            if (length == 0) throw new IOException("The guarded download destination path could not be resolved.");
            if (length >= buffer.Capacity) {
                buffer = new StringBuilder(checked((int)length + 1));
                length = GetFinalPathNameByHandle(handle, buffer, (uint)buffer.Capacity, 0);
                if (length == 0 || length >= buffer.Capacity) {
                    throw new IOException("The guarded download destination path could not be resolved.");
                }
            }
            string value = buffer.ToString();
            const string uncPrefix = @"\\?\UNC\";
            const string devicePrefix = @"\\?\";
            if (value.StartsWith(uncPrefix, StringComparison.OrdinalIgnoreCase)) {
                return @"\\" + value.Substring(uncPrefix.Length);
            }
            return value.StartsWith(devicePrefix, StringComparison.OrdinalIgnoreCase)
                ? value.Substring(devicePrefix.Length)
                : value;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint GetFinalPathNameByHandle(SafeFileHandle file,
            StringBuilder filePath, uint filePathLength, uint flags);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(SafeFileHandle file,
            out ByHandleFileInformation information);

        [DllImport("libc", EntryPoint = "open", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int OpenUnix(string path, int flags, int mode);

        [DllImport("libc", EntryPoint = "close", SetLastError = true)]
        private static extern int CloseUnix(int descriptor);

        [DllImport("System.Native", EntryPoint = "SystemNative_FStat", SetLastError = true)]
        private static extern int GetUnixFileStatus(IntPtr descriptor, out UnixFileStatus status);

        [StructLayout(LayoutKind.Sequential)]
        private struct ByHandleFileInformation {
            internal uint FileAttributes;
            internal FileTime CreationTime;
            internal FileTime LastAccessTime;
            internal FileTime LastWriteTime;
            internal uint VolumeSerialNumber;
            internal uint FileSizeHigh;
            internal uint FileSizeLow;
            internal uint NumberOfLinks;
            internal uint FileIndexHigh;
            internal uint FileIndexLow;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FileTime {
            internal uint LowDateTime;
            internal uint HighDateTime;
        }

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
    }
}
