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

            IntPtr openedDescriptor = stream.SafeFileHandle.DangerousGetHandle();
            if (GetUnixFileStatus(openedDescriptor, out UnixFileStatus opened) != 0 ||
                GetUnixHardLinkCount(openedDescriptor) != 1) {
                throw new IOException("The guarded download destination identity could not be verified.");
            }
            int descriptor = OpenUnixNoFollow(path, createNew: false);
            if (descriptor < 0) {
                throw new IOException("The guarded download destination path no longer references a regular file.");
            }
            try {
                if (GetUnixFileStatus(new IntPtr(descriptor), out UnixFileStatus current) != 0 ||
                    GetUnixHardLinkCount(new IntPtr(descriptor)) != 1 ||
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
            var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
            FileStream? unixStream = null;
            try {
                IntPtr openedDescriptor = handle.DangerousGetHandle();
                if (GetUnixFileStatus(openedDescriptor, out UnixFileStatus status) != 0 ||
                    (status.Mode & 0xF000) != 0x8000 || GetUnixHardLinkCount(openedDescriptor) != 1) {
                    throw new IOException("The guarded download destination is not a regular file.");
                }

                unixStream = new FileStream(handle, FileAccess.ReadWrite, bufferSize, isAsync: false);
                EnsurePathReferencesHandle(path, unixStream);
                return unixStream;
            } catch {
                unixStream?.Dispose();
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

        private static uint GetUnixHardLinkCount(IntPtr descriptor) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                if (MacFStat(descriptor, out MacFileStatus status) != 0) {
                    throw new IOException("The guarded download destination link count could not be inspected.");
                }
                return status.HardLinkCount;
            }

            try {
                const int atEmptyPath = 0x1000;
                const uint statxBasicStats = 0x000007ff;
                if (LinuxStatX(descriptor, string.Empty, atEmptyPath, statxBasicStats,
                        out LinuxFileStatus status) != 0) {
                    throw new IOException("The guarded download destination link count could not be inspected.");
                }
                return status.HardLinkCount;
            } catch (EntryPointNotFoundException exception) {
                throw new IOException("This Unix runtime cannot inspect guarded download hard links safely.", exception);
            }
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

        [DllImport("libc", EntryPoint = "statx", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int LinuxStatX(IntPtr directoryDescriptor, string path, int flags,
            uint mask, out LinuxFileStatus status);

        [DllImport("libc", EntryPoint = "fstat", SetLastError = true)]
        private static extern int MacFStat(IntPtr descriptor, out MacFileStatus status);

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
        }

        [StructLayout(LayoutKind.Explicit, Size = 256)]
        private struct LinuxFileStatus {
            [FieldOffset(16)] internal uint HardLinkCount;
        }

        [StructLayout(LayoutKind.Explicit, Size = 144)]
        private struct MacFileStatus {
            [FieldOffset(6)] internal ushort HardLinkCount;
        }
    }
}
