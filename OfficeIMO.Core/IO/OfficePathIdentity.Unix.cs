using Microsoft.Win32.SafeHandles;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.Internal {
    internal static partial class OfficePathIdentity {
        private const int AtFdcwd = -100;
        private const int AtEmptyPath = 0x1000;
        private const uint StatxBasicStats = 0x000007ff;
        private const uint StatxRequiredIdentity = 0x00000107;
        private const int ErrorNoEntry = 2;
        private const int ErrorPermissionDenied = 1;
        private const int ErrorAccessDenied = 13;
        private const int ErrorNotDirectory = 20;
        private const int ErrorInvalidArgument = 22;
        private const int ErrorFunctionNotImplemented = 38;
        private const int ErrorOperationNotSupported = 95;
        private const int MacPathConfCaseSensitive = 11;
        private const int LinuxOpenReadOnly = 0;
        private const int LinuxOpenDirectory = 0x10000;
        private const int LinuxOpenCloseOnExec = 0x80000;
        private const int LinuxCaseFoldFlag = 0x40000000;
        private const ulong LinuxGetFileFlags64 = 0x80086601;
        private const ulong LinuxGetFileFlags32 = 0x80046601;
        private const ulong LinuxExtFileSystem = 0x0000ef53;
        private const ulong LinuxF2fsFileSystem = 0xf2f52010;
        private const ulong LinuxBcachefsFileSystem = 0xca451a4e;
        private const ulong LinuxBtrfsFileSystem = 0x9123683e;
        private const ulong LinuxXfsFileSystem = 0x58465342;
        private const ulong LinuxTmpfsFileSystem = 0x01021994;
        private const ulong LinuxOverlayFileSystem = 0x794c7630;
        private const ulong LinuxMsDosFileSystem = 0x00004d44;
        private const ulong LinuxExFatFileSystem = 0x2011bab0;
        private const uint UnixFileTypeMask = 0xf000;
        private const uint UnixDirectoryType = 0x4000;

        private static FileStream OpenUnixRegularFileForRead(string path, int bufferSize) {
            int nonBlocking = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? 0x0004 : 0x0800;
            int closeOnExec = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? 0x01000000 : 0x00080000;
            int noFollow = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? 0x00000100 : 0x00020000;
            int descriptor = LinuxOpen(path, nonBlocking | closeOnExec | noFollow);
            if (descriptor < 0) throw UnixIdentityError(path);

            var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
            try {
                if (!GetUnixMetadata(handle).IsRegularFile) {
                    throw new InvalidDataException("The filesystem entry is not a regular file.");
                }
                return new FileStream(handle, FileAccess.Read, bufferSize, isAsync: false);
            } catch {
                handle.Dispose();
                throw;
            }
        }

        private static string ResolveUnixExistingPath(string path) {
            IntPtr pointer = IntPtr.Zero;
            try {
                pointer = RealPath(path, IntPtr.Zero);
                if (pointer == IntPtr.Zero) {
                    throw new IOException("Unable to resolve physical path '" + path + "' (OS error " +
                        Marshal.GetLastWin32Error() + ").");
                }
                string? value = Marshal.PtrToStringAnsi(pointer);
                if (string.IsNullOrEmpty(value)) throw new IOException("The resolved physical path was empty.");
                return value!;
            } catch (DllNotFoundException exception) {
                throw new PlatformNotSupportedException("The Unix runtime does not expose realpath.", exception);
            } catch (EntryPointNotFoundException exception) {
                throw new PlatformNotSupportedException("The Unix runtime does not expose realpath.", exception);
            } finally {
                if (pointer != IntPtr.Zero) Free(pointer);
            }
        }

        private static bool TryGetUnixMetadata(string path, out OfficeFileMetadata metadata) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) return TryGetLinuxMetadata(path, out metadata);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return TryGetMacMetadata(path, out metadata);
            throw new PlatformNotSupportedException("Unix file identity is not supported on this platform.");
        }

        private static OfficeFileMetadata GetUnixMetadata(SafeFileHandle handle) {
            int descriptor = checked((int)handle.DangerousGetHandle().ToInt64());
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) return GetLinuxMetadata(descriptor);
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return GetMacMetadata(descriptor);
            throw new PlatformNotSupportedException("Unix file identity is not supported on this platform.");
        }

        private static bool TryGetLinuxMetadata(string path, out OfficeFileMetadata metadata) {
            try {
                if (LinuxStatX(AtFdcwd, path, 0, StatxBasicStats, out LinuxStatx status) == 0) {
                    if (!HasRequiredLinuxStatxMetadata(status.Mask)) {
                        return TryGetLinuxLegacyMetadata(path, out metadata);
                    }
                    metadata = CreateLinuxMetadata(status);
                    return true;
                }
                int error = Marshal.GetLastWin32Error();
                if (ShouldFallbackFromLinuxStatx(error)) {
                    return TryGetLinuxLegacyMetadata(path, out metadata);
                }
                if (IsConfirmedMissing(error)) {
                    metadata = default(OfficeFileMetadata);
                    return false;
                }
                throw UnixIdentityError(path, error);
            } catch (EntryPointNotFoundException) {
                return TryGetLinuxLegacyMetadata(path, out metadata);
            }
        }

        private static OfficeFileMetadata GetLinuxMetadata(int descriptor) {
            try {
                if (LinuxStatX(descriptor, string.Empty, AtEmptyPath, StatxBasicStats, out LinuxStatx status) != 0) {
                    int error = Marshal.GetLastWin32Error();
                    if (ShouldFallbackFromLinuxStatx(error)) return GetLinuxLegacyMetadata(descriptor);
                    throw UnixIdentityError("open descriptor", error);
                }
                if (!HasRequiredLinuxStatxMetadata(status.Mask)) {
                    return GetLinuxLegacyMetadata(descriptor);
                }
                return CreateLinuxMetadata(status);
            } catch (EntryPointNotFoundException) {
                return GetLinuxLegacyMetadata(descriptor);
            }
        }

        internal static bool ShouldFallbackFromLinuxStatx(int error) =>
            error == ErrorPermissionDenied ||
            error == ErrorAccessDenied ||
            error == ErrorInvalidArgument ||
            error == ErrorFunctionNotImplemented ||
            error == ErrorOperationNotSupported;

        internal static bool HasRequiredLinuxStatxMetadata(uint mask) =>
            (mask & StatxRequiredIdentity) == StatxRequiredIdentity;

        private static bool TryGetLinuxLegacyMetadata(string path, out OfficeFileMetadata metadata) {
            IntPtr buffer = Marshal.AllocHGlobal(256);
            try {
                ZeroBuffer(buffer, 256);
                if (LinuxStat(path, buffer) == 0) {
                    metadata = ReadLinuxLegacyMetadata(buffer);
                    return true;
                }
                int error = Marshal.GetLastWin32Error();
                if (IsConfirmedMissing(error)) {
                    metadata = default(OfficeFileMetadata);
                    return false;
                }
                throw UnixIdentityError(path, error);
            } finally {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private static OfficeFileMetadata GetLinuxLegacyMetadata(int descriptor) {
            IntPtr buffer = Marshal.AllocHGlobal(256);
            try {
                ZeroBuffer(buffer, 256);
                if (LinuxFStat(descriptor, buffer) != 0) throw UnixIdentityError("open descriptor");
                return ReadLinuxLegacyMetadata(buffer);
            } finally {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private static OfficeFileMetadata ReadLinuxLegacyMetadata(IntPtr buffer) {
            ulong encodedDevice;
            ulong inode;
            ulong linkCount;
            uint mode;
            if (RuntimeInformation.ProcessArchitecture == Architecture.X64) {
                encodedDevice = unchecked((ulong)Marshal.ReadInt64(buffer, 0));
                inode = unchecked((ulong)Marshal.ReadInt64(buffer, 8));
                linkCount = unchecked((ulong)Marshal.ReadInt64(buffer, 16));
                mode = unchecked((uint)Marshal.ReadInt32(buffer, 24));
            } else if (RuntimeInformation.ProcessArchitecture == Architecture.Arm64) {
                encodedDevice = unchecked((ulong)Marshal.ReadInt64(buffer, 0));
                inode = unchecked((ulong)Marshal.ReadInt64(buffer, 8));
                mode = unchecked((uint)Marshal.ReadInt32(buffer, 16));
                linkCount = unchecked((uint)Marshal.ReadInt32(buffer, 20));
            } else {
                throw new PlatformNotSupportedException(
                    "The Linux stat fallback supports x64 and ARM64 process ABIs.");
            }
            ulong device = ((ulong)LinuxDeviceMajor(encodedDevice) << 32) | LinuxDeviceMinor(encodedDevice);
            return new OfficeFileMetadata(new OfficePhysicalFileIdentity(string.Empty, device, inode),
                linkCount, mode, (mode & UnixFileTypeMask) == UnixDirectoryType);
        }

        private static uint LinuxDeviceMajor(ulong device) =>
            unchecked((uint)(((device & 0x00000000000fff00UL) >> 8) |
                             ((device & 0xfffff00000000000UL) >> 32)));

        private static uint LinuxDeviceMinor(ulong device) =>
            unchecked((uint)((device & 0x00000000000000ffUL) |
                             ((device & 0x00000ffffff00000UL) >> 12)));

        private static OfficeFileMetadata CreateLinuxMetadata(LinuxStatx status) {
            ulong device = ((ulong)status.DeviceMajor << 32) | status.DeviceMinor;
            var identity = new OfficePhysicalFileIdentity(string.Empty, device, status.Inode);
            return new OfficeFileMetadata(identity, status.HardLinkCount, status.Mode,
                (status.Mode & UnixFileTypeMask) == UnixDirectoryType);
        }

        private static bool TryGetMacMetadata(string path, out OfficeFileMetadata metadata) {
            IntPtr buffer = Marshal.AllocHGlobal(256);
            try {
                ZeroBuffer(buffer, 256);
                int result = RuntimeInformation.ProcessArchitecture == Architecture.X64
                    ? MacStatInode64(path, buffer)
                    : RuntimeInformation.ProcessArchitecture == Architecture.Arm64
                        ? MacStat(path, buffer)
                        : throw new PlatformNotSupportedException("macOS file identity supports x64 and ARM64.");
                if (result == 0) {
                    metadata = ReadMacMetadata(buffer);
                    return true;
                }
                int error = Marshal.GetLastWin32Error();
                if (IsConfirmedMissing(error)) {
                    metadata = default(OfficeFileMetadata);
                    return false;
                }
                throw UnixIdentityError(path, error);
            } finally {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private static OfficeFileMetadata GetMacMetadata(int descriptor) {
            IntPtr buffer = Marshal.AllocHGlobal(256);
            try {
                ZeroBuffer(buffer, 256);
                int result = RuntimeInformation.ProcessArchitecture == Architecture.X64
                    ? MacFStatInode64(descriptor, buffer)
                    : RuntimeInformation.ProcessArchitecture == Architecture.Arm64
                        ? MacFStat(descriptor, buffer)
                        : throw new PlatformNotSupportedException("macOS file identity supports x64 and ARM64.");
                if (result != 0) throw UnixIdentityError("open descriptor");
                return ReadMacMetadata(buffer);
            } finally {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private static OfficeFileMetadata ReadMacMetadata(IntPtr buffer) {
            uint device = unchecked((uint)Marshal.ReadInt32(buffer, 0));
            uint mode = unchecked((ushort)Marshal.ReadInt16(buffer, 4));
            ulong linkCount = unchecked((ushort)Marshal.ReadInt16(buffer, 6));
            ulong inode = unchecked((ulong)Marshal.ReadInt64(buffer, 8));
            return new OfficeFileMetadata(
                new OfficePhysicalFileIdentity(string.Empty, device, inode), linkCount, mode,
                (mode & UnixFileTypeMask) == UnixDirectoryType);
        }

        private static bool TryGetUnixDirectoryCaseInsensitive(string directory, out bool caseInsensitive) {
            caseInsensitive = true;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                long result = MacPathConf(directory, MacPathConfCaseSensitive);
                if (result == 0 || result == 1) {
                    caseInsensitive = result == 0;
                    return true;
                }
                return false;
            }
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) return false;

            int descriptor = LinuxOpen(directory,
                LinuxOpenReadOnly | LinuxOpenDirectory | LinuxOpenCloseOnExec);
            if (descriptor < 0) return false;
            try {
                ulong request = IntPtr.Size == 8 ? LinuxGetFileFlags64 : LinuxGetFileFlags32;
                int flagsResult = LinuxIoctl(descriptor, request, out int flags);
                if (!CanClassifyLinuxCaseBehavior(flagsResult)) return false;
                if ((flags & LinuxCaseFoldFlag) != 0) {
                    caseInsensitive = true;
                    return true;
                }
                return TryGetLinuxFileSystemCaseBehavior(directory, out caseInsensitive);
            } finally {
                LinuxClose(descriptor);
            }
        }

        internal static bool CanClassifyLinuxCaseBehavior(int fileFlagsResult) => fileFlagsResult == 0;

        private static bool TryGetLinuxFileSystemCaseBehavior(string directory, out bool caseInsensitive) {
            caseInsensitive = true;
            IntPtr buffer = Marshal.AllocHGlobal(256);
            try {
                ZeroBuffer(buffer, 256);
                if (LinuxStatFs(directory, buffer) != 0) return false;
                ulong fileSystem = IntPtr.Size == 8
                    ? unchecked((ulong)Marshal.ReadInt64(buffer, 0))
                    : unchecked((uint)Marshal.ReadInt32(buffer, 0));
                return TryClassifyLinuxFileSystemCaseBehavior(fileSystem, out caseInsensitive);
            } finally {
                Marshal.FreeHGlobal(buffer);
            }
        }

        internal static bool TryClassifyLinuxFileSystemCaseBehavior(ulong fileSystem,
            out bool caseInsensitive) {
            if (fileSystem == LinuxMsDosFileSystem || fileSystem == LinuxExFatFileSystem) {
                caseInsensitive = true;
                return true;
            }
            if (fileSystem == LinuxExtFileSystem || fileSystem == LinuxF2fsFileSystem ||
                fileSystem == LinuxBcachefsFileSystem || fileSystem == LinuxBtrfsFileSystem ||
                fileSystem == LinuxXfsFileSystem || fileSystem == LinuxTmpfsFileSystem ||
                fileSystem == LinuxOverlayFileSystem) {
                caseInsensitive = false;
                return true;
            }
            caseInsensitive = false;
            return false;
        }

        private static bool TryReadUnixLinkTarget(string path, out string? target) {
            byte[] buffer = new byte[32768];
            IntPtr length = ReadLink(path, buffer, (UIntPtr)buffer.Length);
            long value = length.ToInt64();
            if (value < 0) {
                int error = Marshal.GetLastWin32Error();
                target = null;
                if (error == ErrorNoEntry || error == ErrorNotDirectory || error == ErrorInvalidArgument) return false;
                throw new IOException("Unable to inspect symbolic-link metadata for '" + path +
                    "' (OS error " + error + ").");
            }
            if (value >= buffer.Length) throw new IOException("The symbolic-link target is too long.");
            target = Encoding.UTF8.GetString(buffer, 0, checked((int)value));
            return true;
        }

        private static bool IsUnixLink(string path) => TryReadUnixLinkTarget(path, out _);

        private static void ZeroBuffer(IntPtr buffer, int length) {
            for (int index = 0; index < length; index++) Marshal.WriteByte(buffer, index, 0);
        }

        private static bool IsConfirmedMissing(int error) =>
            error == ErrorNoEntry || error == ErrorNotDirectory;

        private static IOException UnixIdentityError(string path) => UnixIdentityError(path, Marshal.GetLastWin32Error());

        private static IOException UnixIdentityError(string path, int error) =>
            new IOException("Unable to inspect physical file identity for '" + path + "' (OS error " + error + ").");

        [StructLayout(LayoutKind.Explicit, Size = 256)]
        private struct LinuxStatx {
            [FieldOffset(0)] internal uint Mask;
            [FieldOffset(16)] internal uint HardLinkCount;
            [FieldOffset(28)] internal ushort Mode;
            [FieldOffset(32)] internal ulong Inode;
            [FieldOffset(136)] internal uint DeviceMajor;
            [FieldOffset(140)] internal uint DeviceMinor;
        }

        [DllImport("libc", EntryPoint = "realpath", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr RealPath(string path, IntPtr resolvedPath);

        [DllImport("libc", EntryPoint = "free")]
        private static extern void Free(IntPtr pointer);

        [DllImport("libc", EntryPoint = "readlink", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern IntPtr ReadLink(string path, byte[] buffer, UIntPtr bufferSize);

        [DllImport("libc", EntryPoint = "statx", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern int LinuxStatX(int directoryDescriptor, string path, int flags,
            uint mask, out LinuxStatx status);

        [DllImport("libc", EntryPoint = "stat", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern int LinuxStat(string path, IntPtr buffer);

        [DllImport("libc", EntryPoint = "fstat", SetLastError = true)]
        private static extern int LinuxFStat(int descriptor, IntPtr buffer);

        [DllImport("libc", EntryPoint = "open", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern int LinuxOpen(string path, int flags);

        [DllImport("libc", EntryPoint = "ioctl", SetLastError = true)]
        private static extern int LinuxIoctl(int descriptor, ulong request, out int flags);

        [DllImport("libc", EntryPoint = "close", SetLastError = true)]
        private static extern int LinuxClose(int descriptor);

        [DllImport("libc", EntryPoint = "statfs", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern int LinuxStatFs(string path, IntPtr buffer);

        [DllImport("libc", EntryPoint = "pathconf", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern long MacPathConf(string path, int name);

        [DllImport("libc", EntryPoint = "stat$INODE64", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern int MacStatInode64(string path, IntPtr buffer);

        [DllImport("libc", EntryPoint = "stat", CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern int MacStat(string path, IntPtr buffer);

        [DllImport("libc", EntryPoint = "fstat$INODE64", SetLastError = true)]
        private static extern int MacFStatInode64(int descriptor, IntPtr buffer);

        [DllImport("libc", EntryPoint = "fstat", SetLastError = true)]
        private static extern int MacFStat(int descriptor, IntPtr buffer);
    }
}
