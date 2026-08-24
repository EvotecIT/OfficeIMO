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
        private const uint StatxRequiredIdentity = 0x00000105;
        private const int ErrorNoEntry = 2;
        private const int ErrorNotDirectory = 20;
        private const int ErrorInvalidArgument = 22;
        private const uint UnixFileTypeMask = 0xf000;
        private const uint UnixDirectoryType = 0x4000;

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
                    metadata = CreateLinuxMetadata(status);
                    return true;
                }
                int error = Marshal.GetLastWin32Error();
                if (IsConfirmedMissing(error)) {
                    metadata = default(OfficeFileMetadata);
                    return false;
                }
                throw UnixIdentityError(path, error);
            } catch (EntryPointNotFoundException exception) {
                throw new PlatformNotSupportedException("This Linux runtime does not expose statx.", exception);
            }
        }

        private static OfficeFileMetadata GetLinuxMetadata(int descriptor) {
            try {
                if (LinuxStatX(descriptor, string.Empty, AtEmptyPath, StatxBasicStats, out LinuxStatx status) != 0) {
                    throw UnixIdentityError("open descriptor");
                }
                return CreateLinuxMetadata(status);
            } catch (EntryPointNotFoundException exception) {
                throw new PlatformNotSupportedException("This Linux runtime does not expose statx.", exception);
            }
        }

        private static OfficeFileMetadata CreateLinuxMetadata(LinuxStatx status) {
            if ((status.Mask & StatxRequiredIdentity) != StatxRequiredIdentity) {
                throw new IOException("Linux did not return the required file type, link count, and inode metadata.");
            }
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
