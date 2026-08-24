using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using OfficeIMO.Internal;

namespace OfficeIMO.Core.Internal {
    /// <summary>Creates owner-only, non-shareable temporary files that the operating system deletes on close.</summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal static class OfficeTemporaryFile {
        internal static FileStream Create(
            string prefix,
            string suffix,
            FileOptions options,
            out string path) {
            if (string.IsNullOrWhiteSpace(prefix)) throw new ArgumentException("Temporary file prefix cannot be empty.", nameof(prefix));
            if (suffix == null) throw new ArgumentNullException(nameof(suffix));

            path = Path.Combine(Path.GetTempPath(), prefix + Guid.NewGuid().ToString("N") + suffix);
            return CreateAtPath(path, 81920, options | FileOptions.DeleteOnClose);
        }

        internal static FileStream CreateAtPath(
            string path,
            int bufferSize,
            FileOptions options) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Temporary file path cannot be empty.", nameof(path));
            if (bufferSize < 1) throw new ArgumentOutOfRangeException(nameof(bufferSize));
#if NET6_0_OR_GREATER
            var streamOptions = new FileStreamOptions {
                Mode = FileMode.CreateNew,
                Access = FileAccess.ReadWrite,
                Share = FileShare.None,
                BufferSize = bufferSize,
                Options = options
            };
            if (!OperatingSystem.IsWindows()) {
                streamOptions.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            }
            return new FileStream(path, streamOptions);
#else
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                return CreateUnixOwnerOnly(path, bufferSize, options);
            }
            return new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.None, bufferSize, options);
#endif
        }

        internal static FileStream CreateUnixOwnerOnly(string path, int bufferSize, FileOptions options) {
            int descriptor = OpenFile(path, GetExclusiveCreateFlags(), 0x180U);
            if (descriptor < 0) {
                throw new IOException(
                    "Unable to create an owner-only temporary file (OS error "
                    + Marshal.GetLastWin32Error() + ").");
            }

            var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
            try {
                if (ChangeDescriptorMode(descriptor, 0x180U) != 0) {
                    throw new IOException(
                        "Unable to secure the owner-only temporary file (OS error "
                        + Marshal.GetLastWin32Error() + ").");
                }
                var stream = new UnixOwnerOnlyFileStream(handle, path, bufferSize, options);
                handle = null!;
                return stream;
            } catch {
                handle?.Dispose();
                TryDelete(path);
                throw;
            }
        }

        private sealed class UnixOwnerOnlyFileStream : FileStream {
            private readonly bool _deleteOnClose;
            private readonly string _path;
            private bool _deleted;

            internal UnixOwnerOnlyFileStream(
                SafeFileHandle handle,
                string path,
                int bufferSize,
                FileOptions options)
                : base(
                    handle,
                    FileAccess.ReadWrite,
                    bufferSize,
                    (options & FileOptions.Asynchronous) != 0) {
                _path = path;
                _deleteOnClose = (options & FileOptions.DeleteOnClose) != 0;
            }

            protected override void Dispose(bool disposing) {
                try {
                    base.Dispose(disposing);
                } finally {
                    if (_deleteOnClose && !_deleted) {
                        _deleted = true;
                        TryDelete(_path);
                    }
                }
            }
        }

        internal static int GetExclusiveCreateFlags() {
            const int openReadWrite = 0x0002;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                const int openCreate = 0x0200;
                const int openExclusive = 0x0800;
                const int openCloseOnExec = 0x01000000;
                return openReadWrite | openCreate | openExclusive | openCloseOnExec;
            }
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
                const int linuxOpenCreate = 0x0040;
                const int linuxOpenExclusive = 0x0080;
                const int linuxOpenCloseOnExec = 0x00080000;
                return openReadWrite | linuxOpenCreate | linuxOpenExclusive | linuxOpenCloseOnExec;
            }
            throw new PlatformNotSupportedException(
                "This Unix platform does not expose a supported exclusive-create flag layout.");
        }

        internal static void CopyUnixFileMode(string sourcePath, string destinationPath) {
#if NET6_0_OR_GREATER
            if (!OperatingSystem.IsWindows()) {
                File.SetUnixFileMode(destinationPath, File.GetUnixFileMode(sourcePath));
            }
#else
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                CopyUnixFileModePortable(sourcePath, destinationPath);
            }
#endif
        }

        internal static void ApplyDefaultUnixCreationMode(string path) {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

            string? directory = Path.GetDirectoryName(path);
            string probePath = Path.Combine(
                string.IsNullOrEmpty(directory) ? "." : directory,
                ".officeimo-mode-" + Guid.NewGuid().ToString("N") + ".tmp");
            try {
                using (new FileStream(
                    probePath,
                    FileMode.CreateNew,
                    FileAccess.ReadWrite,
                    FileShare.None)) {
                }
                CopyUnixFileMode(probePath, path);
            } finally {
                TryDelete(probePath);
            }
        }

        internal static void CopyUnixFileModePortable(string sourcePath, string destinationPath) {
            MethodInfo? getMode = typeof(File).GetMethod(
                "GetUnixFileMode",
                BindingFlags.Public | BindingFlags.Static,
                binder: null,
                types: new[] { typeof(string) },
                modifiers: null);
            Type? modeType = getMode?.ReturnType;
            MethodInfo? setMode = modeType == null ? null : typeof(File).GetMethod(
                "SetUnixFileMode",
                BindingFlags.Public | BindingFlags.Static,
                binder: null,
                types: new[] { typeof(string), modeType },
                modifiers: null);
            if (getMode != null && setMode != null) {
                object mode = getMode.Invoke(null, new object[] { sourcePath })!;
                setMode.Invoke(null, new[] { (object)destinationPath, mode });
                return;
            }

            CopyUnixFileModeNative(sourcePath, destinationPath);
        }

        internal static void CopyUnixFileModeNative(string sourcePath, string destinationPath) {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.OSX) &&
                !RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
                throw new PlatformNotSupportedException(
                    "This Unix platform does not expose a supported native file-mode layout.");
            }
            uint mode = OfficePathIdentity.GetMetadata(sourcePath).UnixMode & 0x0FFFU;
            if (ChangeFileMode(destinationPath, mode) != 0) {
                throw new IOException(
                    "Unable to preserve existing Unix file permissions (OS error "
                    + Marshal.GetLastWin32Error() + ").");
            }
        }

        private static void TryDelete(string path) {
            try { File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }

        [DllImport("libc", EntryPoint = "open", SetLastError = true)]
        private static extern int OpenFile(string path, int flags, uint mode);

        [DllImport("libc", EntryPoint = "chmod", SetLastError = true)]
        private static extern int ChangeFileMode(string path, uint mode);

        [DllImport("libc", EntryPoint = "fchmod", SetLastError = true)]
        private static extern int ChangeDescriptorMode(int descriptor, uint mode);
    }
}
