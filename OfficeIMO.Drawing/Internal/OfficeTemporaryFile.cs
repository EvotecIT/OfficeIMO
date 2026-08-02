using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Drawing.Internal {
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
            string? directory = Path.GetDirectoryName(path);
            string templatePath = Path.Combine(
                string.IsNullOrEmpty(directory) ? "." : directory,
                ".officeimo-XXXXXX");
            byte[] template = Encoding.UTF8.GetBytes(templatePath + "\0");
            int descriptor = CreateOwnerOnlyTemporaryFile(template);
            if (descriptor < 0) {
                throw new IOException(
                    "Unable to create an owner-only temporary file (OS error "
                    + Marshal.GetLastWin32Error() + ").");
            }

            var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
            int terminator = Array.IndexOf(template, (byte)0);
            string generatedPath = Encoding.UTF8.GetString(template, 0, terminator < 0 ? template.Length : terminator);
            bool targetLinked = false;
            try {
                if (LinkFile(generatedPath, path) != 0) {
                    throw new IOException(
                        "Unable to reserve the owner-only temporary-file path (OS error "
                        + Marshal.GetLastWin32Error() + ").");
                }
                targetLinked = true;
                if (UnlinkFile(generatedPath) != 0) {
                    throw new IOException(
                        "Unable to remove the owner-only temporary-file staging link (OS error "
                        + Marshal.GetLastWin32Error() + ").");
                }
                generatedPath = string.Empty;
                handle.Dispose();
                handle = null!;
                return new FileStream(path, FileMode.Open, FileAccess.ReadWrite,
                    FileShare.None, bufferSize, options);
            } catch {
                handle?.Dispose();
                if (targetLinked) TryDelete(path);
                if (generatedPath.Length > 0) TryDelete(generatedPath);
                throw;
            }
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
            int modeOffset;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
                modeOffset = 4;
            } else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) {
                modeOffset = RuntimeInformation.OSArchitecture == Architecture.X64 ? 24 : 16;
            } else {
                throw new PlatformNotSupportedException(
                    "This Unix platform does not expose a supported native file-mode layout.");
            }

            IntPtr statBuffer = Marshal.AllocHGlobal(512);
            try {
                if (StatFile(sourcePath, statBuffer) != 0) {
                    throw new IOException(
                        "Unable to inspect existing Unix file permissions (OS error "
                        + Marshal.GetLastWin32Error() + ").");
                }

                uint mode = unchecked((uint)Marshal.ReadInt32(statBuffer, modeOffset)) & 0x0FFFU;
                if (ChangeFileMode(destinationPath, mode) != 0) {
                    throw new IOException(
                        "Unable to preserve existing Unix file permissions (OS error "
                        + Marshal.GetLastWin32Error() + ").");
                }
            } finally {
                Marshal.FreeHGlobal(statBuffer);
            }
        }

        private static void TryDelete(string path) {
            try { File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }

        [DllImport("libc", EntryPoint = "mkstemp", SetLastError = true)]
        private static extern int CreateOwnerOnlyTemporaryFile([In, Out] byte[] template);

        [DllImport("libc", EntryPoint = "link", SetLastError = true)]
        private static extern int LinkFile(string existingPath, string newPath);

        [DllImport("libc", EntryPoint = "unlink", SetLastError = true)]
        private static extern int UnlinkFile(string path);

        [DllImport("libc", EntryPoint = "stat", SetLastError = true)]
        private static extern int StatFile(string path, IntPtr buffer);

        [DllImport("libc", EntryPoint = "chmod", SetLastError = true)]
        private static extern int ChangeFileMode(string path, uint mode);
    }
}
