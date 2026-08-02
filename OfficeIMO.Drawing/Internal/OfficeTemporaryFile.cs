using System;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>Creates owner-only, non-shareable temporary files that the operating system deletes on close.</summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal static class OfficeTemporaryFile {
        private const uint OwnerFileMode = 0x180; // 0600

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
            var stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.None, bufferSize, options);
            try {
                if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                    && ChangeMode(path, OwnerFileMode) != 0) {
                    throw new IOException(
                        "Unable to restrict temporary-file permissions (OS error "
                        + Marshal.GetLastWin32Error() + ").");
                }
                return stream;
            } catch {
                stream.Dispose();
                TryDelete(path);
                throw;
            }
#endif
        }

        private static void TryDelete(string path) {
            try { File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }

        [DllImport("libc", EntryPoint = "chmod", SetLastError = true)]
        private static extern int ChangeMode(string path, uint mode);
    }
}
