using System;
using System.IO;

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
#if NET6_0_OR_GREATER
            var streamOptions = new FileStreamOptions {
                Mode = FileMode.CreateNew,
                Access = FileAccess.ReadWrite,
                Share = FileShare.None,
                BufferSize = 81920,
                Options = options | FileOptions.DeleteOnClose
            };
            if (!OperatingSystem.IsWindows()) {
                streamOptions.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            }
            return new FileStream(path, streamOptions);
#else
            return new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.None, 81920, options | FileOptions.DeleteOnClose);
#endif
        }
    }
}
