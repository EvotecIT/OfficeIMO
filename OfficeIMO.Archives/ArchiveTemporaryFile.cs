using System;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;

namespace OfficeIMO.Archives;

internal static class ArchiveTemporaryFile {
    internal static Stream Create() {
        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-archive-" + Guid.NewGuid().ToString("N"));
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            Directory.CreateDirectory(directory);
        } else if (CreatePrivateDirectoryUnix(directory, 0x1C0) != 0) {
            throw new IOException("Unable to create private ZIP snapshot directory (OS error " +
                Marshal.GetLastWin32Error().ToString(CultureInfo.InvariantCulture) + ").");
        }
        PrivateSnapshotFileStream output;
        try {
            output = new PrivateSnapshotFileStream(Path.Combine(directory, "source.tmp"), directory);
        } catch {
            TryDeleteSnapshotDirectory(directory);
            throw;
        }
        return output;
    }

    private static void TryDeleteSnapshotDirectory(string directory) {
        try {
            Directory.Delete(directory, recursive: true);
        } catch (IOException) {
            // DeleteOnClose removes the source; directory cleanup is best effort.
        } catch (UnauthorizedAccessException) {
            // DeleteOnClose removes the source; directory cleanup is best effort.
        }
    }

    private sealed class PrivateSnapshotFileStream : FileStream {
        private readonly string _directory;

        internal PrivateSnapshotFileStream(string path, string directory) : base(path,
            FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920,
            FileOptions.DeleteOnClose | FileOptions.SequentialScan) {
            _directory = directory;
        }

        protected override void Dispose(bool disposing) {
            try {
                base.Dispose(disposing);
            } finally {
                if (disposing) TryDeleteSnapshotDirectory(_directory);
            }
        }
    }

    [DllImport("libc", EntryPoint = "mkdir", SetLastError = true)]
    private static extern int CreatePrivateDirectoryUnix(string path, uint mode);
}
