using OfficeIMO.Core.Internal;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeIMO.Zip;

public static partial class ZipTraversal {
    /// <summary>
    /// Validates compressed size and physical entry count on a seekable ZIP source before opening a <see cref="ZipArchive"/>.
    /// The source position is restored. This check is safe only if the caller keeps the source bytes immutable until its archive has been read.
    /// Prefer a stream or path Traverse overload for mutable or untrusted sources.
    /// </summary>
    public static void ValidateSource(Stream zipStream, ZipTraversalOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (zipStream == null) throw new ArgumentNullException(nameof(zipStream));
        if (!zipStream.CanRead || !zipStream.CanSeek) {
            throw new ArgumentException("ZIP source must be readable and seekable for preflight validation.", nameof(zipStream));
        }

        ZipTraversalOptions effective = Normalize(options);
        cancellationToken.ThrowIfCancellationRequested();
        long originalPosition = zipStream.Position;
        try {
            long archiveLength = zipStream.Length;
            if (archiveLength > effective.MaxArchiveBytes) {
                throw new InvalidDataException($"ZIP source exceeds MaxArchiveBytes ({effective.MaxArchiveBytes}).");
            }

            zipStream.Position = 0;
            OfficeArchiveSafety.ZipCentralDirectoryScanResult scan =
                OfficeArchiveSafety.ScanZipCentralDirectory(zipStream, archiveLength,
                    effective.MaxPhysicalEntries, cancellationToken);
            if (!scan.IsValid) {
                throw new InvalidDataException(scan.Error ?? "ZIP central directory is malformed.");
            }
            if (scan.LimitExceeded) {
                throw new InvalidDataException(
                    $"ZIP source exceeds MaxPhysicalEntries ({effective.MaxPhysicalEntries}).");
            }
        } finally {
            zipStream.Position = originalPosition;
        }
    }

    private static Stream CreateBoundedSnapshot(Stream source, long maximumBytes,
        CancellationToken cancellationToken) {
        if (!source.CanSeek) return SnapshotCurrentPosition(source, maximumBytes, cancellationToken);

        long originalPosition = source.Position;
        try {
            if (source.Length > maximumBytes) {
                throw new InvalidDataException($"ZIP source exceeds MaxArchiveBytes ({maximumBytes}).");
            }
            source.Position = 0;
            return SnapshotCurrentPosition(source, maximumBytes, cancellationToken);
        } finally {
            source.Position = originalPosition;
        }
    }

    private static Stream SnapshotCurrentPosition(Stream source, long maximumBytes,
        CancellationToken cancellationToken) {
        string directory = Path.Combine(Path.GetTempPath(),
            "officeimo-zip-" + Guid.NewGuid().ToString("N"));
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
        try {
            var buffer = new byte[81920];
            long total = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                if (total > maximumBytes - read) {
                    throw new InvalidDataException($"ZIP source exceeds MaxArchiveBytes ({maximumBytes}).");
                }
                output.Write(buffer, 0, read);
                total += read;
            }
            output.Position = 0;
            return output;
        } catch {
            output.Dispose();
            throw;
        }
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
