using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.Archives;

public static partial class ArchiveReader {
    /// <summary>
    /// Validates compressed size and physical entry count on a seekable ZIP source before opening a <see cref="ZipArchive"/>.
    /// The source position is restored. This check is safe only if the caller keeps the source bytes immutable until its archive has been read.
    /// Prefer a stream or path TraverseZip overload for mutable or untrusted sources.
    /// </summary>
    public static void ValidateZipSource(Stream zipStream, ArchiveTraversalOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (zipStream == null) throw new ArgumentNullException(nameof(zipStream));
        if (!zipStream.CanRead || !zipStream.CanSeek) {
            throw new ArgumentException("ZIP source must be readable and seekable for preflight validation.", nameof(zipStream));
        }

        ArchiveTraversalOptions effective = Normalize(options);
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
        var output = ArchiveTemporaryFile.Create();
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

}
