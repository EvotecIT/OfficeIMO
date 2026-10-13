using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Archives;

/// <summary>Creates passive CAB/MSI inventories and bounded whole-member reads without filesystem extraction or installer execution.</summary>
public static class ArchiveInspector {
    /// <summary>Snapshots one package into bounded memory and indexes its authored members without decoding their payloads.</summary>
    /// <exception cref="InvalidDataException">The input is malformed, inconsistent or exceeds a configured bound.</exception>
    /// <exception cref="NotSupportedException">The input uses an unsupported container or metadata profile.</exception>
    public static ArchiveInspection Open(string packagePath, ArchiveInspectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(packagePath)) throw new ArgumentException("A package path is required.", nameof(packagePath));
        using var source = new FileStream(packagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Open(source, options, cancellationToken);
    }

    /// <summary>Snapshots a readable seekable stream from its beginning, restores its original position on success or failure,
    /// and leaves it open. Later reads use the immutable snapshot and their own cancellation tokens.</summary>
    /// <exception cref="InvalidDataException">The input is malformed, inconsistent or exceeds a configured bound.</exception>
    /// <exception cref="NotSupportedException">The input uses an unsupported container or metadata profile.</exception>
    public static ArchiveInspection Open(Stream source, ArchiveInspectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead || !source.CanSeek) throw new ArgumentException("Inspection requires a readable seekable stream.", nameof(source));
        var requested = options ?? new ArchiveInspectionOptions();
        var limits = new ArchiveLimits(new ArchiveExtractionOptions {
            MaxInputBytes = requested.MaxInputBytes, MaxEntryCount = requested.MaxEntryCount,
            MaxCabinetCount = requested.MaxCabinetCount, MaxEntryBytes = requested.MaxEntryBytes,
            MaxExpandedBytes = requested.MaxExpandedBytes, MaxDecompressionWorkBytes = requested.MaxDecompressionWorkBytes,
            MaxMetadataBytes = requested.MaxMetadataBytes, MaxTableRows = requested.MaxTableRows,
            MaxStringPoolEntries = requested.MaxStringPoolEntries, MaxPathDepth = requested.MaxPathDepth,
            MaxPathLength = requested.MaxPathLength
        });
        int directoryRecords = requested.MaxDirectoryRecords;
        if (directoryRecords < 1) throw new ArgumentOutOfRangeException(nameof(requested.MaxDirectoryRecords));
        cancellationToken.ThrowIfCancellationRequested();
        long originalPosition = source.Position;
        ArchiveInspection? inspection = null;
        try {
            long length = source.Length;
            if (length < 0 || length > limits.Input || length > int.MaxValue)
                throw new InvalidDataException("Package exceeds the configured input or managed buffer limit.");
            source.Position = 0;
            var data = new byte[(int)length];
            int offset = 0;
            while (offset < data.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(data, offset, Math.Min(65536, data.Length - offset));
                if (read <= 0) throw new InvalidDataException("Package ended before its declared stream length.");
                offset += read;
            }
            cancellationToken.ThrowIfCancellationRequested();
            inspection = new ArchiveInspection(data, limits, directoryRecords, cancellationToken);
            return inspection;
        } finally {
            try { source.Position = originalPosition; }
            catch { inspection?.Dispose(); throw; }
        }
    }
}
