using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Archives;

/// <summary>Creates passive ZIP/CAB/MSI inventories and bounded whole-member reads without filesystem extraction or installer execution.</summary>
public static class ArchiveInspector {
    /// <summary>Snapshots one package into bounded memory and indexes its authored members without decoding their payloads.</summary>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured or fixed managed admission ceiling is exhausted.</exception>
    /// <exception cref="InvalidDataException">The input is malformed or inconsistent.</exception>
    /// <exception cref="NotSupportedException">The input uses an unsupported container or metadata profile.</exception>
    public static ArchiveInspection Open(string packagePath, ArchiveInspectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(packagePath)) throw new ArgumentException("A package path is required.", nameof(packagePath));
        using var source = new FileStream(packagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Open(source, options, cancellationToken);
    }

    /// <summary>Snapshots a readable seekable stream from its beginning, restores its original position on success or failure,
    /// and leaves it open. Nonseekable streams are consumed from their current position through EOF under the input limit.
    /// Later reads use the immutable snapshot and their own cancellation tokens.</summary>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured or fixed managed admission ceiling is exhausted.</exception>
    /// <exception cref="InvalidDataException">The input is malformed or inconsistent.</exception>
    /// <exception cref="NotSupportedException">The input uses an unsupported container or metadata profile.</exception>
    public static ArchiveInspection Open(Stream source, ArchiveInspectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("Inspection requires a readable stream.", nameof(source));
        var requested = options ?? new ArchiveInspectionOptions();
        var limits = new ArchiveLimits(new ArchiveExtractionOptions {
            OperationBudget = requested.OperationBudget, ZipLegacyCodePage = requested.ZipLegacyCodePage,
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
        if (!source.CanSeek) return SnapshotNonseekable(source, limits, directoryRecords, cancellationToken);
        long originalPosition = source.Position;
        ArchiveInspection? inspection = null;
        try {
            long length = source.Length;
            ArchiveLimits.Check(length, Math.Min(limits.Input, int.MaxValue), nameof(requested.MaxInputBytes));
            limits.Operation.AdmitInput(length);
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

    private static ArchiveInspection SnapshotNonseekable(Stream source, ArchiveLimits limits, int directoryRecords,
        CancellationToken token) {
        long maximum = Math.Min(limits.Input, int.MaxValue);
        using var snapshot = new MemoryStream();
        var buffer = new byte[65536];
        // Admit even an empty container so nested depth remains authoritative.
        limits.Operation.AdmitInput(0);
        while (true) {
            token.ThrowIfCancellationRequested();
            long remaining = maximum - snapshot.Length;
            int read = source.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining + 1));
            if (read == 0) break;
            ArchiveLimits.Check(read, remaining, nameof(ArchiveInspectionOptions.MaxInputBytes));
            limits.Operation.AdmitInput(read);
            snapshot.Write(buffer, 0, read);
        }
        token.ThrowIfCancellationRequested();
        return new ArchiveInspection(snapshot.ToArray(), limits, directoryRecords, token);
    }
}
