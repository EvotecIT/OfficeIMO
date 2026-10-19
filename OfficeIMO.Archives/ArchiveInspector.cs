using System;
using System.IO;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives;

/// <summary>Creates passive ZIP/CAB/MSI inventories and bounded whole-member reads without filesystem extraction or installer execution.</summary>
public static class ArchiveInspector {
    private const long MaximumInMemorySnapshotBytes = 64L * 1024 * 1024;
    /// <summary>Snapshots one package and indexes its authored members without decoding their payloads. ZIP input above 64 MiB uses private temporary storage.</summary>
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
        var limits = CreateLimits(requested);
        int directoryRecords = requested.MaxDirectoryRecords;
        cancellationToken.ThrowIfCancellationRequested();
        if (!source.CanSeek) return SnapshotNonseekable(source, limits, directoryRecords, cancellationToken);
        long originalPosition = source.Position;
        ArchiveInspection? inspection = null;
        try {
            long length = source.Length;
            ArchiveLimits.Check(length, limits.Input, nameof(requested.MaxInputBytes));
            limits.Operation.AdmitInput(length);
            source.Position = 0;
            ArchiveFormat format = ArchiveExtractor.Detect(source, cancellationToken);
            source.Position = 0;
            if (format == ArchiveFormat.Zip && length > MaximumInMemorySnapshotBytes) {
                Stream snapshot = ArchiveTemporaryFile.Create();
                try {
                    CopySnapshot(source, snapshot, length, cancellationToken);
                    snapshot.Position = 0;
                    inspection = new ArchiveInspection(snapshot, limits, directoryRecords, cancellationToken);
                    return inspection;
                } catch { snapshot.Dispose(); throw; }
            }
            ArchiveLimits.Check(length, Math.Min(limits.Input, int.MaxValue), nameof(requested.MaxInputBytes));
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

    /// <summary>Indexes an existing immutable ZIP snapshot without copying it. By default inspection takes ownership;
    /// the caller disposes it on admission failure. With leaveOpen, the caller must retain the snapshot unchanged
    /// throughout inspection; a non-owning view preserves its position and lifetime, including on failure.</summary>
    internal static ArchiveInspection OpenZipSnapshot(Stream snapshot, ArchiveInspectionOptions options, CancellationToken token,
        bool leaveOpen = false) {
        var limits = CreateLimits(options);
        ArchiveLimits.Check(snapshot.Length, limits.Input, nameof(options.MaxInputBytes));
        limits.Operation.AdmitInput(snapshot.Length);
        Stream input = leaveOpen ? new SeekableReadWindowStream(snapshot, 0, snapshot.Length, preserveSourcePosition: true) : snapshot;
        return new ArchiveInspection(input, limits, options.MaxDirectoryRecords, token);
    }

    private static ArchiveLimits CreateLimits(ArchiveInspectionOptions requested) {
        var limits = new ArchiveLimits(new ArchiveExtractionOptions {
            OperationBudget = requested.OperationBudget, ZipLegacyCodePage = requested.ZipLegacyCodePage,
            MaxInputBytes = requested.MaxInputBytes, MaxEntryCount = requested.MaxEntryCount,
            MaxCabinetCount = requested.MaxCabinetCount, MaxEntryBytes = requested.MaxEntryBytes,
            MaxExpandedBytes = requested.MaxExpandedBytes, MaxDecompressionWorkBytes = requested.MaxDecompressionWorkBytes,
            MaxMetadataBytes = requested.MaxMetadataBytes, MaxTableRows = requested.MaxTableRows,
            MaxStringPoolEntries = requested.MaxStringPoolEntries, MaxPathDepth = requested.MaxPathDepth,
            MaxPathLength = requested.MaxPathLength
        });
        if (requested.MaxDirectoryRecords < 1) throw new ArgumentOutOfRangeException(nameof(requested.MaxDirectoryRecords));
        if (requested.MaxTemporaryBytes < 0) throw new ArgumentOutOfRangeException(nameof(requested.MaxTemporaryBytes));
        limits.TemporaryBytes = requested.MaxTemporaryBytes;
        return limits;
    }

    private static ArchiveInspection SnapshotNonseekable(Stream source, ArchiveLimits limits, int directoryRecords,
        CancellationToken token) {
        long maximum = limits.Input;
        Stream? snapshot = new MemoryStream();
        try {
            var buffer = new byte[65536];
            // Admit even an empty container so nested depth remains authoritative.
            limits.Operation.AdmitInput(0);
            while (true) {
                token.ThrowIfCancellationRequested();
                long remaining = maximum - snapshot.Length;
                int count = remaining >= buffer.Length ? buffer.Length : (int)remaining + 1;
                int read = source.Read(buffer, 0, count);
                if (read == 0) break;
                if (read > remaining) throw ArchiveLimits.Exceeded(nameof(ArchiveInspectionOptions.MaxInputBytes), maximum);
                limits.Operation.AdmitInput(read);
                if (snapshot is MemoryStream && snapshot.Length + read > MaximumInMemorySnapshotBytes) {
                    Stream file = ArchiveTemporaryFile.Create();
                    try {
                        snapshot.Position = 0;
                        CopySnapshot(snapshot, file, snapshot.Length, token);
                    } catch { file.Dispose(); throw; }
                    snapshot.Dispose(); snapshot = file;
                }
                snapshot.Write(buffer, 0, read);
            }
            token.ThrowIfCancellationRequested();
            snapshot.Position = 0;
            if (ArchiveExtractor.Detect(snapshot, token) == ArchiveFormat.Zip) {
                var inspection = new ArchiveInspection(snapshot, limits, directoryRecords, token);
                snapshot = null;
                return inspection;
            }
            ArchiveLimits.Check(snapshot.Length, Math.Min(limits.Input, int.MaxValue), nameof(ArchiveInspectionOptions.MaxInputBytes));
            snapshot.Position = 0;
            using var bytes = new MemoryStream((int)snapshot.Length);
            CopySnapshot(snapshot, bytes, snapshot.Length, token);
            return new ArchiveInspection(bytes.GetBuffer(), limits, directoryRecords, token);
        } finally { snapshot?.Dispose(); }
    }

    private static void CopySnapshot(Stream source, Stream destination, long length, CancellationToken token) {
        var buffer = new byte[65536];
        long remaining = length;
        while (remaining > 0) {
            token.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
            if (read == 0) throw new InvalidDataException("Package ended before its declared stream length.");
            destination.Write(buffer, 0, read); remaining -= read;
        }
        token.ThrowIfCancellationRequested();
    }
}
