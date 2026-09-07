using System.Buffers.Binary;
using System.Text.Json;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed record RecoveryCleanupResult(int RemovedFiles, long RemovedBytes, int FailedFiles);

internal sealed partial class PdfWorkspaceRecoveryStore {
    internal event EventHandler? MaintenanceCompleted;

    /// <summary>Removes expired known snapshots and abandoned staging files without racing writers.</summary>
    internal Task<RecoveryCleanupResult> CleanupExpiredAsync(CancellationToken token = default) =>
        CleanupAsync(clearAll: false, token);

    /// <summary>Removes recognized recovery data. The coordination file and unrelated files are retained.</summary>
    internal Task<RecoveryCleanupResult> ClearAllAsync(CancellationToken token = default) =>
        CleanupAsync(clearAll: true, token);

    private async Task<RecoveryCleanupResult> CleanupAsync(bool clearAll, CancellationToken token) {
        if (!Directory.Exists(_root)) return new(0, 0, 0);
        FileStream lease = await AcquireExclusiveAsync(token).ConfigureAwait(false);
        try {
            using (lease) return await Task.Run(() => CleanupUnderLease(clearAll, token), token).ConfigureAwait(false);
        } finally {
            MaintenanceCompleted?.Invoke(this, EventArgs.Empty);
        }
    }

    internal bool HasSnapshotFiles(string sourcePath) {
        string key = CreateKey(Canonicalize(sourcePath));
        return File.Exists(Path.Combine(_root, key + ".recovery")) ||
            (File.Exists(Path.Combine(_root, key + ".pdf")) && File.Exists(Path.Combine(_root, key + ".json")));
    }

    private RecoveryCleanupResult CleanupUnderLease(bool clearAll, CancellationToken token) {
        int removed = 0;
        var failedPaths = new HashSet<string>(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
        long bytes = 0;
        DateTimeOffset now = DateTimeOffset.UtcNow;
        foreach (string path in Directory.EnumerateFiles(_root, "*", SearchOption.TopDirectoryOnly)) {
            token.ThrowIfCancellationRequested();
            string name = Path.GetFileName(path);
            bool record = IsRecordName(name);
            bool staging = IsStagingName(name);
            if (!record && !staging) continue;
            try {
                if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) { failedPaths.Add(path); continue; }
                bool remove = clearAll || (staging
                    ? File.GetLastWriteTimeUtc(path) < now.UtcDateTime.AddDays(-1)
                    : IsExpiredRecord(path, now));
                if (!remove) continue;
                if (Path.GetExtension(path) == ".json") {
                    DeleteDataFile(Path.ChangeExtension(path, ".pdf"));
                }
                DeleteDataFile(path);
            } catch (FileNotFoundException) {
                // Another actor may already have removed an obsolete legacy file.
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                failedPaths.Add(path);
            }
        }
        return new(removed, bytes, failedPaths.Count);

        void DeleteDataFile(string path) {
            try {
                if (!File.Exists(path)) return;
                if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) { failedPaths.Add(path); return; }
                long length = new FileInfo(path).Length;
                File.Delete(path);
                removed++;
                bytes += length;
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                failedPaths.Add(path);
            }
        }
    }

    private static bool IsExpiredRecord(string path, DateTimeOffset now) {
        string extension = Path.GetExtension(path);
        if (extension == ".pdf") {
            string metadataPath = Path.ChangeExtension(path, ".json");
            if (File.Exists(metadataPath)) return IsExpiredLegacyMetadata(metadataPath, now);
            return File.GetLastWriteTimeUtc(path) < (now - Retention).UtcDateTime;
        }
        if (extension == ".json") return IsExpiredLegacyMetadata(path, now);

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        Span<byte> header = stackalloc byte[20];
        if (stream.Length < header.Length) return IsOldMalformed(path, now);
        stream.ReadExactly(header);
        // Preserve unknown future formats rather than deleting data this version cannot interpret.
        if (!header[..8].SequenceEqual(SnapshotMagic)) return false;
        int metadataLength = BinaryPrimitives.ReadInt32LittleEndian(header[8..12]);
        if (metadataLength <= 0 || metadataLength > MaximumMetadataBytes || stream.Length < 20L + metadataLength) {
            return IsOldMalformed(path, now);
        }
        byte[] metadata = new byte[metadataLength];
        stream.ReadExactly(metadata);
        return IsExpiredMetadata(metadata, schemaVersion: 2, path, now);
    }

    private static bool IsExpiredLegacyMetadata(string path, DateTimeOffset now) {
        if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) return false;
        byte[]? metadata = ReadBounded(path, MaximumMetadataBytes);
        return metadata is null ? IsOldMalformed(path, now) : IsExpiredMetadata(metadata, schemaVersion: 1, path, now);
    }

    private static bool IsExpiredMetadata(byte[] bytes, int schemaVersion, string path, DateTimeOffset now) {
        try {
            RecoveryMetadata? metadata = JsonSerializer.Deserialize<RecoveryMetadata>(bytes);
            if (metadata is null) return IsOldMalformed(path, now);
            if (metadata.SchemaVersion != schemaVersion) return false;
            if (metadata.UpdatedAt == default || metadata.UpdatedAt > now.AddDays(1) || metadata.Revision < 0 ||
                string.IsNullOrWhiteSpace(metadata.SourcePath) ||
                !IsFingerprint(metadata.BaseFingerprint) || !IsFingerprint(metadata.RecoveryFingerprint)) {
                return IsOldMalformed(path, now);
            }
            return metadata.UpdatedAt < now - Retention;
        } catch (JsonException) {
            return IsOldMalformed(path, now);
        }
    }

    private static bool IsOldMalformed(string path, DateTimeOffset now) =>
        File.GetLastWriteTimeUtc(path) < (now - Retention).UtcDateTime;

    private static bool IsRecordName(string name) => name.Length > 33 && IsHexKey(name.AsSpan(0, 32)) &&
        name[32..] is ".recovery" or ".pdf" or ".json";

    private static bool IsStagingName(string name) {
        const string prefix = ".officeimo-";
        if (name.StartsWith(prefix, StringComparison.Ordinal) && name.EndsWith(".tmp", StringComparison.Ordinal)) {
            return IsHexKey(name.AsSpan(prefix.Length, name.Length - prefix.Length - 4));
        }
        int marker = name.IndexOf(".tmp-", StringComparison.Ordinal);
        return marker > 0 && IsRecordName(name[..marker]) && IsHexKey(name.AsSpan(marker + 5));
    }

    private static bool IsHexKey(ReadOnlySpan<char> value) {
        if (value.Length != 32) return false;
        foreach (char character in value) if (!(character is >= '0' and <= '9' or >= 'a' and <= 'f')) return false;
        return true;
    }

    private static bool IsFingerprint(string? value) => value is { Length: 64 } && value.All(Uri.IsHexDigit);
}
