using System.Security.Cryptography;
using System.Buffers.Binary;
using System.Text;
using System.Text.Json;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspaceRecoveryStore {
    internal const long MaximumSnapshotBytes = 512L * 1024 * 1024;
    internal static readonly TimeSpan Retention = TimeSpan.FromDays(30);
    private const int MaximumMetadataBytes = 64 * 1024;
    private static ReadOnlySpan<byte> SnapshotMagic => "OIMORCV2"u8;
    private readonly string _root;
    private readonly SemaphoreSlim _persistenceGate = new(1, 1);
    private bool _persistenceEnabled;

    internal PdfWorkspaceRecoveryStore(string? root = null, bool persistenceEnabled = true) {
        _persistenceEnabled = persistenceEnabled;
        _root = root ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OfficeIMO",
            "Studio",
            "Recovery");
    }

    /// <summary>Commits the preference after active writes finish. Failure retains the previous policy.</summary>
    internal async Task SetPersistenceAsync(bool enabled, Action persistPreference, CancellationToken token = default) {
        ArgumentNullException.ThrowIfNull(persistPreference);
        await _persistenceGate.WaitAsync(token);
        try {
            token.ThrowIfCancellationRequested();
            persistPreference();
            _persistenceEnabled = enabled;
        } finally {
            _persistenceGate.Release();
        }
    }

    /// <summary>Stores edits when enabled; returns null without touching storage when opted out.</summary>
    internal async Task<string?> WriteAsync(
        string sourcePath,
        string baseFingerprint,
        byte[] bytes,
        long revision,
        CancellationToken cancellationToken) {
        await _persistenceGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (!_persistenceEnabled) return null;
            return await WriteEnabledAsync(sourcePath, baseFingerprint, bytes, revision, cancellationToken).ConfigureAwait(false);
        } finally {
            _persistenceGate.Release();
        }
    }

    private async Task<string> WriteEnabledAsync(
        string sourcePath,
        string baseFingerprint,
        byte[] bytes,
        long revision,
        CancellationToken cancellationToken) {
        if (bytes.LongLength > MaximumSnapshotBytes) {
            throw new IOException("The document exceeds the 512 MiB recovery snapshot limit.");
        }
        string canonicalPath = Canonicalize(sourcePath);
        string key = CreateKey(canonicalPath);
        using FileStream lease = await AcquireExclusiveAsync(cancellationToken).ConfigureAwait(false);
        string snapshotPath = Path.Combine(_root, key + ".recovery");
        string recoveryFingerprint = Fingerprint(bytes);

        byte[] metadata = JsonSerializer.SerializeToUtf8Bytes(new RecoveryMetadata(
            canonicalPath,
            baseFingerprint,
            recoveryFingerprint,
            revision,
            DateTimeOffset.UtcNow) { SchemaVersion = 2 });
        if (metadata.Length > MaximumMetadataBytes) throw new InvalidDataException("Recovery metadata exceeds its size limit.");
        await WriteSnapshotAsync(snapshotPath, metadata, bytes, cancellationToken).ConfigureAwait(false);
        // Migrate only after the complete replacement has been published successfully.
        TryDelete(Path.Combine(_root, key + ".pdf"));
        TryDelete(Path.Combine(_root, key + ".json"));
        return snapshotPath;
    }

    internal async Task DeleteAsync(string sourcePath, CancellationToken token = default) {
        token.ThrowIfCancellationRequested();
        if (!Directory.Exists(_root)) return;
        using FileStream lease = await AcquireExclusiveAsync(token).ConfigureAwait(false);
        string key = CreateKey(Canonicalize(sourcePath));
        // Remove legacy data first so a failed deletion cannot expose older edits.
        File.Delete(Path.Combine(_root, key + ".pdf"));
        File.Delete(Path.Combine(_root, key + ".json"));
        File.Delete(Path.Combine(_root, key + ".recovery"));
    }

    internal string? Find(string sourcePath, string baseFingerprint) {
        return ReadSnapshot(sourcePath, baseFingerprint)?.Path;
    }

    internal byte[]? ReadVerifiedSnapshot(string sourcePath, string baseFingerprint) =>
        ReadSnapshot(sourcePath, baseFingerprint)?.Bytes;

    private (string Path, byte[] Bytes)? ReadSnapshot(string sourcePath, string baseFingerprint) {
        string canonicalPath = Canonicalize(sourcePath);
        string key = CreateKey(canonicalPath);
        string pdfPath = Path.Combine(_root, key + ".pdf");
        string metadataPath = Path.Combine(_root, key + ".json");
        string snapshotPath = Path.Combine(_root, key + ".recovery");

        try {
            if (File.Exists(snapshotPath)) {
                // Never fall back to older edits when a newer snapshot exists but is invalid.
                byte[]? snapshot = ReadCurrentSnapshot(snapshotPath, canonicalPath, baseFingerprint);
                return snapshot is null ? null : (snapshotPath, snapshot);
            }
            if (!File.Exists(pdfPath) || !File.Exists(metadataPath)) return null;
            byte[]? metadataBytes = ReadBounded(metadataPath, MaximumMetadataBytes);
            if (metadataBytes is null) return null;
            RecoveryMetadata? metadata = JsonSerializer.Deserialize<RecoveryMetadata>(metadataBytes);
            if (!IsValidMetadata(metadata, canonicalPath, baseFingerprint, schemaVersion: 1)) return null;
            byte[]? bytes = ReadBounded(pdfPath, MaximumSnapshotBytes);
            if (bytes is null) return null;
            return string.Equals(Fingerprint(bytes), metadata!.RecoveryFingerprint, StringComparison.OrdinalIgnoreCase) ? (pdfPath, bytes) : null;
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            return null;
        }
    }

    internal static string Fingerprint(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));

    private static bool IsValidMetadata(RecoveryMetadata? metadata, string path, string fingerprint, int schemaVersion) {
        DateTimeOffset now = DateTimeOffset.UtcNow;
        return metadata is not null && metadata.SchemaVersion == schemaVersion &&
            metadata.UpdatedAt >= now - Retention && metadata.UpdatedAt <= now.AddDays(1) &&
            metadata.Revision >= 0 && PathsEqual(path, metadata.SourcePath) &&
            string.Equals(fingerprint, metadata.BaseFingerprint, StringComparison.OrdinalIgnoreCase);
    }

    private static byte[]? ReadCurrentSnapshot(string path, string sourcePath, string baseFingerprint) {
        // Delete sharing allows atomic replacement while this reader retains the old file handle.
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
        long length = stream.Length;
        if (length < 20 || length > 20 + MaximumMetadataBytes + MaximumSnapshotBytes) return null;
        Span<byte> header = stackalloc byte[20];
        stream.ReadExactly(header);
        if (!header[..8].SequenceEqual(SnapshotMagic)) return null;
        int metadataLength = BinaryPrimitives.ReadInt32LittleEndian(header[8..12]);
        long pdfLength = BinaryPrimitives.ReadInt64LittleEndian(header[12..]);
        if (metadataLength <= 0 || metadataLength > MaximumMetadataBytes ||
            pdfLength <= 0 || pdfLength > MaximumSnapshotBytes || length != 20L + metadataLength + pdfLength) return null;
        byte[] metadataBytes = new byte[metadataLength];
        stream.ReadExactly(metadataBytes);
        RecoveryMetadata? metadata = JsonSerializer.Deserialize<RecoveryMetadata>(metadataBytes);
        if (!IsValidMetadata(metadata, sourcePath, baseFingerprint, schemaVersion: 2)) return null;
        byte[] bytes = new byte[checked((int)pdfLength)];
        stream.ReadExactly(bytes);
        return stream.ReadByte() == -1 &&
            string.Equals(Fingerprint(bytes), metadata!.RecoveryFingerprint, StringComparison.OrdinalIgnoreCase) ? bytes : null;
    }

    private static byte[]? ReadBounded(string path, long maximumBytes) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        long length = stream.Length;
        if (length <= 0 || length > maximumBytes) return null;
        var bytes = new byte[checked((int)length)];
        stream.ReadExactly(bytes);
        // Also reject a file that grew through an already-open writer on platforms
        // where sharing flags cannot exclude that writer.
        return stream.ReadByte() == -1 ? bytes : null;
    }

    private static async Task WriteSnapshotAsync(string path, byte[] metadata, byte[] bytes, CancellationToken cancellationToken) {
        string temporaryPath = string.Empty;
        try {
            byte[] header = new byte[20];
            SnapshotMagic.CopyTo(header);
            BinaryPrimitives.WriteInt32LittleEndian(header.AsSpan(8, 4), metadata.Length);
            BinaryPrimitives.WriteInt64LittleEndian(header.AsSpan(12, 8), bytes.LongLength);
            using (var stream = OfficeFileCommit.CreateTemporaryFile(path, FileOptions.Asynchronous, out temporaryPath, 64 * 1024)) {
                await stream.WriteAsync(header, cancellationToken).ConfigureAwait(false);
                await stream.WriteAsync(metadata, cancellationToken).ConfigureAwait(false);
                await stream.WriteAsync(bytes, cancellationToken).ConfigureAwait(false);
                await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                stream.Flush(flushToDisk: true);
            }
            cancellationToken.ThrowIfCancellationRequested();
            OfficeFileCommit.CommitTemporaryFileAtomically(temporaryPath, path,
                OfficeFileCommit.ConflictPolicy.Replace, OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly);
        } finally {
            TryDelete(temporaryPath);
        }
    }

    private static string Canonicalize(string sourcePath) => OfficeStorageIdentity.Normalize(sourcePath);

    private static string CreateKey(string canonicalPath) {
        string identity = OfficeStorageIdentity.GetPersistenceKey(canonicalPath);
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(identity));
        return Convert.ToHexString(hash.AsSpan(0, 16)).ToLowerInvariant();
    }

    private static bool PathsEqual(string left, string right) => string.Equals(
        OfficeStorageIdentity.GetPersistenceKey(left),
        OfficeStorageIdentity.GetPersistenceKey(right),
        StringComparison.Ordinal);

    private static void TryDelete(string path) {
        try {
            if (File.Exists(path)) File.Delete(path);
        } catch (IOException) {
            // Recovery cleanup is best effort; a future successful save retries it.
        } catch (UnauthorizedAccessException) {
            // Recovery cleanup is best effort; a future successful save retries it.
        }
    }

    private sealed record RecoveryMetadata(
        string SourcePath,
        string BaseFingerprint,
        string RecoveryFingerprint,
        long Revision,
        DateTimeOffset UpdatedAt) {
        public int SchemaVersion { get; init; } = 1;
    }
}
