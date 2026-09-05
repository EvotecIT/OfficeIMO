using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed class PdfWorkspaceRecoveryStore {
    internal const long MaximumSnapshotBytes = 512L * 1024 * 1024;
    internal static readonly TimeSpan Retention = TimeSpan.FromDays(30);
    private readonly string _root;

    internal PdfWorkspaceRecoveryStore(string? root = null) {
        _root = root ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OfficeIMO",
            "Studio",
            "Recovery");
    }

    internal async Task<string> WriteAsync(
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
        Directory.CreateDirectory(_root);
        string pdfPath = Path.Combine(_root, key + ".pdf");
        string metadataPath = Path.Combine(_root, key + ".json");
        string recoveryFingerprint = Fingerprint(bytes);
        await WriteAtomicAsync(pdfPath, bytes, cancellationToken).ConfigureAwait(false);

        byte[] metadata = JsonSerializer.SerializeToUtf8Bytes(new RecoveryMetadata(
            canonicalPath,
            baseFingerprint,
            recoveryFingerprint,
            revision,
            DateTimeOffset.UtcNow));
        await WriteAtomicAsync(metadataPath, metadata, cancellationToken).ConfigureAwait(false);
        return pdfPath;
    }

    internal void Delete(string sourcePath) {
        string key = CreateKey(Canonicalize(sourcePath));
        TryDelete(Path.Combine(_root, key + ".pdf"));
        TryDelete(Path.Combine(_root, key + ".json"));
    }

    internal string? Find(string sourcePath, string baseFingerprint) {
        if (ReadVerifiedSnapshot(sourcePath, baseFingerprint) is null) return null;
        return Path.Combine(_root, CreateKey(Canonicalize(sourcePath)) + ".pdf");
    }

    internal byte[]? ReadVerifiedSnapshot(string sourcePath, string baseFingerprint) {
        string canonicalPath = Canonicalize(sourcePath);
        string key = CreateKey(canonicalPath);
        string pdfPath = Path.Combine(_root, key + ".pdf");
        string metadataPath = Path.Combine(_root, key + ".json");
        if (!File.Exists(pdfPath) || !File.Exists(metadataPath)) return null;

        try {
            byte[]? metadataBytes = ReadBounded(metadataPath, 64 * 1024);
            if (metadataBytes is null) return null;
            RecoveryMetadata? metadata = JsonSerializer.Deserialize<RecoveryMetadata>(metadataBytes);
            DateTimeOffset now = DateTimeOffset.UtcNow;
            if (metadata is null ||
                metadata.SchemaVersion != 1 ||
                metadata.UpdatedAt < now - Retention ||
                metadata.UpdatedAt > now.AddDays(1) ||
                metadata.Revision < 0 ||
                !PathsEqual(canonicalPath, metadata.SourcePath) ||
                !string.Equals(baseFingerprint, metadata.BaseFingerprint, StringComparison.OrdinalIgnoreCase)) {
                return null;
            }
            byte[]? bytes = ReadBounded(pdfPath, MaximumSnapshotBytes);
            if (bytes is null) return null;
            return string.Equals(Fingerprint(bytes), metadata.RecoveryFingerprint, StringComparison.OrdinalIgnoreCase) ? bytes : null;
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            return null;
        }
    }

    internal static string Fingerprint(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));

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

    private static async Task WriteAtomicAsync(string path, byte[] bytes, CancellationToken cancellationToken) {
        string temporaryPath = path + ".tmp-" + Guid.NewGuid().ToString("N");
        try {
            await File.WriteAllBytesAsync(temporaryPath, bytes, cancellationToken).ConfigureAwait(false);
            File.Move(temporaryPath, path, overwrite: true);
        } finally {
            TryDelete(temporaryPath);
        }
    }

    private static string Canonicalize(string sourcePath) => Path.GetFullPath(sourcePath);

    private static string CreateKey(string canonicalPath) {
        string identity = OperatingSystem.IsWindows() ? canonicalPath.ToUpperInvariant() : canonicalPath;
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(identity));
        return Convert.ToHexString(hash.AsSpan(0, 16)).ToLowerInvariant();
    }

    private static bool PathsEqual(string left, string right) => string.Equals(
        Canonicalize(left),
        Canonicalize(right),
        OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);

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
