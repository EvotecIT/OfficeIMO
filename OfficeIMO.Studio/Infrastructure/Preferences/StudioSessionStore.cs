using System.Text.Json;
using OfficeIMO.Internal;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Studio.Infrastructure.Preferences;

/// <summary>A bounded restart record. Document edits remain in the PDF recovery store.</summary>
internal sealed record StudioSessionDocument(string Path, string Fingerprint, StudioDocumentViewState View) {
    public StudioStorageReference? Storage { get; init; }
    public DateTimeOffset LastUsedAt { get; init; } = DateTimeOffset.UtcNow;
}
internal sealed record StudioSessionSnapshot(int SchemaVersion, DateTimeOffset UpdatedAt, string? ActivePath,
    IReadOnlyList<StudioSessionDocument> Documents);

internal sealed class StudioSessionStore(string path) {
    private const int MaximumDocuments = 32;
    private const int MaximumBytes = 512 * 1024;
    private readonly string _path = Path.GetFullPath(path);

    internal StudioSessionSnapshot Load() {
        try {
            if (!File.Exists(_path) || new FileInfo(_path).Length > MaximumBytes) return Empty();
            var snapshot = JsonSerializer.Deserialize<StudioSessionSnapshot>(File.ReadAllBytes(_path));
            if (snapshot?.SchemaVersion != 1 || snapshot.Documents is null ||
                snapshot.UpdatedAt < DateTimeOffset.UtcNow.AddDays(-30) || snapshot.UpdatedAt > DateTimeOffset.UtcNow.AddDays(1)) return Empty();
            return snapshot with { Documents = Normalize(snapshot.Documents) };
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException or ArgumentException or NotSupportedException) {
            return Empty();
        }
    }

    internal void Save(StudioSessionSnapshot snapshot) {
        var bounded = snapshot with { SchemaVersion = 1, UpdatedAt = DateTimeOffset.UtcNow, Documents = Normalize(snapshot.Documents) };
        byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(bounded);
        if (bytes.Length > MaximumBytes) throw new IOException("The session record exceeds its storage limit.");
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        OfficeFileCommit.WriteAllBytes(_path, bytes, OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly);
    }

    internal void Clear() => File.Delete(_path);

    private static IReadOnlyList<StudioSessionDocument> Normalize(IEnumerable<StudioSessionDocument> documents) {
        var result = new List<StudioSessionDocument>();
        var identities = new HashSet<string>(StringComparer.Ordinal);
        foreach (var document in documents) {
            if (document is null || string.IsNullOrWhiteSpace(document.Path) || document.Path.Length > 4096 ||
                document.LastUsedAt < DateTimeOffset.UtcNow.AddDays(-30) || document.LastUsedAt > DateTimeOffset.UtcNow.AddDays(1) ||
                document.Fingerprint is not { Length: 64 } || !document.Fingerprint.All(Uri.IsHexDigit) || document.View is null) continue;
            try {
                if (!Path.IsPathFullyQualified(document.Path) && !Uri.TryCreate(document.Path, UriKind.Absolute, out _)) continue;
                string location = OfficeStorageIdentity.Normalize(document.Path);
                string identity = OfficeStorageIdentity.GetPersistenceKey(location);
                StudioStorageReference? reference = document.Storage;
                if (reference is not null && (reference.Name is null || reference.Name.Length > 4096 ||
                    reference.Bookmark?.Length > 32768 ||
                    OfficeStorageIdentity.GetPersistenceKey(reference.Location) != identity)) continue;
                if (!identities.Add(identity)) continue;
                result.Add(document with { Path = location, View = document.View.Normalize() });
                if (result.Count == MaximumDocuments) break;
            } catch (Exception error) when (error is ArgumentException or NotSupportedException or UriFormatException) {
                // A malformed record does not hide the other recoverable documents.
            }
        }
        return result;
    }
    private static StudioSessionSnapshot Empty() => new(1, DateTimeOffset.UtcNow, null, []);
}
