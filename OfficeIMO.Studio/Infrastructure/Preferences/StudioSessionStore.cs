using System.Text.Json;

namespace OfficeIMO.Studio.Infrastructure.Preferences;

/// <summary>A bounded restart record. Document edits remain in the PDF recovery store.</summary>
internal sealed record StudioSessionDocument(string Path, string Fingerprint, StudioDocumentViewState View) {
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
        string temporary = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            File.WriteAllBytes(temporary, bytes);
            File.Move(temporary, _path, overwrite: true);
        } finally {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    internal void Clear() => File.Delete(_path);

    private static IReadOnlyList<StudioSessionDocument> Normalize(IEnumerable<StudioSessionDocument> documents) => documents
        .Where(document => document is not null && !string.IsNullOrWhiteSpace(document.Path) &&
            Path.IsPathFullyQualified(document.Path) && document.Path.Length <= 4096 &&
            document.LastUsedAt >= DateTimeOffset.UtcNow.AddDays(-30) && document.LastUsedAt <= DateTimeOffset.UtcNow.AddDays(1) &&
            document.Fingerprint is { Length: 64 } && document.Fingerprint.All(Uri.IsHexDigit) && document.View is not null)
        .Select(document => document with { Path = Path.GetFullPath(document.Path), View = document.View.Normalize() })
        .DistinctBy(document => document.Path, OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal)
        .Take(MaximumDocuments).ToArray();

    private static StudioSessionSnapshot Empty() => new(1, DateTimeOffset.UtcNow, null, []);
}
