using System.Text.Json;
using System.Security;

namespace OfficeIMO.Studio.Features.Home;

internal interface IRecentDocumentStore {
    IReadOnlyList<RecentDocumentViewModel> Load();

    void Save(IReadOnlyList<RecentDocumentViewModel> documents);

    void Clear();
}

internal sealed class JsonRecentDocumentStore : IRecentDocumentStore {
    private const int MaximumEntries = 12;
    private const int MaximumBytes = 512 * 1024;
    private readonly string _path;
    private readonly Func<bool> _enabled;

    public JsonRecentDocumentStore(string path, Func<bool>? enabled = null) {
        _path = Path.GetFullPath(path);
        _enabled = enabled ?? (() => true);
    }

    public static JsonRecentDocumentStore CreateDefault() =>
        new(Infrastructure.StudioDataPaths.CreateDefault().RecentDocumentsPath);

    public IReadOnlyList<RecentDocumentViewModel> Load() {
        if (!_enabled()) return [];
        try {
            if (!File.Exists(_path) || new FileInfo(_path).Length > MaximumBytes) return [];
            RecentDocumentEntry?[]? entries = JsonSerializer.Deserialize<RecentDocumentEntry?[]>(File.ReadAllText(_path));
            if (entries is null) return [];

            var documents = new List<RecentDocumentViewModel>(MaximumEntries);
            foreach (RecentDocumentEntry? entry in entries) {
                if (documents.Count == MaximumEntries) break;
                if (entry is null || string.IsNullOrWhiteSpace(entry.Path)) continue;
                try {
                    if (entry.Reference is { } reference && (reference.Name is null || reference.Name.Length > 4096 ||
                        reference.Bookmark?.Length > 32768 || OfficeIMO.Internal.OfficeStorageIdentity.Normalize(reference.Location) !=
                        OfficeIMO.Internal.OfficeStorageIdentity.Normalize(entry.Path))) continue;
                    var document = new RecentDocumentViewModel(entry.Path, entry.OpenedAt) { StorageReference = entry.Reference };
                    if (OfficeIMO.Internal.OfficeStorageIdentity.GetLocalPath(document.Path) is null ||
                        entry.Reference?.Bookmark is not null || File.Exists(document.Path)) documents.Add(document);
                } catch (Exception exception) when (exception is IOException or ArgumentException or NotSupportedException or SecurityException) {
                    // Ignore one malformed entry without discarding otherwise useful history.
                }
            }
            return documents;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException or SecurityException) {
            return [];
        }
    }

    public void Save(IReadOnlyList<RecentDocumentViewModel> documents) {
        if (!_enabled()) return;
        try {
            string? directory = Path.GetDirectoryName(_path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            RecentDocumentEntry[] entries = documents
                .Take(MaximumEntries)
                .Select(static document => new RecentDocumentEntry(document.Path, document.OpenedAt, document.StorageReference))
                .ToArray();
            string json = JsonSerializer.Serialize(entries, new JsonSerializerOptions { WriteIndented = true });
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(json);
            if (bytes.Length > MaximumBytes) throw new IOException("Recent document references exceed their storage limit.");
            OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytes(_path, bytes,
                OfficeIMO.Core.Internal.OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            // Recent history is a convenience. A read-only profile must not prevent document work.
        }
    }

    public void Clear() => File.Delete(_path);

    private sealed record RecentDocumentEntry(string Path, DateTimeOffset OpenedAt, Infrastructure.StudioStorageReference? Reference = null);
}
