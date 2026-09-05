using System.Text.Json;
using System.Security;

namespace OfficeIMO.Studio.Features.Home;

internal interface IRecentDocumentStore {
    IReadOnlyList<RecentDocumentViewModel> Load();

    void Save(IReadOnlyList<RecentDocumentViewModel> documents);
}

internal sealed class JsonRecentDocumentStore : IRecentDocumentStore {
    private const int MaximumEntries = 12;
    private readonly string _path;

    public JsonRecentDocumentStore(string path) => _path = Path.GetFullPath(path);

    public static JsonRecentDocumentStore CreateDefault() =>
        new(Infrastructure.StudioDataPaths.CreateDefault().RecentDocumentsPath);

    public IReadOnlyList<RecentDocumentViewModel> Load() {
        try {
            if (!File.Exists(_path)) return [];
            RecentDocumentEntry?[]? entries = JsonSerializer.Deserialize<RecentDocumentEntry?[]>(File.ReadAllText(_path));
            if (entries is null) return [];

            var documents = new List<RecentDocumentViewModel>(MaximumEntries);
            foreach (RecentDocumentEntry? entry in entries) {
                if (documents.Count == MaximumEntries) break;
                if (entry is null || string.IsNullOrWhiteSpace(entry.Path)) continue;
                try {
                    var document = new RecentDocumentViewModel(entry.Path, entry.OpenedAt);
                    if (File.Exists(document.Path)) documents.Add(document);
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
        try {
            string? directory = Path.GetDirectoryName(_path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            RecentDocumentEntry[] entries = documents
                .Take(MaximumEntries)
                .Select(static document => new RecentDocumentEntry(document.Path, document.OpenedAt))
                .ToArray();
            string json = JsonSerializer.Serialize(entries, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_path, json);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            // Recent history is a convenience. A read-only profile must not prevent document work.
        }
    }

    private sealed record RecentDocumentEntry(string Path, DateTimeOffset OpenedAt);
}
