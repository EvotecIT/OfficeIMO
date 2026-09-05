using OfficeIMO.Studio.Features.Home;

namespace OfficeIMO.Studio.Tests;

public sealed class RecentDocumentStoreTests {
    [Fact]
    public void StorePersistsAtMostTwelveExistingDocumentsInOrder() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-recents-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try {
            var documents = Enumerable.Range(1, 15)
                .Select(index => {
                    string path = Path.Combine(root, $"document-{index:00}.pdf");
                    File.WriteAllText(path, "fixture");
                    return new RecentDocumentViewModel(path, DateTimeOffset.UtcNow.AddMinutes(-index));
                })
                .ToArray();
            var store = new JsonRecentDocumentStore(Path.Combine(root, "recent-documents.json"));

            store.Save(documents);
            IReadOnlyList<RecentDocumentViewModel> loaded = store.Load();

            Assert.Equal(12, loaded.Count);
            Assert.Equal(documents.Take(12).Select(document => document.Path), loaded.Select(document => document.Path));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void StoreTreatsCorruptHistoryAsEmpty() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-recents-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "recent-documents.json");
        File.WriteAllText(path, "{not valid json");

        try {
            var store = new JsonRecentDocumentStore(path);

            Assert.Empty(store.Load());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void StoreSkipsSemanticallyInvalidEntriesWithoutCrashingStartup() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-recents-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "recent-documents.json");
        File.WriteAllText(path, """[null,{"Path":"\u0000","OpenedAt":"2026-09-01T00:00:00+00:00"}]""");

        try {
            var store = new JsonRecentDocumentStore(path);

            Assert.Empty(store.Load());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}
