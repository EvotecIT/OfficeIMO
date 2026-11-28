using System.Text.Json;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioFolderPermissionTests {
    [Fact]
    public async Task SelectingAChildDoesNotCreateItAndTheFolderGrantSurvivesRestart() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-folder-grant-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            using var storage = new StudioStorageAccess();
            string folder = (await storage.RegisterFolderAsync([TestLocalOutputFolder.Create(root)], default))!;
            string destination = storage.SelectLocalFileDestination(folder, "new.pdf");
            Assert.False(File.Exists(destination));
            Assert.False(storage.UsesProviderPublication(destination));
            Assert.Empty(Directory.EnumerateFileSystemEntries(root));
            StudioStorageReference reference = storage.Describe(destination);
            Assert.Equal(root, reference.PermissionFolder);
            reference = JsonSerializer.Deserialize<StudioStorageReference>(JsonSerializer.Serialize(reference))!;
            File.WriteAllBytes(destination, [1, 2, 3]);
            using var restarted = new StudioStorageAccess();
            restarted.Remember(reference);
            Assert.Equal(new byte[] { 1, 2, 3 }, (await restarted.ReadSnapshotAsync(destination, default)).Bytes);
            string outside = Path.Combine(Path.GetDirectoryName(root)!, "outside.pdf");
            restarted.Remember(reference with { Location = outside });
            await Assert.ThrowsAsync<IOException>(() => restarted.ReadSnapshotAsync(outside, default));
        } finally { Directory.Delete(root, true); }
    }
}
