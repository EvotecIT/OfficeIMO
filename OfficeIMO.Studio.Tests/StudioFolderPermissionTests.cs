using System.Text;
using System.Text.Json;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioFolderPermissionTests {
    [Theory]
    [InlineData("macOS", "ios")]
    [InlineData("ios", "macOS")]
    public void PickerBookmarksRetainTheirNativePayloadAndRejectOtherPlatforms(string platform, string other) {
        byte[] payload = [1, 2, 3, 4];
        byte[] envelope = new byte[16 + payload.Length];
        Encoding.ASCII.GetBytes("ava.v1." + platform).CopyTo(envelope, 0);
        payload.CopyTo(envelope, 16);
        string bookmark = Convert.ToBase64String(envelope);
        Assert.Equal(payload, StudioStorageAccess.DecodeFolderBookmark(bookmark, platform));
        Assert.Throws<IOException>(() => StudioStorageAccess.DecodeFolderBookmark(bookmark, other));
        envelope[0] = 0;
        Assert.Throws<IOException>(() => StudioStorageAccess.DecodeFolderBookmark(Convert.ToBase64String(envelope), platform));
        Assert.Throws<IOException>(() => StudioStorageAccess.DecodeFolderBookmark("bad", platform));
        Assert.Throws<IOException>(() => StudioStorageAccess.DecodeFolderBookmark(Convert.ToBase64String(new byte[24017]), platform));
    }

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
