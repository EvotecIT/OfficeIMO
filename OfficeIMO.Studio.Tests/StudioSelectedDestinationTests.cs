using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSelectedDestinationTests {
    [Fact]
    public async Task SelectingAndCancellingANewDestinationDoesNotCreateIt() {
        using var storage = new StudioStorageAccess();
        var folder = new StudioProviderOutputFolderTests.OutputFolder(hierarchical: true);
        string parent = (await storage.RegisterFolderAsync([folder.Item], default))!;
        string destination = await storage.SelectFileDestinationAsync(parent, "output.pdf", default);
        Assert.Equal(0, folder.Creations);
        await Assert.ThrowsAsync<OperationCanceledException>(() => storage.PublishAsync(destination, [1, 2, 3], null,
            _ => throw new OperationCanceledException(), default));
        Assert.Equal(0, folder.Creations);
        Assert.Empty(folder.Files);
    }

    [Fact]
    public async Task AFileAppearingAfterSelectionIsNotTruncated() {
        using var storage = new StudioStorageAccess();
        var folder = new StudioProviderOutputFolderTests.OutputFolder(hierarchical: true);
        string parent = (await storage.RegisterFolderAsync([folder.Item], default))!;
        string destination = await storage.SelectFileDestinationAsync(parent, "output.pdf", default);
        var appeared = new TestStorageFile(destination, [7, 8, 9], "output.pdf");
        folder.Files[appeared.Name] = appeared;
        await Assert.ThrowsAsync<IOException>(() => storage.PublishAsync(destination, [1, 2, 3], null,
            _ => Task.CompletedTask, default));
        Assert.Equal(new byte[] { 7, 8, 9 }, appeared.Bytes);
        Assert.Equal(0, appeared.Writes);
        Assert.Equal(0, folder.Creations);
        // Explicit reselection grants replacement of the file the user can now see.
        string selectedAgain = await storage.SelectFileDestinationAsync(parent, "output.pdf", default);
        Assert.Equal(destination, selectedAgain);
        await storage.PublishAsync(selectedAgain, [1, 2, 3], null, _ => Task.CompletedTask, default);
        Assert.Equal(new byte[] { 1, 2, 3 }, appeared.Bytes);
        Assert.Equal(1, appeared.Writes);
    }
}
