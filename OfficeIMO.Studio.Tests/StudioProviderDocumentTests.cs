using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioProviderDocumentTests {
    [Fact]
    public async Task ImportProviderBatchIncludesAllPagesAndHasOneUndoStep() {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://documents/source", CreatePdf());
        var first = new TestStorageFile("content://documents/import-one", CreatePdf(2));
        var second = new TestStorageFile("content://documents/import-two", CreatePdf(3));
        string location = await storage.RegisterAsync(source.Item, default);
        var imports = await storage.RegisterManyAsync([first.Item, second.Item], default);
        using var root = new TestDirectory();
        using var workspace = await PdfWorkspace.OpenAsync(location, default, new(root.Path), storage: storage);
        Assert.Equal(5, await workspace.ImportAsync(imports, 1, default));
        Assert.Equal(6, workspace.Pages.Count);
        Assert.True(workspace.IsDirty);
        await workspace.UndoAsync(default);
        Assert.Single(workspace.Pages);
        await workspace.RedoAsync(default);
        await workspace.SaveAsync(null, default);
        Assert.Equal(6, PdfDocument.Load(source.Bytes).Inspect().PageCount);
        Assert.Equal(first.Reads, first.ClosedReads);
        Assert.Equal(second.Reads, second.ClosedReads);
        Assert.Equal(0, first.Writes + second.Writes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FailedImportBatchPreservesWorkspaceAndClosesProviderStreams(bool denyRead) {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://documents/source", CreatePdf());
        var first = new TestStorageFile("content://documents/import-one", CreatePdf(2));
        var second = new TestStorageFile("content://documents/import-two", [1, 2, 3]) { DenyRead = denyRead };
        string location = await storage.RegisterAsync(source.Item, default);
        var imports = await storage.RegisterManyAsync([first.Item, second.Item], default);
        using var root = new TestDirectory();
        using var workspace = await PdfWorkspace.OpenAsync(location, default, new(root.Path), storage: storage);
        byte[] original = workspace.CreateDocumentSnapshot().ToBytes();
        await Assert.ThrowsAnyAsync<Exception>(() => workspace.ImportAsync(imports, 1, default));
        Assert.Equal(original, workspace.CreateDocumentSnapshot().ToBytes());
        Assert.False(workspace.IsDirty);
        Assert.Equal(first.Reads, first.ClosedReads);
        Assert.Equal(denyRead ? 0 : second.Reads, second.ClosedReads);
        Assert.Equal(0, source.Writes + first.Writes + second.Writes);
    }

    [Fact]
    public async Task CancelledBatchRegistrationReleasesEveryUnretainedPickerItem() {
        using var storage = new StudioStorageAccess();
        var first = new TestStorageFile("content://documents/first", CreatePdf());
        var second = new TestStorageFile("content://documents/second", CreatePdf());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            storage.RegisterManyAsync([first.Item, second.Item, second.Item], cancellation.Token));
        Assert.Equal(1, first.Disposals);
        Assert.Equal(1, second.Disposals);
    }

    [Fact]
    public async Task BoundedProviderReadClosesStreamWhenBatchAllowanceIsExceeded() {
        using var storage = new StudioStorageAccess();
        var file = new TestStorageFile("content://documents/large", new byte[100]);
        string location = await storage.RegisterAsync(file.Item, default);
        await Assert.ThrowsAsync<InvalidDataException>(() => storage.ReadSnapshotAsync(location, default, 20));
        Assert.Equal(file.Reads, file.ClosedReads);
    }

    [Fact]
    public async Task FailedSaveOfUneditedSourceStillRequiresPreservingItsOriginalBytes() {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://documents/source", CreatePdf()) { FailWrite = true };
        byte[] original = source.Bytes;
        string location = await storage.RegisterAsync(source.Item, default);
        using var root = new TestDirectory();
        var recovery = new PdfWorkspaceRecoveryStore(root.Path);
        using var workspace = await PdfWorkspace.OpenAsync(location, default, recovery, storage: storage);
        Assert.False(workspace.IsDirty);
        await Assert.ThrowsAsync<IOException>(() => workspace.SaveAsync(null, default));
        Assert.True(workspace.IsDirty);
        Assert.Equal(original, recovery.ReadVerifiedSnapshot(location, workspace.BaseFingerprint));
    }

    [Fact]
    public async Task ProviderDocumentEditsSaveVerifiedBytesAndReleaseReadScopes() {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://documents/opaque-id", CreatePdf());
        string location = await storage.RegisterAsync(source.Item, default);
        using var root = new TestDirectory();
        var recovery = new PdfWorkspaceRecoveryStore(root.Path);
        using var workspace = await PdfWorkspace.OpenAsync(location, default, recovery, storage: storage);
        Assert.Equal(source.Name, workspace.FileName);
        Assert.Equal(location, workspace.Path);
        await workspace.DuplicateAsync([1], default);
        Assert.NotNull(new PdfWorkspaceRecoveryStore(root.Path).ReadVerifiedSnapshot(location, workspace.BaseFingerprint));
        await workspace.SaveAsync(null, default);
        Assert.False(workspace.IsDirty);
        Assert.False(workspace.HasRecovery);
        Assert.Equal(2, PdfDocument.Load(source.Bytes).Inspect().Pages.Count);
        Assert.Equal(source.Reads, source.ClosedReads);
        Assert.Equal(1, source.Writes);
        Assert.Equal(0, source.Disposals);
        storage.Dispose();
        Assert.Equal(1, source.Disposals);
    }

    [Fact]
    public async Task ChangedProviderSourceIsNotOverwritten() {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://documents/source", CreatePdf());
        string location = await storage.RegisterAsync(source.Item, default);
        using var root = new TestDirectory();
        using var workspace = await PdfWorkspace.OpenAsync(location, default, new(root.Path), storage: storage);
        await workspace.DuplicateAsync([1], default);
        byte[] external = CreatePdf(3);
        source.Bytes = external;
        await Assert.ThrowsAsync<IOException>(() => workspace.SaveAsync(null, default));
        Assert.Same(external, source.Bytes);
        Assert.Equal(0, source.Writes);
        Assert.True(workspace.IsDirty);
        Assert.NotNull(new PdfWorkspaceRecoveryStore(root.Path).ReadVerifiedSnapshot(location, workspace.BaseFingerprint));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task FailedPublicationRetainsEditsAndRecoveryForLocalSaveAs(bool interrupted) {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://documents/source", CreatePdf()) {
            FailWrite = interrupted, CorruptWrite = !interrupted
        };
        string location = await storage.RegisterAsync(source.Item, default);
        using var root = new TestDirectory();
        using var workspace = await PdfWorkspace.OpenAsync(location, default, new(root.Path), storage: storage);
        await workspace.DuplicateAsync([1], default);
        await Assert.ThrowsAsync<IOException>(() => workspace.SaveAsync(null, default));
        Assert.True(workspace.IsDirty);
        Assert.NotNull(new PdfWorkspaceRecoveryStore(root.Path).ReadVerifiedSnapshot(location, workspace.BaseFingerprint));
        Assert.Equal(location, workspace.Path);
        Assert.Equal(2, workspace.CreateDocumentSnapshot().Inspect().Pages.Count);
        await workspace.UndoAsync(default);
        Assert.True(workspace.IsDirty);
        await workspace.RedoAsync(default);
        string destination = System.IO.Path.Combine(root.Path, "recovered.pdf");
        await workspace.SaveAsync(destination, default);
        Assert.False(workspace.IsDirty);
        Assert.Equal(2, PdfDocument.Load(destination).Inspect().Pages.Count);
        Assert.Equal(source.Reads, source.ClosedReads);
    }

    [Fact]
    public async Task ProviderSaveAsKeepsOriginalAndChecksLiveOutputOwnership() {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://documents/source", CreatePdf());
        var destination = new TestStorageFile("content://documents/copy", []);
        string location = await storage.RegisterAsync(source.Item, default);
        string output = await storage.RegisterAsync(destination.Item, default);
        using var root = new TestDirectory();
        bool allowed = false;
        using var workspace = await PdfWorkspace.OpenAsync(location, default, new(root.Path), storage: storage,
            canPublishOutput: (_, _) => ValueTask.FromResult(allowed));
        await workspace.DuplicateAsync([1], default);
        await Assert.ThrowsAsync<IOException>(() => workspace.SaveAsync(output, default));
        Assert.Equal(0, destination.Writes);
        allowed = true;
        await workspace.SaveAsync(output, default);
        Assert.Single(PdfDocument.Load(source.Bytes).Inspect().Pages);
        Assert.Equal(2, PdfDocument.Load(destination.Bytes).Inspect().Pages.Count);
        Assert.Equal(output, workspace.Path);
        Assert.False(workspace.IsDirty);
    }

    [Fact]
    public async Task SessionBookmarkRestoresStreamAccessWithoutLocalPath() {
        using var root = new TestDirectory();
        var source = new TestStorageFile("content://documents/source", CreatePdf());
        StudioSessionSnapshot snapshot;
        using (var storage = new StudioStorageAccess()) {
            string location = await storage.RegisterAsync(source.Item, default);
            var store = new StudioSessionStore(System.IO.Path.Combine(root.Path, "session.json"));
            store.Save(new(1, DateTimeOffset.UtcNow, location, [new(location,
                await storage.FingerprintAsync(location, default), new()) { Storage = storage.Describe(location) }]));
            snapshot = store.Load();
        }
        var restored = new TestStorageFile(source.Location.AbsoluteUri, source.Bytes, bookmark: source.Bookmark);
        using var reopened = new StudioStorageAccess();
        reopened.Attach(() => restored.CreateProvider());
        StudioSessionDocument document = Assert.Single(snapshot.Documents);
        reopened.Remember(document.Storage!);
        using var workspace = await PdfWorkspace.OpenAsync(document.Path, default, new(root.Path), storage: reopened);
        Assert.Equal(document.Fingerprint, workspace.BaseFingerprint);
        Assert.Equal(restored.Name, workspace.FileName);
    }

    [Fact]
    public async Task ProviderCaseVariantsKeepSeparateRecoveryIdentities() {
        using var root = new TestDirectory();
        var recovery = new PdfWorkspaceRecoveryStore(root.Path);
        const string upper = "content://documents/CaseSensitive";
        const string lower = "content://documents/casesensitive";
        string fingerprint = new('A', 64);
        await recovery.WriteAsync(upper, fingerprint, [1, 2], 1, default);
        await recovery.WriteAsync(lower, fingerprint, [3, 4], 1, default);
        Assert.Equal(new byte[] { 1, 2 }, recovery.ReadVerifiedSnapshot(upper, fingerprint));
        Assert.Equal(new byte[] { 3, 4 }, recovery.ReadVerifiedSnapshot(lower, fingerprint));
    }

    [Fact]
    public void ProviderCaseVariantsKeepSeparateReadingAndRestartRecords() {
        using var root = new TestDirectory();
        const string upper = "content://documents/CaseSensitive";
        const string lower = "content://documents/casesensitive";
        string viewPath = System.IO.Path.Combine(root.Path, "views.json");
        var views = new StudioDocumentViewStore(viewPath);
        views.Put(upper, new() { PageNumber = 4 });
        views.Put(lower, new() { PageNumber = 7 });
        var reopened = new StudioDocumentViewStore(viewPath);
        Assert.Equal(4, reopened.Get(upper).PageNumber);
        Assert.Equal(7, reopened.Get(lower).PageNumber);
        var session = new StudioSessionStore(System.IO.Path.Combine(root.Path, "session.json"));
        session.Save(new(1, DateTimeOffset.UtcNow, upper, [
            new(upper, new string('A', 64), new()), new(lower, new string('B', 64), new())
        ]));
        Assert.Equal(2, session.Load().Documents.Count);
    }

    [Fact]
    public async Task ReselectingDocumentRefreshesExpiredAccess() {
        using var storage = new StudioStorageAccess();
        var previous = new TestStorageFile("content://documents/source", CreatePdf()) { DenyRead = true };
        string location = await storage.RegisterAsync(previous.Item, default);
        var current = new TestStorageFile(location, CreatePdf());
        await storage.RegisterAsync(current.Item, default);
        Assert.NotEmpty((await storage.ReadSnapshotAsync(location, default)).Bytes);
        Assert.Equal(current.Bookmark, storage.Describe(location).Bookmark);
        storage.Dispose();
        Assert.Equal(1, previous.Disposals);
        Assert.Equal(1, current.Disposals);
    }

    [Fact]
    public async Task RevokedBookmarkLeavesVerifiedRecoveryAvailable() {
        using var root = new TestDirectory();
        using var storage = new StudioStorageAccess();
        var file = new TestStorageFile("content://documents/unavailable", CreatePdf());
        storage.Attach(() => file.CreateProvider(denyBookmark: true));
        var reference = new StudioStorageReference(file.Location.AbsoluteUri, file.Name, file.Bookmark);
        storage.Remember(reference);
        var recovery = new PdfWorkspaceRecoveryStore(root.Path);
        string fingerprint = new('A', 64);
        byte[] edits = CreatePdf(2);
        await recovery.WriteAsync(reference.Location, fingerprint, edits, 1, default);
        await Assert.ThrowsAsync<IOException>(() => storage.ReadSnapshotAsync(reference.Location, default));
        Assert.Equal(edits, recovery.ReadVerifiedSnapshot(reference.Location, fingerprint));
        Assert.Equal(0, file.Reads);
        Assert.Equal(0, file.Writes);
    }

    [Fact]
    public async Task ProtectedAndExtractedCopiesUseProviderPublicationWithoutChangingSource() {
        using var root = new TestDirectory();
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://documents/source", CreatePdf(2));
        var protectedCopy = new TestStorageFile("content://documents/protected", []);
        var extractedCopy = new TestStorageFile("content://documents/extracted", []);
        string location = await storage.RegisterAsync(source.Item, default);
        string protectedLocation = await storage.RegisterAsync(protectedCopy.Item, default);
        string extractedLocation = await storage.RegisterAsync(extractedCopy.Item, default);
        using var workspace = await PdfWorkspace.OpenAsync(location, default, new(root.Path), storage: storage);
        await workspace.SaveProtectedCopyAsync(protectedLocation, new PdfStandardEncryptionOptions("reader-password"), null, default);
        await workspace.ExtractAsync([2], extractedLocation, default);
        Assert.Equal(2, PdfDocument.Load(protectedCopy.Bytes, new PdfLoadOptions { Password = "reader-password" }).Inspect().Pages.Count);
        Assert.Single(PdfDocument.Load(extractedCopy.Bytes).Inspect().Pages);
        Assert.Equal(0, source.Writes);
        Assert.Equal(location, workspace.Path);
        Assert.False(workspace.IsDirty);
    }

    internal static byte[] CreatePdf(int count = 1) {
        var document = PdfDocument.Create(compose => {
            for (int index = 0; index < count; index++) compose.Page(page => page.Size(300, 400));
        });
        using var stream = new MemoryStream();
        document.Save(stream);
        return stream.ToArray();
    }

    private sealed class TestDirectory : IDisposable {
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "studio-provider-tests-" + Guid.NewGuid().ToString("N"));
        internal TestDirectory() => Directory.CreateDirectory(Path);
        public void Dispose() => Directory.Delete(Path, true);
    }
}
