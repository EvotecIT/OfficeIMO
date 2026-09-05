using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSessionTests {
    [Fact]
    public void SessionStoreBoundsRecordsAndRejectsExpiredOrDamagedSnapshots() {
        string root = NewFolder();
        try {
            string path = Path.Combine(root, "session.json");
            var store = new StudioSessionStore(path);
            var entries = Enumerable.Range(0, 40).Select(index => new StudioSessionDocument(
                Path.Combine(root, index + ".pdf"), new string('A', 64), new() { PageNumber = -50 })).ToArray();
            store.Save(new(1, DateTimeOffset.UtcNow, entries[0].Path, entries));
            Assert.Equal(32, store.Load().Documents.Count);
            Assert.All(store.Load().Documents, document => Assert.Equal(1, document.View.PageNumber));
            File.WriteAllText(path, JsonSerializer.Serialize(new StudioSessionSnapshot(1, DateTimeOffset.UtcNow.AddDays(-31), null, entries)));
            Assert.Empty(store.Load().Documents);
            File.WriteAllText(path, "{ broken");
            Assert.Empty(store.Load().Documents);
            File.WriteAllText(path, JsonSerializer.Serialize(new StudioSessionSnapshot(99, DateTimeOffset.UtcNow, null, entries)));
            Assert.Empty(store.Load().Documents);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task RestartRestoresTabsActiveDocumentAndReadingPositionWithoutApplyingEdits() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = NewFolder();
            try {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                string first = CreatePdf(root, "first.pdf", 2), second = CreatePdf(root, "second.pdf", 1);
                using (var host = Host(services)) {
                    using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(null));
                    await host.OpenDocumentAsync(first);
                    var document = host.ActiveDocument;
                    document.SetOrganizerSelection([document.OrganizerPages[0]]);
                    await document.DuplicateSelectedCommand.ExecuteAsync(null);
                    document.SelectedPage = document.Pages[1];
                    document.ActualSizeCommand.Execute(null);
                    document.ZoomInCommand.Execute(null);
                    await host.OpenDocumentAsync(second);
                    host.SelectedTab = host.Tabs[0];
                    session.CaptureForShutdown();
                }
                using var reopened = Host(services);
                using var restore = new StudioSessionController(reopened, services, _ => Task.FromResult<string?>(null));
                Assert.Equal(2, restore.Pending.Count);
                await restore.RestoreCommand.ExecuteAsync(null);
                Assert.Empty(restore.Pending);
                Assert.Equal(2, reopened.Tabs.Count);
                Assert.Equal(first, reopened.ActiveDocument.DocumentPath);
                Assert.Equal(2, reopened.ActiveDocument.SelectedPage!.PageNumber);
                Assert.Equal(1.25, reopened.ActiveDocument.Zoom);
                Assert.Equal(2, reopened.ActiveDocument.Pages.Count);
                Assert.False(reopened.ActiveDocument.IsDirty);
                Assert.True(reopened.ActiveDocument.HasRecovery);
                await reopened.ActiveDocument.RestoreRecoveryCommand.ExecuteAsync(null);
                await reopened.ActiveDocument.SaveCommand.ExecuteAsync(null);
                restore.CaptureForShutdown();
                var saved = new StudioSessionStore(services.Paths.SessionPath).Load();
                Assert.Equal(PdfWorkspaceRecoveryStore.Fingerprint(File.ReadAllBytes(first)), saved.Documents[0].Fingerprint);
                Assert.False(reopened.ActiveDocument.IsDirty);
            } finally { Directory.Delete(root, true); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ChangedAndMissingSourcesRequireChoicesAndRecoveryNeverReplacesAFile() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = NewFolder();
            try {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                string source = CreatePdf(root, "source.pdf", 1);
                using (var host = Host(services)) {
                    using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(null));
                    await host.OpenDocumentAsync(source);
                    host.ActiveDocument.SetOrganizerSelection([host.ActiveDocument.OrganizerPages[0]]);
                    await host.ActiveDocument.DuplicateSelectedCommand.ExecuteAsync(null);
                    session.CaptureForShutdown();
                }
                CreatePdf(root, "source.pdf", 3);
                byte[] changed = File.ReadAllBytes(source);
                string destination = source;
                using var restored = Host(services);
                using var choices = new StudioSessionController(restored, services, _ => Task.FromResult<string?>(destination));
                await choices.RestoreCommand.ExecuteAsync(null);
                var pending = Assert.Single(choices.Pending);
                Assert.Empty(restored.Tabs);
                Assert.True(pending.SourceExists);
                Assert.False(pending.SourceUnchanged);
                Assert.True(pending.HasRecovery);
                await choices.RecoverCopyCommand.ExecuteAsync(pending);
                Assert.True(choices.HasError);
                Assert.Equal(changed, File.ReadAllBytes(source));
                File.Delete(source);
                await choices.InspectAsync();
                Assert.False(pending.SourceExists);
                destination = Path.Combine(root, "recovered.pdf");
                await choices.RecoverCopyCommand.ExecuteAsync(pending);
                Assert.False(choices.HasError);
                Assert.False(File.Exists(source));
                Assert.Equal(destination, restored.ActiveDocument.DocumentPath);
                Assert.Equal(2, restored.ActiveDocument.Pages.Count);
                Assert.Empty(choices.Pending);
            } finally { Directory.Delete(root, true); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ManuallyReopeningAndClosingAPreviousDocumentRemovesItsRestartEntry() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = NewFolder();
            try {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                string source = CreatePdf(root, "source.pdf", 1);
                using (var original = Host(services)) {
                    using var saved = new StudioSessionController(original, services, _ => Task.FromResult<string?>(null));
                    await original.OpenDocumentAsync(source);
                    saved.CaptureForShutdown();
                }
                using var host = Host(services);
                using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(null));
                Assert.Single(session.Pending);
                await host.OpenDocumentAsync(source);
                Assert.Empty(session.Pending);
                await host.CloseSelectedTabAsync();
                session.Flush();
                Assert.Empty(new StudioSessionStore(services.Paths.SessionPath).Load().Documents);
            } finally { Directory.Delete(root, true); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task DisablingSessionMemoryClearsPathsAndStopsNewSnapshots() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = NewFolder();
            try {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                using var host = Host(services);
                using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(null));
                await host.OpenDocumentAsync(CreatePdf(root, "private.pdf", 1));
                session.Flush();
                Assert.True(File.Exists(services.Paths.SessionPath));
                services.Preferences.Update(value => value with { RememberSession = false });
                session.Flush();
                Assert.False(File.Exists(services.Paths.SessionPath));
                Assert.False(session.HasPending);
            } finally { Directory.Delete(root, true); }
            return true;
        }, CancellationToken.None);
    }

    private static StudioDocumentTabHost Host(StudioApplicationServices services) => new(open => new MainWindowViewModel(
        _ => Task.FromResult<string?>(null), openDocumentInTab: open, services: services), _ => { });
    private static string NewFolder() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-session-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        return root;
    }
    private static string CreatePdf(string root, string name, int pages) {
        string path = Path.Combine(root, name);
        PdfDocument.Create(builder => { for (int i = 0; i < pages; i++) builder.Page(page => page.Size(600, 800)); }).Save(path);
        return path;
    }
}
