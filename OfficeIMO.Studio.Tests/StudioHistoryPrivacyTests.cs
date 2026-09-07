using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioHistoryPrivacyTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SessionChoicesRestoreTheirReadingStateWithoutEnablingDocumentHistory(bool recoverCopy) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Avalonia.Application.Current!).Services;
            string source = Path.Combine(services.Paths.Root, "session-source.pdf");
            string destination = Path.Combine(services.Paths.Root, "recovered.pdf");
            PdfDocument.Create(builder => {
                builder.Page(page => page.Size(600, 800));
                builder.Page(page => page.Size(600, 800));
            }).Save(source);
            byte[] bytes = File.ReadAllBytes(source);
            string fingerprint = OfficeIMO.Studio.Features.Workspace.PdfWorkspaceRecoveryStore.Fingerprint(bytes);
            services.DocumentHistory.SetRememberHistory(false);
            services.DocumentHistory.RestartSession.Save(new(1, DateTimeOffset.UtcNow, source,
                [new(source, fingerprint, new() { PageNumber = 2, Zoom = 1.8, ZoomMode = ViewerZoomMode.Custom, FocusReading = true })]));
            if (recoverCopy) {
                await services.Recovery.WriteAsync(source, fingerprint, bytes, 1, CancellationToken.None);
                File.Delete(source);
            }
            using var host = Host(services);
            using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(destination));
            if (recoverCopy) await session.RecoverCopyCommand.ExecuteAsync(Assert.Single(session.Pending));
            else await session.RestoreCommand.ExecuteAsync(null);
            Assert.Empty(session.Pending);
            Assert.Equal(2, host.ActiveDocument.SelectedPage?.PageNumber);
            Assert.Equal(1.8, host.ActiveDocument.Zoom, 5);
            Assert.True(host.ActiveDocument.IsFocusReading);
            Assert.False(File.Exists(services.Paths.DocumentViewsPath));
            Assert.False(File.Exists(services.Paths.RecentDocumentsPath));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task PrivateWorkDoesNotUpdateHistoryViewsSessionsOrRecoveryAcrossRestart() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Avalonia.Application.Current!).Services;
            string source = CreatePdf(services.Paths.Root, "ordinary.pdf");
            string privateSource = CreatePdf(services.Paths.Root, "private.pdf");
            using var host = Host(services);
            using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(null));
            await host.OpenDocumentAsync(source);
            host.ActiveDocument.SaveDocumentViewState();
            session.Flush();
            byte[] recent = File.ReadAllBytes(services.Paths.RecentDocumentsPath);
            byte[] views = File.ReadAllBytes(services.Paths.DocumentViewsPath);
            var settings = host.ActiveDocument.Settings;
            settings.ToggleDocumentHistoryCommand.Execute(null);
            settings.RememberSession = false;
            await settings.ToggleRecoveryPersistenceCommand.ExecuteAsync(null);
            Assert.False(settings.RememberDocumentHistory);
            Assert.False(settings.RememberSession);
            Assert.False(settings.CreateRecoverySnapshots);
            await host.OpenDocumentAsync(privateSource);
            host.ActiveDocument.ToggleFocusReadingCommand.Execute(null);
            host.ActiveDocument.SetOrganizerSelection([host.ActiveDocument.OrganizerPages[0]]);
            await host.ActiveDocument.DuplicateSelectedCommand.ExecuteAsync(null);
            host.ActiveDocument.SaveDocumentViewState();
            session.Flush();
            Assert.Equal(recent, File.ReadAllBytes(services.Paths.RecentDocumentsPath));
            Assert.Equal(views, File.ReadAllBytes(services.Paths.DocumentViewsPath));
            Assert.False(File.Exists(services.Paths.SessionPath));
            Assert.False(services.Recovery.HasSnapshotFiles(privateSource));
            Assert.DoesNotContain(host.ActiveDocument.RecentDocuments, item => item.Path == privateSource);

            var restarted = StudioApplicationServices.Create(services.Paths);
            using var second = Host(restarted);
            using var secondSession = new StudioSessionController(second, restarted, _ => Task.FromResult<string?>(null));
            Assert.Empty(restarted.DocumentHistory.RecentDocuments.Load());
            Assert.False(restarted.DocumentViews.Get(privateSource).FocusReading);
            Assert.Empty(secondSession.Pending);
            await second.OpenDocumentAsync(privateSource);
            second.ActiveDocument.SaveDocumentViewState();
            secondSession.Flush();
            Assert.Equal(recent, File.ReadAllBytes(services.Paths.RecentDocumentsPath));
            Assert.Equal(views, File.ReadAllBytes(services.Paths.DocumentViewsPath));
            settings.RequestHistoryClearCommand.Execute(null);
            settings.ClearHistoryCommand.Execute(null);
            Assert.False(File.Exists(services.Paths.RecentDocumentsPath));
            Assert.False(File.Exists(services.Paths.DocumentViewsPath));
            Assert.False(File.Exists(services.Paths.SessionPath));
            Assert.True(host.ActiveDocument.IsDirty);
            Assert.Single(PdfDocument.Load(privateSource).Read().Pages);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task PartialHistoryCleanupReportsFailureAndClearsOnlySuccessfullyDeletedRecords() {
        if (!OperatingSystem.IsWindows()) return;
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Avalonia.Application.Current!).Services;
            string source = CreatePdf(services.Paths.Root, "existing.pdf");
            string fingerprint = OfficeIMO.Studio.Features.Workspace.PdfWorkspaceRecoveryStore.Fingerprint(File.ReadAllBytes(source));
            services.DocumentHistory.RecentDocuments.Save([new(source, DateTimeOffset.UtcNow)]);
            services.DocumentViews.Put(source, new() { Zoom = 2 });
            services.DocumentHistory.RestartSession.Save(new(1, DateTimeOffset.UtcNow, source, [new(source, fingerprint, new())]));
            await services.Recovery.WriteAsync(source, fingerprint, File.ReadAllBytes(source), 1, CancellationToken.None);
            using var host = Host(services);
            using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(null));
            Assert.Single(session.Pending);
            var settings = host.ActiveDocument.Settings;
            using (var held = new FileStream(services.Paths.RecentDocumentsPath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                settings.RequestHistoryClearCommand.Execute(null);
                settings.ClearHistoryCommand.Execute(null);
                Assert.True(settings.HasHistoryStatus);
                Assert.Single(host.ActiveDocument.RecentDocuments);
                Assert.True(File.Exists(services.Paths.RecentDocumentsPath));
                Assert.False(File.Exists(services.Paths.DocumentViewsPath));
                Assert.False(File.Exists(services.Paths.SessionPath));
                Assert.Empty(session.Pending);
                Assert.Equal(1, services.DocumentViews.Get(source).Zoom);
                session.Flush();
                Assert.Null(services.DocumentHistory.RestartSession.Load().ActivePath);
                Assert.Empty(services.DocumentHistory.RestartSession.Load().Documents);
            }
            settings.RequestHistoryClearCommand.Execute(null);
            settings.ClearHistoryCommand.Execute(null);
            Assert.Empty(host.ActiveDocument.RecentDocuments);
            Assert.False(File.Exists(services.Paths.RecentDocumentsPath));
            Assert.True(services.Recovery.HasSnapshotFiles(source));
            Assert.True(File.Exists(source));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task FailedPrivacyPreferenceWriteRestoresDisplayedStateAndAllowsRetry() {
        if (!OperatingSystem.IsWindows()) return;
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(() => {
            var services = ((App)Avalonia.Application.Current!).Services;
            services.Preferences.Update(current => current with { RememberDocumentHistory = false });
            using var host = Host(services);
            var settings = host.ActiveDocument.Settings;
            using (var held = new FileStream(services.Paths.PreferencesPath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                settings.ToggleDocumentHistoryCommand.Execute(null);
                Assert.False(settings.RememberDocumentHistory);
                Assert.False(services.Preferences.Current.RememberDocumentHistory);
                Assert.True(settings.HasHistoryStatus);
                settings.RememberSession = false;
                Assert.True(settings.RememberSession);
                Assert.True(services.Preferences.Current.RememberSession);
            }
            settings.ToggleDocumentHistoryCommand.Execute(null);
            Assert.True(settings.RememberDocumentHistory);
            Assert.False(settings.HasHistoryStatus);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task SessionOptOutReportsRetainedLockedRecordAndExplicitCleanupCanRetry() {
        if (!OperatingSystem.IsWindows()) return;
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(() => {
            var services = ((App)Avalonia.Application.Current!).Services;
            services.DocumentHistory.RestartSession.Save(new(1, DateTimeOffset.UtcNow, null, []));
            using var host = Host(services);
            var settings = host.ActiveDocument.Settings;
            using (var held = new FileStream(services.Paths.SessionPath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                settings.RememberSession = false;
                Assert.False(settings.RememberSession);
                Assert.False(services.Preferences.Current.RememberSession);
                Assert.True(settings.HasHistoryStatus);
                Assert.True(File.Exists(services.Paths.SessionPath));
            }
            settings.RequestHistoryClearCommand.Execute(null);
            settings.ClearHistoryCommand.Execute(null);
            Assert.False(File.Exists(services.Paths.SessionPath));
            return true;
        }, CancellationToken.None);
    }

    private static StudioDocumentTabHost Host(StudioApplicationServices services) => new(open => new MainWindowViewModel(
        _ => Task.FromResult<string?>(null), openDocumentInTab: open, services: services,
        recentDocumentStore: services.DocumentHistory.RecentDocuments), _ => { });

    private static string CreatePdf(string root, string name) {
        string path = Path.Combine(root, name);
        PdfDocument.Create(builder => builder.Page(page => page.Size(600, 800))).Save(path);
        return path;
    }
}
