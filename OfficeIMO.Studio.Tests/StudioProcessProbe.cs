using Avalonia;
using Avalonia.Headless;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

/// <summary>Child-process entry point for crash/restart acceptance using the real Studio services.</summary>
internal static class StudioProcessProbe {
    private static string _profile = string.Empty;

    public static AppBuilder BuildAvaloniaApp() => AppBuilder
        .Configure(() => new App(StudioApplicationServices.Create(new StudioDataPaths(_profile))))
        .UseSkia().UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = false });

    public static async Task<int> Main(string[] args) {
        if (args is not ["--studio-process-probe", var mode, var root] ||
            mode is not ("write" or "private-write" or "restore" or "recover-missing" or "private-verify")) return 2;
        root = Path.GetFullPath(root);
        if (!Directory.Exists(root)) return 2;
        _profile = Path.Combine(root, "profile");
        var app = HeadlessUnitTestSession.StartNew(typeof(StudioProcessProbe), AvaloniaTestIsolationLevel.PerTest);
        try {
            await app.Dispatch(async () => {
                var services = ((App)Application.Current!).Services;
                string source = Path.Combine(root, "source.pdf");
                string destination = Path.Combine(root, "recovered.pdf");
                using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(
                    _ => Task.FromResult<string?>(null), pickSavePdf: _ => Task.FromResult<string?>(destination),
                    openDocumentInTab: open, services: services, recentDocumentStore: services.DocumentHistory.RecentDocuments), _ => { });
                using var session = new StudioSessionController(host, services, _ => Task.FromResult<string?>(destination));
                if (mode is "write" or "private-write") {
                    if (mode == "private-write") {
                        host.ActiveDocument.Settings.RememberSession = false;
                        host.ActiveDocument.Settings.ToggleDocumentHistoryCommand.Execute(null);
                        await host.ActiveDocument.Settings.ToggleRecoveryPersistenceCommand.ExecuteAsync(null);
                    }
                    await host.OpenDocumentAsync(source);
                    await DuplicateAsync(host.ActiveDocument);
                    host.ActiveDocument.RestoreSessionViewState(new() { PageNumber = 2, Zoom = 1.6, ZoomMode = ViewerZoomMode.Custom });
                    session.Flush();
                    if (mode == "write") {
                        Assert.Single(services.DocumentHistory.RestartSession.Load().Documents);
                        Assert.True(services.Recovery.HasSnapshotFiles(source));
                    } else AssertPrivate(services, source);
                    Console.WriteLine("READY");
                    Console.Out.Flush();
                    // The parent kills this process. No Dispose, close handler, or shutdown flush can run.
                    await Task.Delay(Timeout.InfiniteTimeSpan);
                } else if (mode == "private-verify") {
                    AssertPrivate(services, source);
                    Assert.Empty(session.Pending);
                    await host.OpenDocumentAsync(source);
                    await DuplicateAsync(host.ActiveDocument);
                    await host.ActiveDocument.SaveAsCommand.ExecuteAsync(null);
                    AssertPrivate(services, source);
                    Assert.Equal(2, host.ActiveDocument.Pages.Count);
                    Assert.False(host.ActiveDocument.IsDirty);
                } else {
                    Assert.Single(session.Pending);
                    await session.InspectAsync();
                    if (mode == "recover-missing") {
                        Assert.False(session.Pending[0].SourceExists);
                        Assert.True(session.Pending[0].HasRecovery);
                        await session.RecoverCopyCommand.ExecuteAsync(session.Pending[0]);
                    } else {
                        Assert.True(session.Pending[0].SourceUnchanged);
                        await session.RestoreCommand.ExecuteAsync(null);
                        Assert.Single(host.ActiveDocument.Pages);
                        Assert.True(host.ActiveDocument.HasRecovery);
                        await host.ActiveDocument.RestoreRecoveryCommand.ExecuteAsync(null);
                        await host.ActiveDocument.SaveAsCommand.ExecuteAsync(null);
                    }
                    Assert.Empty(session.Pending);
                    Assert.Equal(2, host.ActiveDocument.Pages.Count);
                    Assert.False(host.ActiveDocument.IsDirty);
                    Assert.Equal(1.6, host.ActiveDocument.Zoom, 5);
                    Assert.True(File.Exists(destination));
                }
                Console.WriteLine("VERIFIED");
                return true;
            }, CancellationToken.None);
            return 0;
        } catch (Exception error) {
            Console.Error.WriteLine(error);
            return 1;
        } finally {
            // Dispatch can complete on the UI thread; disposal must not join that same thread.
            await Task.Run(app.Dispose);
        }
    }

    private static async Task DuplicateAsync(MainWindowViewModel document) {
        document.SetOrganizerSelection([Assert.Single(document.OrganizerPages)]);
        await document.DuplicateSelectedCommand.ExecuteAsync(null);
        Assert.True(document.IsDirty);
        Assert.Equal(2, document.Pages.Count);
    }

    private static void AssertPrivate(StudioApplicationServices services, string source) {
        Assert.False(services.Preferences.Current.RememberSession);
        Assert.False(services.Preferences.Current.RememberDocumentHistory);
        Assert.False(services.Preferences.Current.CreateRecoverySnapshots);
        Assert.False(File.Exists(services.Paths.SessionPath));
        Assert.False(File.Exists(services.Paths.DocumentViewsPath));
        Assert.False(File.Exists(services.Paths.RecentDocumentsPath));
        Assert.False(services.Recovery.HasSnapshotFiles(source));
    }
}
