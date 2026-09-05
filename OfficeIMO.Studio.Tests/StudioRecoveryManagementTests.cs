using Avalonia.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioRecoveryManagementTests {
    [Fact]
    public async Task FailedExplicitDiscardRetainsRecoveryChoiceAndReportsFailure() {
        if (!OperatingSystem.IsWindows()) return; // Windows sharing denies deletion of an open file.
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = Path.Combine(Path.GetTempPath(), "studio-discard-failure-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try {
                string source = Path.Combine(root, "source.pdf");
                PdfDocument.Create(builder => builder.Page(page => page.Size(600, 800))).Save(source);
                byte[] original = File.ReadAllBytes(source);
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                string snapshot = await services.Recovery.WriteAsync(source,
                    PdfWorkspaceRecoveryStore.Fingerprint(original), original, 1, CancellationToken.None);
                using var reader = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
                await reader.OpenDocumentAsync(source);
                Assert.True(reader.HasRecovery);
                using (var locked = new FileStream(snapshot, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    await reader.DiscardRecoveryCommand.ExecuteAsync(null);
                    Assert.True(reader.HasRecovery);
                    Assert.False(string.IsNullOrWhiteSpace(reader.ErrorMessage));
                    Assert.NotEqual(services.Localizer.Get("Workspace.RecoveryDiscarded"), reader.OperationStatus);
                }
                await reader.DiscardRecoveryCommand.ExecuteAsync(null);
                Assert.False(reader.HasRecovery);
                Assert.Null(reader.ErrorMessage);
                Assert.Equal(original, File.ReadAllBytes(source));
            } finally { Directory.Delete(root, recursive: true); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ConfirmedCleanupPreservesOpenEditsAndSourcesAndRefreshesRecoveryChoices() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = Path.Combine(Path.GetTempPath(), "studio-recovery-settings-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try {
                string source = Path.Combine(root, "source.pdf");
                PdfDocument.Create(builder => builder.Page(page => page.Size(600, 800))).Save(source);
                byte[] original = File.ReadAllBytes(source);
                string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(original);
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                using var editing = Host(services);
                using var savedSession = new StudioSessionController(editing, services, _ => Task.FromResult<string?>(null));
                await editing.OpenDocumentAsync(source);
                var document = editing.ActiveDocument;
                document.SetOrganizerSelection([document.OrganizerPages[0]]);
                await document.DuplicateSelectedCommand.ExecuteAsync(null);
                savedSession.CaptureForShutdown();

                using var pendingHost = Host(services);
                using var pending = new StudioSessionController(pendingHost, services, _ => Task.FromResult<string?>(null));
                await pending.InspectAsync();
                Assert.True(Assert.Single(pending.Pending).HasRecovery);
                using var reader = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
                await reader.OpenDocumentAsync(source);
                Assert.True(reader.HasRecovery);
                bool notified = false;
                reader.PropertyChanged += (_, args) => { if (args.PropertyName == nameof(reader.HasRecovery)) notified = true; };

                var settings = pendingHost.ActiveDocument.Settings;
                await settings.ClearRecoveryCommand.ExecuteAsync(null);
                Assert.NotNull(services.Recovery.Find(source, fingerprint));
                settings.RequestRecoveryClearCommand.Execute(null);
                Assert.True(settings.ConfirmRecoveryClear);
                settings.CancelRecoveryClearCommand.Execute(null);
                Assert.False(settings.ConfirmRecoveryClear);
                Assert.NotNull(services.Recovery.Find(source, fingerprint));

                settings.RequestRecoveryClearCommand.Execute(null);
                await settings.ClearRecoveryCommand.ExecuteAsync(null);
                Dispatcher.UIThread.RunJobs();
                Assert.False(settings.ConfirmRecoveryClear);
                Assert.False(settings.IsRecoveryBusy);
                Assert.True(settings.HasRecoveryStatus);
                Assert.Null(services.Recovery.Find(source, fingerprint));
                Assert.False(Assert.Single(pending.Pending).HasRecovery);
                Assert.False(reader.HasRecovery);
                Assert.True(notified);
                Assert.True(document.IsDirty);
                Assert.Equal(2, document.Pages.Count);
                Assert.Equal(original, File.ReadAllBytes(source));
                Assert.False(reader.IsDirty);
            } finally {
                Directory.Delete(root, recursive: true);
            }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task StartupRunsExpiryCleanupInTheConfiguredProfile() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = Path.Combine(Path.GetTempPath(), "studio-recovery-startup-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try {
                var services = StudioApplicationServices.Create(new StudioDataPaths(root));
                string snapshot = await services.Recovery.WriteAsync(Path.Combine(root, "source.pdf"),
                    PdfWorkspaceRecoveryStore.Fingerprint([1]), [1], 1, CancellationToken.None);
                await File.WriteAllBytesAsync(snapshot, [0]);
                File.SetLastWriteTimeUtc(snapshot, DateTime.UtcNow.AddDays(-31));
                var completed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                services.Recovery.MaintenanceCompleted += (_, _) => completed.TrySetResult();
                var window = new MainWindow(services);
                try {
                    window.Show();
                    await completed.Task.WaitAsync(TimeSpan.FromSeconds(5));
                    Assert.False(File.Exists(snapshot));
                    Assert.True(window.ViewModel.IsHomeMode);
                } finally {
                    window.Close();
                }
            } finally {
                Directory.Delete(root, recursive: true);
            }
            return true;
        }, CancellationToken.None);
    }

    private static StudioDocumentTabHost Host(StudioApplicationServices services) => new(open => new MainWindowViewModel(
        _ => Task.FromResult<string?>(null), openDocumentInTab: open, services: services), _ => { });
}
