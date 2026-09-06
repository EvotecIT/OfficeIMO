using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSessionCloseTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CloseCoordinatesTheEntireRestoreIncludingDocumentsOpenedAfterTheDialog(bool cancel) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            byte[] plain = PdfDocument.Create(document => document.Page(page => page.Size(200, 300))).ToBytes();
            byte[] encrypted = PdfDocument.Load(plain).Security.Encrypt(new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" }).Pdf;
            string first = Path.Combine(services.Paths.Root, "first.pdf"), second = Path.Combine(services.Paths.Root, "second.pdf");
            File.WriteAllBytes(first, encrypted); File.WriteAllBytes(second, encrypted);
            var firstReached = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var secondReached = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var firstRelease = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var secondRelease = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            int prompts = 0;
            using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                services: services, openDocumentInTab: open, promptPdfPassword: async (_, _, token) => {
                    bool initial = ++prompts == 1;
                    (initial ? firstReached : secondReached).TrySetResult();
                    await (initial ? firstRelease : secondRelease).Task.WaitAsync(token);
                    return "open";
                }), _ => { });
            using var restore = new StudioSessionController(host, services, _ => Task.FromResult<string?>(null));
            foreach (string path in new[] { first, second })
                restore.Pending.Add(new(new(path, PdfWorkspaceRecoveryStore.Fingerprint(encrypted), new())));
            var owner = new Window();
            Task? restoring = null;
            try {
                owner.Show();
                restoring = restore.RestoreCommand.ExecuteAsync(null);
                await firstReached.Task.WaitAsync(TimeSpan.FromSeconds(10));
                var dialog = new ActiveOperationsDialog(host.OperationDocuments, services.Localizer, host, restore);
                Task<bool> closing = dialog.ShowDialog<bool>(owner);
                dialog.GetVisualDescendants().OfType<Button>().Single(button => button.Name == (cancel ? "CancelWorkAndClose" : "WaitAndClose"))
                    .RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                if (cancel) {
                    await restoring.WaitAsync(TimeSpan.FromSeconds(10));
                    Assert.True(await closing.WaitAsync(TimeSpan.FromSeconds(5)));
                    Assert.Equal(1, prompts);
                    Assert.Empty(host.Tabs);
                    Assert.Equal(2, restore.Pending.Count);
                    Assert.Equal(encrypted, File.ReadAllBytes(first));
                    Assert.Equal(encrypted, File.ReadAllBytes(second));
                    restore.CaptureForShutdown();
                    Assert.Equal(2, services.DocumentHistory.RestartSession.Load().Documents.Count);
                    return true;
                }
                firstRelease.TrySetResult();
                await secondReached.Task.WaitAsync(TimeSpan.FromSeconds(10));
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                Assert.True(restore.IsBusy);
                Assert.False(closing.IsCompleted);
                secondRelease.TrySetResult();
                await restoring.WaitAsync(TimeSpan.FromSeconds(10));
                Assert.True(await closing.WaitAsync(TimeSpan.FromSeconds(5)));
                Assert.Equal(2, host.Tabs.Count);
            } finally {
                firstRelease.TrySetResult(); secondRelease.TrySetResult();
                if (restoring is not null) await restoring;
                foreach (Window dialog in owner.OwnedWindows.ToArray()) dialog.Close(false);
                owner.Close();
            }
            return true;
        }, CancellationToken.None);
    }
}
