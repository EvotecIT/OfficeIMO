using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioMobileImportLifetimeTests {
    [Theory]
    [InlineData("open")]
    [InlineData("close")]
    [InlineData("dispose")]
    [InlineData("comparison")]
    [InlineData("health")]
    [InlineData("health-comparison")]
    [InlineData("page-export")]
    [InlineData("print")]
    [InlineData("ocr")]
    public async Task DelayedImportBelongsToTheInitiatingDocumentAndStopsBeforeCreatingACopy(string entry) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-import-lifetime-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                var selection = new TaskCompletionSource<IStorageFile?>(TaskCreationOptions.RunContinuationsAsynchronously);
                CancellationToken importToken = default;
                var view = new MobileWorkspaceView();
                var host = new MobileDocumentHost(services, view, _ => Task.CompletedTask,
                    () => throw new NotSupportedException("This test supplies its PDF picker directly."));
                using var controller = new MobileDocumentController(services, token => {
                    importToken = token;
                    return selection.Task; // A platform picker may return after cancellation was requested.
                }, _ => Task.CompletedTask, host);
                view.Connect(controller);
                var window = new Window { Content = view, Width = 1024, Height = 768 };
                try {
                    window.Show();
                    await controller.OpenSampleAsync();
                    var document = controller.Document;
                    string[] before = Directory.GetFiles(services.LocalDocuments!.Path, "*.pdf", SearchOption.AllDirectories);
                    var source = new TestStorageFile("content://mobile/late.pdf", File.ReadAllBytes(before[0]), "late.pdf");
                    IAsyncRelayCommand command = entry switch {
                        "comparison" => document.OpenComparisonCommand,
                        "health" => document.DocumentHealth.ChooseInputCommand,
                        "health-comparison" => document.DocumentHealth.ChooseComparisonCommand,
                        "page-export" => document.OutputWorkbench.PageExport.ChooseInputCommand,
                        "print" => document.OutputWorkbench.PrintPreview.ChooseInputCommand,
                        "ocr" => document.OcrWorkbench.ChooseInputCommand,
                        _ => document.OpenCommand
                    };
                    Task importing = command.ExecuteAsync(null);
                    Assert.False(importing.IsCompleted);
                    Assert.True(document.CanCancelOperation);
                    Task? closing = null;
                    if (entry == "dispose") controller.Dispose();
                    else if (entry == "close") {
                        closing = controller.Tabs.CloseTabAsync(Assert.Single(controller.Tabs.Tabs));
                        window.UpdateLayout();
                        var dialog = Assert.IsType<ActiveOperationsDialogContent>(view.FindControl<ContentControl>("DialogContent")!.Content);
                        var cancel = Assert.Single(dialog.GetVisualDescendants().OfType<Button>(), button => button.Name == "CancelWorkAndClose");
                        cancel.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                        Assert.False(closing.IsCompleted);
                    } else document.CancelCurrentOperation();
                    Assert.True(importToken.IsCancellationRequested);
                    selection.SetResult(source.Item);
                    await importing;
                    if (closing is not null) {
                        Dispatcher.UIThread.RunJobs();
                        await closing.WaitAsync(TimeSpan.FromSeconds(10));
                    }
                    Assert.Equal(entry is "close" or "dispose" ? 0 : 1, controller.Tabs.Tabs.Count);
                    Assert.Equal(before, Directory.GetFiles(services.LocalDocuments.Path, "*.pdf", SearchOption.AllDirectories));
                    Assert.False(document.CanCancelOperation);
                    Assert.Equal(1, source.Disposals);
                } finally {
                    selection.TrySetResult(null);
                    window.Close(); services.Storage.Dispose();
                }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }
}
