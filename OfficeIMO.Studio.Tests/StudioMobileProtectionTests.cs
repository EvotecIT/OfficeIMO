using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioMobileProtectionTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task SavedFolderCopySurvivesRestartAndSharingPreservesTheSavedFile(bool existing) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-save-as-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var paths = new StudioDataPaths(Path.Combine(root, "data"));
                var documents = new StudioLocalDocumentRoot(Path.Combine(root, "documents"));
                var services = StudioApplicationServices.Create(paths, documents);
                string outputPath = Path.Combine(root, "output", "saved.pdf");
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
                if (existing) File.WriteAllBytes(outputPath, StudioProviderDocumentTests.CreatePdf(2));
                IStorageFolder folder = TestLocalOutputFolder.Create(Path.GetDirectoryName(outputPath)!);
                var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
                ((TestStorageFile.StorageProxy)(object)provider).Call = (method, _) => method switch {
                    "OpenFolderPickerAsync" => Task.FromResult<IReadOnlyList<IStorageFolder>>([folder]),
                    _ => throw new NotSupportedException(method)
                };
                var view = new MobileWorkspaceView();
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null),
                    _ => Task.CompletedTask, new MobileDocumentHost(services, view, _ => Task.CompletedTask, () => provider));
                view.Connect(controller);
                var window = new Window { Content = view, Width = 1024, Height = 768 };
                try {
                    window.Show();
                    await controller.OpenSampleAsync();
                    string originalPath = controller.Document.DocumentPath!;
                    byte[] original = File.ReadAllBytes(originalPath);
                    Task saving = controller.Document.SaveAsCommand.ExecuteAsync(null);
                    await ChooseName(view, window, saving, services, "saved.pdf", () => controller.Document.ErrorMessage);
                    if (existing) Click(window, await WaitForDialog<MobileReplaceFileDialogContent>(view, saving, () => controller.Document.ErrorMessage), services.Localizer.Get("Common.Save"));
                    await saving;
                    Assert.False(controller.Document.HasError, controller.Document.ErrorMessage);
                    Assert.Equal(outputPath, controller.Document.DocumentPath);
                    Assert.Equal(original, File.ReadAllBytes(originalPath));
                    controller.Dispose();
                    services.Storage.Dispose();
                    var restartedServices = StudioApplicationServices.Create(paths, documents);
                    string? sharedPath = null;
                    using var restarted = new MobileDocumentController(restartedServices, _ => Task.FromResult<IStorageFile?>(null), path => {
                        sharedPath = path;
                        Assert.Equal(outputPath, path);
                        Assert.Equal(File.ReadAllBytes(outputPath), File.ReadAllBytes(path));
                        return Task.CompletedTask;
                    }, new MobileDocumentHost(restartedServices, view, _ => Task.CompletedTask, () => provider));
                    view.Connect(restarted);
                    await restarted.RestoreAsync();
                    Assert.True(restarted.Document.HasDocument, restarted.Document.ErrorMessage);
                    Assert.Equal(outputPath, restarted.Document.DocumentPath);
                    await restarted.ShareAsync();
                    Assert.NotNull(sharedPath);
                    Assert.True(File.Exists(sharedPath));
                    restartedServices.Storage.Dispose();
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ProtectionPublishesThroughFilesAndTheProtectedCopyCanBeImported(bool existing) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-protection-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                string outputPath = Path.Combine(root, "output", "protected.pdf");
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
                if (existing) File.WriteAllBytes(outputPath, StudioProviderDocumentTests.CreatePdf(2));
                IStorageFolder folder = TestLocalOutputFolder.Create(Path.GetDirectoryName(outputPath)!);
                var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
                ((TestStorageFile.StorageProxy)(object)provider).Call = (method, _) => method switch {
                    "OpenFolderPickerAsync" => Task.FromResult<IReadOnlyList<IStorageFolder>>([folder]),
                    _ => throw new NotSupportedException(method)
                };
                var view = new MobileWorkspaceView();
                var host = new MobileDocumentHost(services, view, _ => Task.CompletedTask, () => provider);
                using var controller = new MobileDocumentController(services,
                    _ => Task.FromResult<IStorageFile?>(new TestStorageFile("content://mobile/import.pdf", File.ReadAllBytes(outputPath), "protected.pdf").Item),
                    _ => Task.CompletedTask, host);
                view.Connect(controller);
                var window = new Window { Content = view, Width = 1024, Height = 768 };
                try {
                    window.Show();
                    await controller.OpenSampleAsync();
                    var document = controller.Document;
                    byte[] original = File.ReadAllBytes(document.DocumentPath!);
                    document.ProtectUserPassword = document.ProtectConfirmPassword = "reader";
                    document.ProtectOwnerPassword = "owner";
                    Task save = document.SaveProtectedCopyCommand.ExecuteAsync(null);
                    await ChooseName(view, window, save, services, "protected.pdf", () => document.ErrorMessage);
                    if (existing) Click(window, await WaitForDialog<MobileReplaceFileDialogContent>(view, save, () => document.ErrorMessage), services.Localizer.Get("Common.Save"));
                    var review = await WaitForDialog<PdfProtectionDialogContent>(view, save, () => document.ErrorMessage);
                    Assert.Empty(window.OwnedWindows);
                    Click(window, review, services.Localizer.Get("Protection.Create"));
                    await WaitUntil(() => view.FindControl<ContentControl>("DialogContent")!.Content is PdfProtectionDialogContent {
                        DataContext: PdfProtectionPreviewViewModel { HasResult: true }
                    }, save, () => document.ErrorMessage);
                    Assert.Equal(3, PdfDocument.Load(new System.IO.MemoryStream(File.ReadAllBytes(outputPath)), new PdfLoadOptions { Password = "reader" }).Inspect().Pages.Count);
                    Assert.Equal(original, File.ReadAllBytes(document.DocumentPath!));
                    Click(window, (Control)view.FindControl<ContentControl>("DialogContent")!.Content!, services.Localizer.Get("Common.Close"));
                    await save;
                    Assert.False(document.HasError, document.ErrorMessage);
                    Task open = document.OpenCommand.ExecuteAsync(null);
                    var password = await WaitForDialog<PdfPasswordDialogContent>(view, open, () => controller.Document.ErrorMessage);
                    Assert.Single(password.GetVisualDescendants().OfType<TextBox>()).Text = "reader";
                    Click(window, password, services.Localizer.Get("Common.Open"));
                    await open;
                    Assert.Equal(2, controller.Tabs.Tabs.Count);
                    Assert.Equal(3, controller.Document.Pages.Count);
                    Assert.True(controller.Document.IsDocumentEncrypted);
                    Assert.False(controller.Document.HasError, controller.Document.ErrorMessage);
                    controller.Dispose();
                    var restartedServices = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                        new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                    using var restarted = new MobileDocumentController(restartedServices, _ => Task.FromResult<IStorageFile?>(null),
                        _ => Task.CompletedTask, new MobileDocumentHost(restartedServices, view, _ => Task.CompletedTask, () => provider));
                    view.Connect(restarted);
                    view.SetInitializing(true);
                    Task restore = restarted.RestoreAsync();
                    password = await WaitForDialog<PdfPasswordDialogContent>(view, restore, () => restarted.Document.ErrorMessage);
                    window.UpdateLayout();
                    Assert.True(password.IsEffectivelyEnabled);
                    Assert.Single(password.GetVisualDescendants().OfType<TextBox>()).Text = "reader";
                    Click(window, password, services.Localizer.Get("Common.Open"));
                    await restore;
                    view.SetInitializing(false);
                    Assert.Equal(2, restarted.Tabs.Tabs.Count);
                    Assert.True(restarted.Document.IsDocumentEncrypted);
                    restartedServices.Storage.Dispose();
                } finally { window.Close(); services.Storage.Dispose(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Fact]
    public async Task CancellingTheFileNameDoesNotOpenAPickerOrModifyTheDestination() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            string root = ((App)Application.Current!).Services.Paths.Root;
            var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "mobile-consent")),
                new StudioLocalDocumentRoot(Path.Combine(root, "mobile-documents")));
            var output = new TestStorageFile("content://mobile/existing.pdf", StudioProviderDocumentTests.CreatePdf(2), "existing.pdf");
            byte[] previous = output.Bytes.ToArray();
            int selections = 0;
            var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
            ((TestStorageFile.StorageProxy)(object)provider).Call = (method, _) => {
                Assert.Equal("SaveFilePickerAsync", method);
                selections++;
                return Task.FromResult<IStorageFile?>(output.Item);
            };
            var view = new MobileWorkspaceView();
            using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null),
                _ => Task.CompletedTask, new MobileDocumentHost(services, view, _ => Task.CompletedTask, () => provider));
            view.Connect(controller);
            var window = new Window { Content = view, Width = 1024, Height = 768 };
            try {
                window.Show();
                await controller.OpenSampleAsync();
                Task saving = controller.Document.SaveAsCommand.ExecuteAsync(null);
                var confirmation = await WaitForDialog<MobileFileNameDialogContent>(view, saving, () => controller.Document.ErrorMessage);
                Click(window, confirmation, services.Localizer.Get("Common.Cancel"));
                await saving;
                Assert.Equal(0, selections);
                Assert.Equal(previous, output.Bytes);
            } finally { window.Close(); services.Storage.Dispose(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellingProtectionOrRejectingAnOwnedDestinationPreservesExistingBytes(bool owned) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            string root = ((App)Application.Current!).Services.Paths.Root;
            var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "destination-data")),
                new StudioLocalDocumentRoot(Path.Combine(root, "destination-documents")));
            string folderPath = Path.Combine(services.Paths.Root, "destinations");
            Directory.CreateDirectory(folderPath);
            string target = Path.Combine(folderPath, "existing.pdf");
            byte[] original = StudioProviderDocumentTests.CreatePdf(2);
            File.WriteAllBytes(target, original);
            IStorageFolder folder = TestLocalOutputFolder.Create(folderPath);
            var targetItem = new TestStorageFile(new Uri(target).AbsoluteUri, original, "existing.pdf");
            int exportPickers = 0;
            var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
            ((TestStorageFile.StorageProxy)(object)provider).Call = (method, _) => {
                if (method == "OpenFolderPickerAsync") return Task.FromResult<IReadOnlyList<IStorageFolder>>([folder]);
                Assert.Equal("SaveFilePickerAsync", method);
                exportPickers++;
                File.WriteAllBytes(target, []); // UIKit publishes its placeholder during selection.
                return Task.FromResult<IStorageFile?>(targetItem.Item);
            };
            var view = new MobileWorkspaceView();
            using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(targetItem.Item),
                _ => Task.CompletedTask, new MobileDocumentHost(services, view, _ => Task.CompletedTask, () => provider));
            view.Connect(controller);
            var window = new Window { Content = view, Width = 1024, Height = 768 };
            try {
                window.Show();
                await controller.OpenSampleAsync();
                var first = Assert.Single(controller.Tabs.Tabs);
                if (owned) {
                    await controller.Document.OpenCommand.ExecuteAsync(null);
                    target = controller.Document.DocumentPath!;
                    folder = TestLocalOutputFolder.Create(Path.GetDirectoryName(target)!);
                    targetItem = new TestStorageFile(new Uri(target).AbsoluteUri, original, "existing.pdf");
                    controller.Tabs.SelectedTab = first;
                }
                var document = controller.Document;
                document.ProtectUserPassword = document.ProtectConfirmPassword = "reader";
                Task operation = (owned ? document.SaveAsCommand : document.SaveProtectedCopyCommand).ExecuteAsync(null);
                // Supports the pre-fix consent/picker path to reproduce its side effect as well.
                await WaitUntil(() => view.FindControl<ContentControl>("DialogContent")!.Content is MobileFileNameDialogContent or ProviderSaveDialogContent,
                    operation, () => document.ErrorMessage);
                if (view.FindControl<ContentControl>("DialogContent")!.Content is ProviderSaveDialogContent oldConsent)
                    Click(window, oldConsent, services.Localizer.Get("Common.Save"));
                else {
                    await ChooseName(view, window, operation, services, "existing.pdf", () => document.ErrorMessage);
                    var replacement = await WaitForDialog<MobileReplaceFileDialogContent>(view, operation, () => document.ErrorMessage);
                    if (!owned) {
                        window.UpdateLayout();
                        using var frame = window.CaptureRenderedFrame();
                        Assert.NotNull(frame);
                        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                        if (!string.IsNullOrWhiteSpace(output)) {
                            Directory.CreateDirectory(output);
                            frame.Save(Path.Combine(output, "mobile-safe-replacement.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                            window.Width = 390; window.Height = 844; window.UpdateLayout();
                            using var compact = window.CaptureRenderedFrame();
                            Assert.NotNull(compact);
                            compact.Save(Path.Combine(output, "mobile-safe-replacement-phone.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                        }
                    }
                    Click(window, replacement, services.Localizer.Get("Common.Save"));
                }
                if (!owned) {
                    var preview = await WaitForDialog<PdfProtectionDialogContent>(view, operation, () => document.ErrorMessage);
                    Click(window, preview, services.Localizer.Get("Common.Cancel"));
                }
                await operation;
                Assert.Equal(original, File.ReadAllBytes(target));
                Assert.Equal(0, exportPickers);
                Assert.Empty(services.Jobs.Entries);
                if (owned) Assert.Contains("already open", document.OperationStatus, StringComparison.OrdinalIgnoreCase);
            } finally { window.Close(); controller.Dispose(); services.Storage.Dispose(); }
            return true;
        }, CancellationToken.None);
    }

    private static async Task ChooseName(MobileWorkspaceView view, Window window, Task operation,
        StudioApplicationServices services, string name, Func<string?> error) {
        var dialog = await WaitForDialog<MobileFileNameDialogContent>(view, operation, error);
        Assert.Single(dialog.GetVisualDescendants().OfType<TextBox>()).Text = name;
        Click(window, dialog, services.Localizer.Get("Mobile.ChooseFolder"));
    }

    private static async Task<T> WaitForDialog<T>(MobileWorkspaceView view, Task operation, Func<string?> error) where T : Control {
        await WaitUntil(() => view.FindControl<ContentControl>("DialogContent")!.Content is T, operation, error);
        return (T)view.FindControl<ContentControl>("DialogContent")!.Content!;
    }

    private static async Task WaitUntil(Func<bool> ready, Task operation, Func<string?> error) {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        while (!ready()) {
            Assert.False(operation.IsCompleted, error() ?? "The operation ended before presenting its review.");
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10, timeout.Token);
        }
    }

    private static void Click(Window window, Control dialog, string label) {
        window.UpdateLayout();
        Dispatcher.UIThread.RunJobs();
        var button = Assert.Single(dialog.GetVisualDescendants().OfType<Button>(), button => Equals(button.Content, label));
        Assert.True(button.IsEffectivelyEnabled);
        button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
    }
}
