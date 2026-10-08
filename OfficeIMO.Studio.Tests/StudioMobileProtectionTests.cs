using System.Reflection;
using Avalonia;
using Avalonia.Controls;
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
    public async Task SavedProviderCopySurvivesRestartAndSharesThroughALocalSnapshot(bool existing) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-save-as-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var paths = new StudioDataPaths(Path.Combine(root, "data"));
                var documents = new StudioLocalDocumentRoot(Path.Combine(root, "documents"));
                var services = StudioApplicationServices.Create(paths, documents);
                var output = new TestStorageFile("content://mobile/saved.pdf", [], "saved.pdf");
                var folder = new StudioProviderOutputFolderTests.OutputFolder(hierarchical: true);
                if (existing) folder.Files[output.Name] = output;
                var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
                ((TestStorageFile.StorageProxy)(object)provider).Call = (method, _) => method switch {
                    "OpenFolderPickerAsync" => Task.FromResult<IReadOnlyList<IStorageFolder>>([folder.Item]),
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
                    await ChooseName(window, view, saving, output.Name);
                    var confirmation = await WaitForDialog<ProviderSaveDialogContent>(view, saving, () => controller.Document.ErrorMessage);
                    Click(window, confirmation, services.Localizer.Get("Common.Save"));
                    await saving;
                    output = folder.Files["saved.pdf"];
                    Assert.Equal(existing ? 0 : 1, folder.Creations);
                    Assert.False(controller.Document.HasError, controller.Document.ErrorMessage);
                    Assert.Equal(output.Location.AbsoluteUri, controller.Document.DocumentPath);
                    Assert.Equal(original, File.ReadAllBytes(originalPath));
                    controller.Dispose();
                    services.Storage.Dispose();
                    var restartedServices = StudioApplicationServices.Create(paths, documents);
                    restartedServices.Storage.Attach(() => output.CreateProvider());
                    string? sharedPath = null;
                    using var restarted = new MobileDocumentController(restartedServices, _ => Task.FromResult<IStorageFile?>(null), path => {
                        sharedPath = path;
                        Assert.StartsWith(documents.Path + Path.DirectorySeparatorChar, path);
                        Assert.Equal(output.Bytes, File.ReadAllBytes(path));
                        return Task.CompletedTask;
                    }, new MobileDocumentHost(restartedServices, view, _ => Task.CompletedTask, () => provider));
                    view.Connect(restarted);
                    await restarted.RestoreAsync();
                    Assert.True(restarted.Document.HasDocument, restarted.Document.ErrorMessage);
                    Assert.Equal(output.Location.AbsoluteUri, restarted.Document.DocumentPath);
                    await restarted.ShareAsync();
                    Assert.NotNull(sharedPath);
                    Assert.False(File.Exists(sharedPath));
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
                var output = new TestStorageFile("content://mobile/protected.pdf", [], "protected.pdf");
                var folder = new StudioProviderOutputFolderTests.OutputFolder(hierarchical: true);
                if (existing) folder.Files[output.Name] = output;
                var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
                ((TestStorageFile.StorageProxy)(object)provider).Call = (method, _) => method switch {
                    "OpenFolderPickerAsync" => Task.FromResult<IReadOnlyList<IStorageFolder>>([folder.Item]),
                    _ => throw new NotSupportedException(method)
                };
                var view = new MobileWorkspaceView();
                var host = new MobileDocumentHost(services, view, _ => Task.CompletedTask, () => provider);
                using var controller = new MobileDocumentController(services,
                    _ => Task.FromResult<IStorageFile?>(new TestStorageFile("content://mobile/import.pdf", output.Bytes, "protected.pdf").Item),
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
                    await ChooseName(window, view, save, output.Name);
                    var confirm = await WaitForDialog<ProviderSaveDialogContent>(view, save, () => document.ErrorMessage);
                    Click(window, confirm, services.Localizer.Get("Common.Save"));
                    var review = await WaitForDialog<PdfProtectionDialogContent>(view, save, () => document.ErrorMessage);
                    Assert.Empty(window.OwnedWindows);
                    Click(window, review, services.Localizer.Get("Protection.Create"));
                    var publication = await WaitForDialog<ProviderSaveDialogContent>(view, save, () => document.ErrorMessage);
                    Click(window, publication, services.Localizer.Get("Common.Save"));
                    await WaitUntil(() => view.FindControl<ContentControl>("DialogContent")!.Content is PdfProtectionDialogContent {
                        DataContext: PdfProtectionPreviewViewModel { HasResult: true }
                    }, save, () => document.ErrorMessage);
                    output = folder.Files["protected.pdf"];
                    Assert.Equal(existing ? 0 : 1, folder.Creations);
                    Assert.Equal(3, PdfDocument.Load(output.Bytes, new PdfLoadOptions { Password = "reader" }).Inspect().Pages.Count);
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

    private static async Task ChooseName(Window window, MobileWorkspaceView view, Task operation, string name) {
        var dialog = await WaitForDialog<MobileFileNameDialogContent>(view, operation, () => null);
        window.UpdateLayout();
        Assert.Single(dialog.GetVisualDescendants().OfType<TextBox>()).Text = name;
        Click(window, dialog, "Choose folder");
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
