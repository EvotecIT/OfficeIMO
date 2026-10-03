using Avalonia;
using Avalonia.iOS;
using Foundation;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Shell;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using UIKit;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.iOS;

// UIKit creates this type through Objective-C rather than a managed constructor call.
[Register("AppDelegate")]
public sealed class AppDelegate : AvaloniaAppDelegate<App> {
    private MobileDocumentController? _documents;
    private MobileWorkspaceView? _workspace;
    private readonly List<NSObject> _observers = [];

    protected override AppBuilder CreateAppBuilder() => AppBuilder.Configure(() => new App(
        StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(
                NSSearchPath.GetDirectories(NSSearchPathDirectory.ApplicationSupportDirectory, NSSearchPathDomain.User, true)[0], "OfficeIMO", "Studio")),
            new StudioLocalDocumentRoot(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Studio")))) {
        SingleViewFactory = CreateWorkspace
    }).UseiOS(this);

    private Control CreateWorkspace(App app) {
        _workspace = new MobileWorkspaceView();
        _documents = new MobileDocumentController(app.Services, PickPdfAsync, ShareAsync);
        _workspace.Connect(_documents);
        _workspace.ShareDocumentAsync = _documents.ShareAsync;
        _workspace.OpenSampleAsync = _documents.OpenSampleAsync;
        TopLevel.SetAutoSafeAreaPadding(_workspace, true);
        _workspace.Loaded += OnWorkspaceLoaded;
        _observers.Add(NSNotificationCenter.DefaultCenter.AddObserver(UIApplication.DidEnterBackgroundNotification, _ => _documents.Suspend()));
        _observers.Add(NSNotificationCenter.DefaultCenter.AddObserver(UIApplication.WillEnterForegroundNotification, _ => _documents.Resume()));
        _observers.Add(NSNotificationCenter.DefaultCenter.AddObserver(UIApplication.DidReceiveMemoryWarningNotification, _ => {
            _documents.Suspend();
            if (UIApplication.SharedApplication.ApplicationState == UIApplicationState.Active) _documents.Resume();
        }));
        return _workspace;
    }

    private async void OnWorkspaceLoaded(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        _workspace!.Loaded -= OnWorkspaceLoaded;
        _workspace.IsEnabled = false;
        try { await _documents!.RestoreAsync(); }
        catch (Exception error) { _documents!.Document.ErrorMessage = error.Message; }
        finally { _workspace.IsEnabled = true; }
    }

    private async Task<IStorageFile?> PickPdfAsync(CancellationToken token) {
        IStorageProvider provider = TopLevel.GetTopLevel(_workspace!)!.StorageProvider;
        var files = await provider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = "Open a PDF", AllowMultiple = false, FileTypeFilter = [new FilePickerFileType("PDF document") { AppleUniformTypeIdentifiers = ["com.adobe.pdf"] }]
        });
        if (token.IsCancellationRequested) { foreach (var file in files) file.Dispose(); token.ThrowIfCancellationRequested(); }
        foreach (var file in files.Skip(1)) file.Dispose();
        return files.FirstOrDefault();
    }

    private static async Task ShareAsync(string path) {
        using NSUrl url = NSUrl.FromFilename(path);
        using var activity = new UIActivityViewController([url], null);
        UIViewController presenter = UIApplication.SharedApplication.ConnectedScenes.OfType<UIWindowScene>()
            .Where(scene => scene.ActivationState == UISceneActivationState.ForegroundActive)
            .SelectMany(scene => scene.Windows).First(window => window.IsKeyWindow).RootViewController!;
        while (presenter.PresentedViewController is { } presented) presenter = presented;
        if (activity.PopoverPresentationController is { } popover) {
            popover.SourceView = presenter.View!;
            var bounds = presenter.View!.Bounds;
            popover.SourceRect = new CoreGraphics.CGRect(bounds.Width - 60, 60, 1, 1);
        }
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        activity.CompletionWithItemsHandler = (_, _, _, error) => {
            if (error is not null) completion.TrySetException(new IOException(error.LocalizedDescription));
            else completion.TrySetResult();
        };
        await presenter.PresentViewControllerAsync(activity, true);
        await completion.Task;
    }

    protected override void Dispose(bool disposing) {
        if (disposing) {
            foreach (var observer in _observers) observer.Dispose();
            _documents?.Dispose();
        }
        base.Dispose(disposing);
    }
}
