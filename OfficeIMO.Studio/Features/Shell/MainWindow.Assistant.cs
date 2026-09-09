using System.Text;
using Avalonia.Input.Platform;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Features.Shell;

public partial class MainWindow {
    private void ConfigureAssistantHost(MainWindowViewModel document) {
        document.Assistant.CopyAnswer = async text => {
            var clipboard = Clipboard ?? throw new IOException("Clipboard is unavailable.");
            await clipboard.SetTextAsync(text);
        };
        document.Assistant.ExportAnswer = (text, isCurrent, token) => ExportAssistantAnswerAsync(text, isCurrent, token);
        document.OcrWorkbench.OpenAssistantOutput = async (path, token) => {
            await TabHost.OpenDocumentAsync(path, token);
            MainWindowViewModel opened = TabHost.ActiveDocument;
            if (opened.HasDocument && opened.DocumentPath is { } openedPath && OfficeStorageIdentity.AreEquivalent(openedPath, path)) {
                opened.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
                opened.IsAssistantVisible = true;
                opened.Assistant.ShowConnections = false;
            }
        };
    }

    private async Task<bool> ExportAssistantAnswerAsync(string text, Func<bool> isCurrent, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (!isCurrent() || !StorageProvider.CanSave) return false;
        IStorageFile? file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = _services.Localizer.GetOrDefault("Assistant.ExportAnswer", "Export answer"),
            SuggestedFileName = "document-answer.txt", DefaultExtension = "txt",
            FileTypeChoices = [new FilePickerFileType(_services.Localizer.GetOrDefault("Assistant.TextFile", "Text document")) {
                Patterns = ["*.txt"], MimeTypes = ["text/plain"], AppleUniformTypeIdentifiers = ["public.plain-text"]
            }]
        });
        string? location = await _services.Storage.RegisterSingleAsync(file is null ? [] : [file], token);
        if (location is null || !isCurrent()) return false;
        void Authorize() {
            token.ThrowIfCancellationRequested();
            _services.Storage.EnsureWritableLocation(location);
            if (!isCurrent() || !TabHost.CanPublishPath(location)) throw new IOException("The source or destination changed. Export the current answer to a separate file.");
        }
        async Task AuthorizeProvider(CancellationToken cancellation) {
            cancellation.ThrowIfCancellationRequested();
            await Dispatcher.UIThread.InvokeAsync(Authorize, DispatcherPriority.Normal, cancellation);
        }
        Authorize();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        if (_services.Storage.UsesProviderPublication(location)) {
            if (!await ConfirmProviderWriteAsync(location)) return false;
            await _services.Storage.PublishAsync(location, bytes, expectedFingerprint: null, AuthorizeProvider, token);
        } else {
            await StudioLocalPublication.WriteAsync(location, bytes, Authorize, token);
        }
        return true;
    }
}
