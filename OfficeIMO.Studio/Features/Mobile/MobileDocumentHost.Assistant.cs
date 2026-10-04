using System.Text;
using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Threading;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Mobile;

internal sealed partial class MobileDocumentHost {
    internal void ConfigureAssistant(MainWindowViewModel document, Func<string, bool> canPublish) {
        services.AiConnections.OpenUri = OpenUriAsync;
        document.Assistant.CopyAnswer = async text => {
            var clipboard = TopLevel.GetTopLevel(workspace)?.Clipboard ?? throw new IOException("Clipboard is unavailable.");
            await clipboard.SetTextAsync(text);
        };
        document.Assistant.ExportAnswer = async (text, isCurrent, token) => {
            token.ThrowIfCancellationRequested();
            if (!isCurrent()) return false;
            string? location = await PickSaveFileAsync(services.Localizer.GetOrDefault("Assistant.ExportAnswer", "Export answer"),
                "document-answer.txt", new("Text document", ["txt"], "text/plain"), token);
            if (location is null || !isCurrent()) return false;
            void Authorize() {
                token.ThrowIfCancellationRequested();
                services.Storage.EnsureWritableLocation(location);
                if (!isCurrent() || !canPublish(location)) throw new IOException("The source or destination changed. Export the current answer to a separate file.");
            }
            Authorize();
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            if (services.Storage.UsesProviderPublication(location)) {
                await services.Storage.PublishAsync(location, bytes, expectedFingerprint: null,
                    async cancellation => await Dispatcher.UIThread.InvokeAsync(Authorize, DispatcherPriority.Normal, cancellation), token);
            } else {
                await StudioLocalPublication.WriteAsync(location, bytes, Authorize, token);
            }
            return true;
        };
    }
}
