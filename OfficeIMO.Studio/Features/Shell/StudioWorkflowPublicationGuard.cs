using Avalonia.Threading;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Evaluates live tab ownership on the UI thread at the shared runner's publication boundary.</summary>
internal sealed class StudioWorkflowPublicationGuard(Func<string, bool, bool> canPublish) : IOfficeWorkflowPublicationGuard {
    public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (Dispatcher.UIThread.CheckAccess()) return canPublish(path, isDirectory);
        return await Dispatcher.UIThread.InvokeAsync(() => {
            cancellationToken.ThrowIfCancellationRequested();
            return canPublish(path, isDirectory);
        }, DispatcherPriority.Normal, cancellationToken);
    }
}
