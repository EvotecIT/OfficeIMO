using Avalonia.Threading;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Evaluates live tab ownership on the UI thread at the shared runner's publication boundary.</summary>
internal sealed class StudioWorkflowPublicationGuard(Func<string, bool, bool> canPublish) : IOfficeWorkflowPublicationGuard {
    // Background completions belong to this application, even after its dispatcher shuts down.
    private readonly Avalonia.Threading.Dispatcher _uiDispatcher = Avalonia.Threading.Dispatcher.UIThread;
    public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_uiDispatcher.CheckAccess()) return canPublish(path, isDirectory);
        return await _uiDispatcher.InvokeAsync(() => {
            cancellationToken.ThrowIfCancellationRequested();
            return canPublish(path, isDirectory);
        }, DispatcherPriority.Normal, cancellationToken);
    }
}
