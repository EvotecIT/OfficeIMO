using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Diagnostics;

namespace OfficeIMO.Studio;

public sealed partial class App {
    private void AttachFileActivation(IClassicDesktopStyleApplicationLifetime desktop, MainWindow window) {
        if (TryGetFeature(typeof(IActivatableLifetime)) is not IActivatableLifetime activation) return;
        activation.Activated += OnActivated;
        desktop.Exit += (_, _) => activation.Activated -= OnActivated;

        async void OnActivated(object? sender, ActivatedEventArgs args) {
            if (args is not FileActivatedEventArgs files) return;
            try {
                Services.Diagnostics.Write(StudioDiagnosticLevel.Information, "FileActivation", "Received");
                if (Dispatcher.UIThread.CheckAccess()) await window.OpenActivatedItemsAsync(files.Files);
                else await Dispatcher.UIThread.InvokeAsync(() => window.OpenActivatedItemsAsync(files.Files));
            } catch (Exception error) when (error is not OutOfMemoryException) {
                Services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "FileActivation", "DispatchFailed", error);
            }
        }
    }
}
