using OfficeIMO.Studio.Infrastructure;
using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Connects the shared restart controller to Avalonia application services.</summary>
internal sealed partial class StudioSessionController : StudioDocumentSessionController<MainWindowViewModel, StudioDocumentTabViewModel> {
    internal StudioSessionController(StudioDocumentTabHost host, StudioApplicationServices services,
        Func<CancellationToken, Task<string?>> pickCopy)
        : base(host.Core, new StudioSessionEnvironment(services), pickCopy, new AvaloniaStudioScheduler()) { }
    [RelayCommand] private Task RestoreAsync() => RestoreDocumentsAsync();
    [RelayCommand] private Task OpenCurrentAsync(StudioSessionItem? item) => OpenCurrentDocumentAsync(item);
    [RelayCommand] private Task RecoverCopyAsync(StudioSessionItem? item) => RecoverDocumentCopyAsync(item);
    [RelayCommand] private void Forget(StudioSessionItem? item) => ForgetDocument(item);
    [RelayCommand] private void RetryStorage() => Flush();
}
