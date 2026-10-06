using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Avalonia binding and command adapter over the shared document-tab lifecycle.</summary>
public sealed partial class StudioDocumentTabHost : ObservableObject, IDisposable {
    internal StudioDocumentTabs<MainWindowViewModel, StudioDocumentTabViewModel> Core { get; }
    internal StudioDocumentTabHost(Func<Func<string, CancellationToken, Task>, MainWindowViewModel> createDocument,
        Action<MainWindowViewModel> activateDocument, Func<MainWindowViewModel, Task<bool>>? prepareActiveClose = null) {
        Core = new(createDocument, activateDocument, (document, close) => new(document, close), prepareActiveClose);
        Core.PropertyChanged += (_, args) => OnPropertyChanged(args.PropertyName);
    }
    public ObservableCollection<StudioDocumentTabViewModel> Tabs => Core.Tabs;
    public StudioDocumentTabViewModel? SelectedTab { get => Core.SelectedTab; set => Core.SelectedTab = value; }
    public bool HasTabs => Core.HasTabs;
    internal MainWindowViewModel ActiveDocument => Core.ActiveDocument;
    internal IEnumerable<MainWindowViewModel> OperationDocuments => Core.OperationDocuments;
    internal bool HasBusyDocuments => Core.HasBusyDocuments;
    internal bool HasDirtyDocuments => Core.HasDirtyDocuments;
    internal bool CanReopenClosedTab => Core.CanReopenClosedTab;
    internal event EventHandler? CloseAllPrepared { add => Core.CloseAllPrepared += value; remove => Core.CloseAllPrepared -= value; }
    [RelayCommand] private Task OpenNewTabAsync() => ActiveDocument.OpenCommand.ExecuteAsync(null);
    internal Task CloseSelectedTabAsync() => Core.CloseSelectedTabAsync();
    internal void SelectRelativeTab(bool previous) => Core.SelectRelativeTab(previous);
    internal void MoveTab(StudioDocumentTabViewModel tab, int index) => Core.MoveTab(tab, index);
    internal bool CanPublishPath(string path) => Core.CanPublishPath(path);
    internal bool CanPublishBookPath(MainWindowViewModel? owner, string path) => Core.CanPublishBookPath(owner, path);
    internal bool CanPublishDirectory(string path) => Core.CanPublishDirectory(path);
    internal bool CanDocumentOwnPath(MainWindowViewModel? document, string path, MainWindowViewModel? ignoreBookOwner = null) =>
        Core.CanDocumentOwnPath(document, path, ignoreBookOwner);
    internal Task OpenDocumentAsync(string path, CancellationToken cancellationToken = default) => Core.OpenDocumentAsync(path, cancellationToken);
    internal Task ReopenClosedTabAsync() => Core.ReopenClosedTabAsync();
    internal Task CloseTabAsync(StudioDocumentTabViewModel tab) => Core.CloseTabAsync(tab);
    internal Task<bool> RequestCloseAllAsync() => Core.RequestCloseAllAsync();
    internal void CancelAllOperations() => Core.CancelAllOperations();
    public void Dispose() => Core.Dispose();
}
