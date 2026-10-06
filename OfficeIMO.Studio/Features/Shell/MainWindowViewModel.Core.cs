namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel : IStudioDocument {
    private event EventHandler<StudioDocumentChangedEventArgs>? CoreDocumentChanged;
    event EventHandler<StudioDocumentChangedEventArgs>? IStudioDocument.Changed {
        add => CoreDocumentChanged += value;
        remove => CoreDocumentChanged -= value;
    }
    private void NotifyCoreDocumentChanged(System.ComponentModel.PropertyChangedEventArgs args) {
        StudioDocumentChangeKind? kind = args.PropertyName switch {
            nameof(DocumentName) => StudioDocumentChangeKind.Name,
            nameof(DocumentPath) or nameof(HasDocument) => StudioDocumentChangeKind.Document,
            nameof(IsDirty) => StudioDocumentChangeKind.Content,
            nameof(SelectedPage) or nameof(Zoom) or nameof(IsFocusReading) => StudioDocumentChangeKind.View,
            _ => null
        };
        if (kind is { } changed) CoreDocumentChanged?.Invoke(this, new(changed));
    }
    string? IStudioDocument.DocumentPath => DocumentPath;
    string? IStudioDocument.LastClosedDocumentPath => LastClosedDocumentPath;
    bool IStudioDocument.HasUnsavedAuxiliaryWork => BookWorkbench.IsDirty;
    IEnumerable<string> IStudioDocument.OwnedOutputLocations => BookWorkbench.OwnedLocations;
    bool IStudioDocument.OwnsOutputPath(string path) => BookWorkbench.OwnsPath(path);
    void IStudioDocument.Deactivate() => DeactivateAssistant();
    void IStudioDocument.SetPresentationActive(bool active) => SetPresentationActive(active);
    Task IStudioDocument.OpenDocumentAsync(string path, CancellationToken token) => OpenDocumentAsync(path, token);
    Task<bool> IStudioDocument.RequestCloseDocumentAsync() => RequestCloseDocumentAsync();
    Task<bool> IStudioDocument.PrepareCloseDocumentAsync() => PrepareCloseDocumentAsync();
    Task<bool> IStudioDocument.CommitPreparedDiscardAsync() => CommitPreparedDiscardAsync();
    void IStudioDocument.CancelPreparedClose() => CancelPreparedClose();
    void IStudioDocument.CompletePreparedClose() => CompletePreparedClose();
    void IStudioDocument.CancelCurrentOperation() => CancelCurrentOperation();
    Infrastructure.Preferences.StudioSessionDocument? IStudioDocument.CaptureSessionDocument() => CaptureSessionDocument();
    void IStudioDocument.RestoreSessionViewState(Infrastructure.Preferences.StudioDocumentViewState state) => RestoreSessionViewState(state);
}
