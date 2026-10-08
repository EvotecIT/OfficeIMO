using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>The document lifecycle consumed by the shared tab and restart controllers.</summary>
internal interface IStudioDocument : IDisposable {
    event EventHandler<StudioDocumentChangedEventArgs>? Changed;
    string DocumentName { get; }
    string? DocumentPath { get; }
    string? LastClosedDocumentPath { get; }
    bool HasDocument { get; }
    bool IsDirty { get; }
    bool HasUnsavedAuxiliaryWork { get; }
    bool CanCancelOperation { get; }
    string? ErrorMessage { get; set; }
    IEnumerable<string> OwnedOutputLocations { get; }
    bool OwnsOutputPath(string path);
    void Deactivate();
    void SetPresentationActive(bool active);
    Task OpenDocumentAsync(string path, CancellationToken token);
    Task<bool> RequestCloseDocumentAsync();
    Task<bool> PrepareCloseDocumentAsync();
    Task<bool> CommitPreparedDiscardAsync();
    void CancelPreparedClose();
    void CompletePreparedClose();
    void CancelCurrentOperation();
    StudioSessionDocument? CaptureSessionDocument();
    void RestoreSessionViewState(StudioDocumentViewState state);
}

internal interface IStudioDocumentTab<out TDocument> : IDisposable where TDocument : IStudioDocument {
    TDocument Document { get; }
    string Title { get; set; }
}

/// <summary>The state changes relevant to shared tab and restart operations.</summary>
internal enum StudioDocumentChangeKind { Name, Document, Content, View }
internal sealed class StudioDocumentChangedEventArgs(StudioDocumentChangeKind kind) : EventArgs {
    internal StudioDocumentChangeKind Kind { get; } = kind;
}
