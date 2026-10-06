using CommunityToolkit.Mvvm.ComponentModel;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Represents one live document workspace in a document tab strip.</summary>
internal sealed partial class StudioDocumentTabState<TDocument> : ObservableObject, IDisposable where TDocument : class, IStudioDocument {
    private bool _disposed;

    internal StudioDocumentTabState(
        TDocument document) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        _title = document.DocumentName;
        Document.Changed += OnDocumentChanged;
    }

    internal TDocument Document { get; }

    public override string ToString() => Title;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DisplayTitle))]
    private string _title;

    /// <summary>The document name without the unsaved-changes marker; the tab shows a dot instead.</summary>
    public string DisplayTitle => Title.TrimEnd(' ', '*');

    public bool IsDirty => Document.IsDirty;

    public string? SourcePath => Document.DocumentPath;

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        Document.Changed -= OnDocumentChanged;
        Document.Dispose();
    }

    private void OnDocumentChanged(object? sender, StudioDocumentChangedEventArgs e) {
        if (e.Kind == StudioDocumentChangeKind.Name) Title = Document.DocumentName;
        else if (e.Kind == StudioDocumentChangeKind.Content) OnPropertyChanged(nameof(IsDirty));
        else if (e.Kind == StudioDocumentChangeKind.Document) OnPropertyChanged(nameof(SourcePath));
    }
}
