using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Owns live document view-models and coordinates tab activation and close semantics.</summary>
public sealed partial class StudioDocumentTabHost : ObservableObject, IDisposable {
    private readonly Func<Func<string, CancellationToken, Task>, MainWindowViewModel> _createDocument;
    private readonly Action<MainWindowViewModel> _activateDocument;
    private MainWindowViewModel _emptyDocument;
    private bool _openingDocument;
    private bool _disposed;

    internal StudioDocumentTabHost(
        Func<Func<string, CancellationToken, Task>, MainWindowViewModel> createDocument,
        Action<MainWindowViewModel> activateDocument) {
        _createDocument = createDocument ?? throw new ArgumentNullException(nameof(createDocument));
        _activateDocument = activateDocument ?? throw new ArgumentNullException(nameof(activateDocument));
        _emptyDocument = _createDocument(OpenDocumentAsync);
    }

    public ObservableCollection<StudioDocumentTabViewModel> Tabs { get; } = new();

    internal event EventHandler? CloseAllPrepared;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasTabs))]
    private StudioDocumentTabViewModel? _selectedTab;

    public bool HasTabs => Tabs.Count > 0;

    internal MainWindowViewModel ActiveDocument => SelectedTab?.Document ?? _emptyDocument;

    internal bool HasBusyDocuments => Tabs.Any(tab => tab.Document.CanCancelOperation) ||
                                      _emptyDocument.CanCancelOperation;

    internal bool HasDirtyDocuments => Tabs.Any(tab => tab.Document.IsDirty);

    partial void OnSelectedTabChanged(StudioDocumentTabViewModel? value) =>
        _activateDocument(value?.Document ?? _emptyDocument);

    [RelayCommand]
    private Task OpenNewTabAsync() => ActiveDocument.OpenCommand.ExecuteAsync(null);

    internal Task CloseSelectedTabAsync() => SelectedTab is null
        ? Task.CompletedTask
        : CloseTabAsync(SelectedTab);

    internal void SelectRelativeTab(bool previous) {
        if (Tabs.Count < 2 || SelectedTab is null) return;
        int current = Tabs.IndexOf(SelectedTab);
        int offset = previous ? -1 : 1;
        SelectedTab = Tabs[(current + offset + Tabs.Count) % Tabs.Count];
    }

    internal bool CanPublishPath(string path) => CanDocumentOwnPath(null, path);

    internal bool CanPublishDirectory(string path) {
        if (string.IsNullOrWhiteSpace(path)) return false;
        try {
            return Tabs.All(tab => tab.Document.DocumentPath is not { Length: > 0 } source ||
                !OfficePathIdentity.IsSameOrDescendant(source, path));
        } catch (Exception exception) when (IsPathIdentityFailure(exception)) {
            return false;
        }
    }

    internal bool CanDocumentOwnPath(MainWindowViewModel? document, string path) {
        if (string.IsNullOrWhiteSpace(path)) return false;
        try {
            string fullPath = Path.GetFullPath(path);
            return Tabs.All(tab => ReferenceEquals(tab.Document, document) ||
                !DocumentOwnsPath(tab.Document, fullPath));
        } catch (Exception exception) when (IsPathIdentityFailure(exception)) {
            // An uninspectable destination cannot safely be authorized for publication.
            return false;
        }
    }

    internal async Task OpenDocumentAsync(string path, CancellationToken cancellationToken = default) {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_openingDocument || string.IsNullOrWhiteSpace(path)) return;

        string fullPath;
        StudioDocumentTabViewModel? existing;
        try {
            fullPath = Path.GetFullPath(path);
            existing = Tabs.FirstOrDefault(tab =>
                DocumentOwnsPath(tab.Document, fullPath));
        } catch (Exception exception) when (IsPathIdentityFailure(exception)) {
            ActiveDocument.ErrorMessage = exception.Message;
            return;
        }
        if (existing is not null) {
            SelectedTab = existing;
            return;
        }

        _openingDocument = true;
        StudioDocumentTabViewModel? previousTab = SelectedTab;
        bool reusedEmptyDocument = Tabs.Count == 0 && !_emptyDocument.HasDocument;
        MainWindowViewModel candidate = reusedEmptyDocument
            ? _emptyDocument
            : _createDocument(OpenDocumentAsync);
        var tab = new StudioDocumentTabViewModel(candidate, CloseTabAsync) { Title = "Opening…" };
        Tabs.Add(tab);
        OnPropertyChanged(nameof(HasTabs));
        SelectedTab = tab;

        try {
            await candidate.OpenDocumentAsync(fullPath, cancellationToken).ConfigureAwait(true);
            if (candidate.HasDocument) {
                tab.Title = candidate.DocumentName;
                if (reusedEmptyDocument) _emptyDocument = _createDocument(OpenDocumentAsync);
                return;
            }

            string? error = candidate.ErrorMessage;
            Tabs.Remove(tab);
            OnPropertyChanged(nameof(HasTabs));
            if (reusedEmptyDocument) {
                tab.Dispose();
                _emptyDocument = _createDocument(OpenDocumentAsync);
            } else {
                tab.Dispose();
            }
            SelectedTab = previousTab is not null && Tabs.Contains(previousTab)
                ? previousTab
                : Tabs.LastOrDefault();
            if (!string.IsNullOrWhiteSpace(error)) ActiveDocument.ErrorMessage = error;
        } finally {
            _openingDocument = false;
        }
    }

    private static bool IsPathIdentityFailure(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException;

    private static bool DocumentOwnsPath(MainWindowViewModel document, string path) =>
        document.DocumentPath is { Length: > 0 } source && OfficePathIdentity.AreEquivalent(source, path);

    internal async Task CloseTabAsync(StudioDocumentTabViewModel tab) {
        if (_disposed || !Tabs.Contains(tab)) return;
        SelectedTab = tab;
        if (!await tab.Document.RequestCloseDocumentAsync().ConfigureAwait(true)) return;

        int index = Tabs.IndexOf(tab);
        Tabs.Remove(tab);
        OnPropertyChanged(nameof(HasTabs));
        if (Tabs.Count == 0) {
            tab.Dispose();
            SelectedTab = null;
            return;
        }

        SelectedTab = Tabs[Math.Min(index, Tabs.Count - 1)];
        tab.Dispose();
    }

    internal async Task<bool> RequestCloseAllAsync() {
        StudioDocumentTabViewModel[] candidates = Tabs.ToArray();
        StudioDocumentTabViewModel? previousSelection = SelectedTab;
        bool prepared = false;
        try {
            foreach (StudioDocumentTabViewModel tab in candidates) {
                SelectedTab = tab;
                if (!await tab.Document.PrepareCloseDocumentAsync().ConfigureAwait(true)) return false;
            }
            foreach (StudioDocumentTabViewModel tab in candidates) {
                SelectedTab = tab;
                if (!await tab.Document.CommitPreparedDiscardAsync().ConfigureAwait(true)) return false;
            }
            prepared = true;
        } finally {
            if (!prepared) {
                foreach (StudioDocumentTabViewModel tab in candidates) tab.Document.CancelPreparedClose();
                if (previousSelection is not null && Tabs.Contains(previousSelection)) SelectedTab = previousSelection;
            }
        }
        if (previousSelection is not null && Tabs.Contains(previousSelection)) SelectedTab = previousSelection;
        CloseAllPrepared?.Invoke(this, EventArgs.Empty);
        foreach (StudioDocumentTabViewModel tab in candidates) {
            tab.Document.CompletePreparedClose();
            Tabs.Remove(tab);
            tab.Dispose();
        }
        OnPropertyChanged(nameof(HasTabs));
        SelectedTab = null;
        return true;
    }

    internal void CancelAllOperations() {
        foreach (StudioDocumentTabViewModel tab in Tabs) tab.Document.CancelCurrentOperation();
        _emptyDocument.CancelCurrentOperation();
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        foreach (StudioDocumentTabViewModel tab in Tabs.ToArray()) tab.Dispose();
        Tabs.Clear();
        _emptyDocument.Dispose();
    }
}
