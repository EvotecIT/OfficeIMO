using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Owns live document view-models and coordinates tab activation and close semantics.</summary>
public sealed partial class StudioDocumentTabHost : ObservableObject, IDisposable {
    private readonly Func<Func<string, CancellationToken, Task>, MainWindowViewModel> _createDocument;
    private readonly Action<MainWindowViewModel> _activateDocument;
    private readonly Func<MainWindowViewModel, Task<bool>>? _prepareActiveClose;
    private MainWindowViewModel _emptyDocument;
    private bool _openingDocument;
    private bool _disposed;

    internal StudioDocumentTabHost(
        Func<Func<string, CancellationToken, Task>, MainWindowViewModel> createDocument,
        Action<MainWindowViewModel> activateDocument,
        Func<MainWindowViewModel, Task<bool>>? prepareActiveClose = null) {
        _createDocument = createDocument ?? throw new ArgumentNullException(nameof(createDocument));
        _activateDocument = activateDocument ?? throw new ArgumentNullException(nameof(activateDocument));
        _prepareActiveClose = prepareActiveClose;
        _emptyDocument = _createDocument(OpenDocumentAsync);
    }

    public ObservableCollection<StudioDocumentTabViewModel> Tabs { get; } = new();

    internal event EventHandler? CloseAllPrepared;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasTabs))]
    private StudioDocumentTabViewModel? _selectedTab;

    public bool HasTabs => Tabs.Count > 0;

    internal MainWindowViewModel ActiveDocument => SelectedTab?.Document ?? _emptyDocument;

    internal IEnumerable<MainWindowViewModel> OperationDocuments => Tabs.Select(tab => tab.Document).Append(_emptyDocument).Distinct();

    internal bool HasBusyDocuments => Tabs.Any(tab => tab.Document.CanCancelOperation) ||
                                      _emptyDocument.CanCancelOperation;

    internal bool HasDirtyDocuments => OperationDocuments.Any(document => document.IsDirty || document.BookWorkbench.IsDirty);

    partial void OnSelectedTabChanged(StudioDocumentTabViewModel? value) {
        ActiveDocument.SetPresentationActive(true);
        _activateDocument(ActiveDocument);
    }

    partial void OnSelectedTabChanging(StudioDocumentTabViewModel? value) {
        ActiveDocument.DeactivateAssistant();
        ActiveDocument.SetPresentationActive(false);
    }

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

    /// <summary>Changes strip order without replacing the live document or its selected workspace.</summary>
    internal void MoveTab(StudioDocumentTabViewModel tab, int index) {
        int current = Tabs.IndexOf(tab);
        if (current < 0 || index < 0 || index >= Tabs.Count || current == index) return;
        var selected = SelectedTab;
        Tabs.Move(current, index);
        SelectedTab = selected;
    }

    internal bool CanPublishPath(string path) => CanDocumentOwnPath(null, path);

    internal bool CanPublishBookPath(MainWindowViewModel? owner, string path) =>
        CanDocumentOwnPath(null, path, ignoreBookOwner: owner);

    internal bool CanPublishDirectory(string path) {
        if (string.IsNullOrWhiteSpace(path)) return false;
        try {
            return OperationDocuments.SelectMany(document => document.BookWorkbench.OwnedLocations
                .Concat(document.DocumentPath is { Length: > 0 } source ? [source] : []))
                .All(source => OfficeStorageIdentity.GetLocalPath(source) is null || !OfficePathIdentity.IsSameOrDescendant(source, path));
        } catch (Exception exception) when (IsPathIdentityFailure(exception)) {
            return false;
        }
    }

    internal bool CanDocumentOwnPath(MainWindowViewModel? document, string path, MainWindowViewModel? ignoreBookOwner = null) {
        if (string.IsNullOrWhiteSpace(path)) return false;
        try {
            string fullPath = OfficeStorageIdentity.Normalize(path);
            return OperationDocuments.All(candidate =>
                (ReferenceEquals(candidate, document) || !DocumentOwnsPath(candidate, fullPath)) &&
                (ReferenceEquals(candidate, ignoreBookOwner) || !candidate.BookWorkbench.OwnsPath(fullPath)));
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
            fullPath = OfficeStorageIdentity.Normalize(path);
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
        document.DocumentPath is { Length: > 0 } source && OfficeStorageIdentity.AreEquivalent(source, path);

    // Paths of tabs closed in this session, newest last; kept in memory only.
    private readonly List<string> _closedPaths = new();

    internal bool CanReopenClosedTab => _closedPaths.Count > 0;

    /// <summary>Reopens the most recently closed document that is not already open.</summary>
    internal async Task ReopenClosedTabAsync() {
        while (_closedPaths.Count > 0 && !_disposed) {
            string path = _closedPaths[^1];
            try {
                if (Tabs.Any(tab => DocumentOwnsPath(tab.Document, path))) {
                    RemoveClosedPath(path);
                    continue;
                }
                await OpenDocumentAsync(path).ConfigureAwait(true);
                if (Tabs.Any(tab => DocumentOwnsPath(tab.Document, path))) RemoveClosedPath(path);
            } catch (Exception exception) when (IsPathIdentityFailure(exception)) {
                ActiveDocument.ErrorMessage = exception.Message;
            }
            return;
        }
    }

    private void RemoveClosedPath(string path) {
        if (_closedPaths.Remove(path)) OnPropertyChanged(nameof(CanReopenClosedTab));
    }

    internal async Task CloseTabAsync(StudioDocumentTabViewModel tab) {
        if (_disposed || !Tabs.Contains(tab)) return;
        SelectedTab = tab;
        if (tab.Document.CanCancelOperation && _prepareActiveClose is not null &&
            !await _prepareActiveClose(tab.Document).ConfigureAwait(true)) return;
        if (_disposed || !Tabs.Contains(tab)) return;
        if (!await tab.Document.RequestCloseDocumentAsync().ConfigureAwait(true)) return;

        int index = Tabs.IndexOf(tab);
        string? closedPath = tab.Document.LastClosedDocumentPath;
        if (!string.IsNullOrEmpty(closedPath)) {
            _closedPaths.Remove(closedPath);
            _closedPaths.Add(closedPath);
            if (_closedPaths.Count > 10) _closedPaths.RemoveAt(0);
            OnPropertyChanged(nameof(CanReopenClosedTab));
        }
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
        if (!await _emptyDocument.PrepareCloseDocumentAsync().ConfigureAwait(true)) return false;
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
