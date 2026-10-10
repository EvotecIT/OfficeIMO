using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Owns document lifecycles and coordinates tab activation and close semantics across UI hosts.</summary>
internal sealed partial class StudioDocumentTabs<TDocument, TTab> : ObservableObject, IDisposable
    where TDocument : class, IStudioDocument
    where TTab : class, IStudioDocumentTab<TDocument> {
    private readonly Func<Func<string, CancellationToken, Task>, TDocument> _createDocument;
    private readonly Func<TDocument, Func<TTab, Task>, TTab> _createTab;
    private readonly Action<TDocument> _activateDocument;
    private readonly Func<TDocument, Task<bool>>? _prepareActiveClose;
    private TDocument _emptyDocument;
    private bool _openingDocument;
    private bool _disposed;
    private readonly HashSet<TDocument> _presentedDocuments = new();

    internal StudioDocumentTabs(
        Func<Func<string, CancellationToken, Task>, TDocument> createDocument,
        Action<TDocument> activateDocument,
        Func<TDocument, Func<TTab, Task>, TTab> createTab,
        Func<TDocument, Task<bool>>? prepareActiveClose = null) {
        _createDocument = createDocument ?? throw new ArgumentNullException(nameof(createDocument));
        _activateDocument = activateDocument ?? throw new ArgumentNullException(nameof(activateDocument));
        _createTab = createTab ?? throw new ArgumentNullException(nameof(createTab));
        _prepareActiveClose = prepareActiveClose;
        _emptyDocument = _createDocument(OpenDocumentAsync);
    }

    public ObservableCollection<TTab> Tabs { get; } = new();

    internal event EventHandler? CloseAllPrepared;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasTabs))]
    private TTab? _selectedTab;

    public bool HasTabs => Tabs.Count > 0;

    internal TDocument ActiveDocument => SelectedTab?.Document ?? _emptyDocument;

    internal IEnumerable<TDocument> OperationDocuments => Tabs.Select(tab => tab.Document).Append(_emptyDocument).Distinct();

    internal bool HasBusyDocuments => Tabs.Any(tab => tab.Document.CanCancelOperation) ||
                                      _emptyDocument.CanCancelOperation;

    internal bool HasDirtyDocuments => OperationDocuments.Any(document => document.IsDirty || document.HasUnsavedAuxiliaryWork);

    partial void OnSelectedTabChanged(TTab? value) {
        ActiveDocument.SetPresentationActive(true);
        _activateDocument(ActiveDocument);
    }

    partial void OnSelectedTabChanging(TTab? value) {
        ActiveDocument.Deactivate();
        if (!_presentedDocuments.Contains(ActiveDocument)) ActiveDocument.SetPresentationActive(false);
    }

    /// <summary>Keeps the render caches of documents shown in independent panes alive while focus moves between them.</summary>
    internal void SetPresentedDocuments(IEnumerable<TDocument> documents) {
        _presentedDocuments.Clear();
        foreach (TDocument document in documents) _presentedDocuments.Add(document);
        foreach (TDocument document in OperationDocuments)
            document.SetPresentationActive(ReferenceEquals(document, ActiveDocument) || _presentedDocuments.Contains(document));
    }

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
    internal void MoveTab(TTab tab, int index) {
        int current = Tabs.IndexOf(tab);
        if (current < 0 || index < 0 || index >= Tabs.Count || current == index) return;
        var selected = SelectedTab;
        Tabs.Move(current, index);
        SelectedTab = selected;
    }

    internal bool CanPublishPath(string path) => CanDocumentOwnPath(null, path);

    internal bool CanPublishBookPath(TDocument? owner, string path) =>
        CanDocumentOwnPath(null, path, ignoreBookOwner: owner);

    internal bool CanPublishDirectory(string path) {
        if (string.IsNullOrWhiteSpace(path)) return false;
        try {
            return OperationDocuments.SelectMany(document => document.OwnedOutputLocations
                .Concat(document.DocumentPath is { Length: > 0 } source ? [source] : []))
                .All(source => OfficeStorageIdentity.GetLocalPath(source) is null || !OfficePathIdentity.IsSameOrDescendant(source, path));
        } catch (Exception exception) when (IsPathIdentityFailure(exception)) {
            return false;
        }
    }

    internal bool CanDocumentOwnPath(TDocument? document, string path, TDocument? ignoreBookOwner = null) {
        if (string.IsNullOrWhiteSpace(path)) return false;
        try {
            string fullPath = OfficeStorageIdentity.Normalize(path);
            return OperationDocuments.All(candidate =>
                (ReferenceEquals(candidate, document) || !DocumentOwnsPath(candidate, fullPath)) &&
                (ReferenceEquals(candidate, ignoreBookOwner) || !candidate.OwnsOutputPath(fullPath)));
        } catch (Exception exception) when (IsPathIdentityFailure(exception)) {
            // An uninspectable destination cannot safely be authorized for publication.
            return false;
        }
    }

    internal async Task OpenDocumentAsync(string path, CancellationToken cancellationToken = default) {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_openingDocument || string.IsNullOrWhiteSpace(path)) return;

        string fullPath;
        TTab? existing;
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
        TTab? previousTab = SelectedTab;
        bool reusedEmptyDocument = Tabs.Count == 0 && !_emptyDocument.HasDocument;
        TDocument candidate = reusedEmptyDocument
            ? _emptyDocument
            : _createDocument(OpenDocumentAsync);
        var tab = _createTab(candidate, CloseTabAsync);
        tab.Title = "Opening…";
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

    private static bool DocumentOwnsPath(TDocument document, string path) =>
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

    internal async Task CloseTabAsync(TTab tab) {
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
        TTab[] candidates = Tabs.ToArray();
        TTab? previousSelection = SelectedTab;
        bool prepared = false;
        try {
            foreach (TTab tab in candidates) {
                SelectedTab = tab;
                if (!await tab.Document.PrepareCloseDocumentAsync().ConfigureAwait(true)) return false;
            }
            foreach (TTab tab in candidates) {
                SelectedTab = tab;
                if (!await tab.Document.CommitPreparedDiscardAsync().ConfigureAwait(true)) return false;
            }
            prepared = true;
        } finally {
            if (!prepared) {
                foreach (TTab tab in candidates) tab.Document.CancelPreparedClose();
                if (previousSelection is not null && Tabs.Contains(previousSelection)) SelectedTab = previousSelection;
            }
        }
        if (previousSelection is not null && Tabs.Contains(previousSelection)) SelectedTab = previousSelection;
        CloseAllPrepared?.Invoke(this, EventArgs.Empty);
        foreach (TTab tab in candidates) {
            tab.Document.CompletePreparedClose();
            Tabs.Remove(tab);
            tab.Dispose();
        }
        OnPropertyChanged(nameof(HasTabs));
        SelectedTab = null;
        return true;
    }

    internal void CancelAllOperations() {
        foreach (TTab tab in Tabs) tab.Document.CancelCurrentOperation();
        _emptyDocument.CancelCurrentOperation();
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        foreach (TTab tab in Tabs.ToArray()) tab.Dispose();
        Tabs.Clear();
        _emptyDocument.Dispose();
    }
}
