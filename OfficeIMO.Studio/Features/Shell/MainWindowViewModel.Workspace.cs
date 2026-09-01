using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private readonly HashSet<int> _organizerSelection = new();
    private CancellationTokenSource? _operationCancellation;
    private bool _disposeWhenIdle;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanStartDocumentTransition))]
    [NotifyPropertyChangedFor(nameof(CanCancelOperation))]
    private bool _isWorkspaceBusy;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOperationStatus))]
    private string? _operationStatus;

    [ObservableProperty]
    private double _operationProgressFraction;

    [ObservableProperty]
    private string _searchQuery = string.Empty;

    [ObservableProperty]
    private PdfSearchHit? _selectedSearchResult;

    [ObservableProperty]
    private PdfBookmarkViewModel? _selectedBookmark;

    [ObservableProperty]
    private double _cropMargin = 12D;

    [ObservableProperty]
    private int _splitPagesPerDocument = 1;

    public ObservableCollection<PdfSearchHit> SearchResults { get; } = new();

    public ObservableCollection<PdfBookmarkViewModel> Bookmarks { get; } = new();

    public bool HasOperationStatus => !string.IsNullOrWhiteSpace(OperationStatus);

    public bool HasOrganizerSelection => _organizerSelection.Count > 0;

    public bool CanMutateSelection => HasOrganizerSelection && CanMutatePages;

    public bool CanExtractSelection => HasOrganizerSelection && CanExtractPages;

    public bool CanDeleteSelection => CanMutateSelection && _workspace is not null && _organizerSelection.Count < _workspace.Pages.Count;

    public string OrganizerSelectionLabel => _organizerSelection.Count == 0
        ? "Select pages"
        : $"{_organizerSelection.Count} of {OrganizerPages.Count} selected";

    partial void OnSelectedSearchResultChanged(PdfSearchHit? value) {
        if (value is not null) NavigateToPage(value.PageNumber);
    }

    partial void OnSelectedBookmarkChanged(PdfBookmarkViewModel? value) {
        if (value?.PageNumber is int pageNumber) NavigateToPage(pageNumber);
    }

    internal void SetOrganizerSelection(IEnumerable<PdfOrganizerPageViewModel> pages) {
        _organizerSelection.Clear();
        foreach (PdfOrganizerPageViewModel page in pages) _organizerSelection.Add(page.PageNumber);
        foreach (PdfOrganizerPageViewModel page in OrganizerPages) {
            page.IsSelected = _organizerSelection.Contains(page.PageNumber);
        }
        NotifyOrganizerSelectionChanged();
    }

    internal void UpdateOrganizerSelection(
        IEnumerable<PdfOrganizerPageViewModel> addedPages,
        IEnumerable<PdfOrganizerPageViewModel> removedPages) {
        foreach (PdfOrganizerPageViewModel page in removedPages) {
            _organizerSelection.Remove(page.PageNumber);
            page.IsSelected = false;
        }
        foreach (PdfOrganizerPageViewModel page in addedPages) {
            _organizerSelection.Add(page.PageNumber);
            page.IsSelected = true;
        }
        NotifyOrganizerSelectionChanged();
    }

    private void NotifyOrganizerSelectionChanged() {
        OnPropertyChanged(nameof(HasOrganizerSelection));
        OnPropertyChanged(nameof(CanMutateSelection));
        OnPropertyChanged(nameof(CanExtractSelection));
        OnPropertyChanged(nameof(CanDeleteSelection));
        OnPropertyChanged(nameof(OrganizerSelectionLabel));
    }

    internal void NavigateToOrganizerPage(int pageNumber) => NavigateToPage(pageNumber);

    internal async Task ReorderByDropAsync(int draggedPageNumber, int targetPageNumber) {
        if (_workspace is null || !CanMutatePages || draggedPageNumber == targetPageNumber) return;
        int[] moved = _organizerSelection.Contains(draggedPageNumber)
            ? GetSelectedPages()
            : [draggedPageNumber];
        if (moved.Contains(targetPageNumber)) return;
        var remaining = Enumerable.Range(1, _workspace.Pages.Count)
            .Where(page => !moved.Contains(page))
            .ToList();
        int targetIndex = remaining.IndexOf(targetPageNumber);
        if (targetIndex < 0) targetIndex = remaining.Count;
        remaining.InsertRange(targetIndex, moved);
        int[] selectionAfter = MapSelectedPagesToReorderedPositions(remaining, moved);
        await RunMutationAsync(
            token => _workspace.ReorderAsync(remaining, token, CreateProgress()),
            CancellationToken.None,
            selectionAfter).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task SaveAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        await RunSaveAsync(path: null, cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task SaveAsAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        string? path = await _pickSavePdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) await RunSaveAsync(path, cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task UndoAsync(CancellationToken cancellationToken) {
        if (_workspace?.CanUndo != true) return;
        bool succeeded = await RunMutationAsync(token => _workspace.UndoAsync(token), cancellationToken).ConfigureAwait(true);
        if (succeeded) OperationStatus = "Undo complete.";
    }

    [RelayCommand]
    private async Task RedoAsync(CancellationToken cancellationToken) {
        if (_workspace?.CanRedo != true) return;
        bool succeeded = await RunMutationAsync(token => _workspace.RedoAsync(token), cancellationToken).ConfigureAwait(true);
        if (succeeded) OperationStatus = "Redo complete.";
    }

    [RelayCommand]
    private async Task RestoreRecoveryAsync(CancellationToken cancellationToken) {
        if (_workspace?.HasRecovery != true) return;
        await RunMutationAsync(token => _workspace.RestoreRecoveryAsync(token), cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private void DiscardRecovery() {
        _workspace?.DiscardRecovery();
        OperationStatus = "Recovery snapshot discarded";
        NotifyWorkspaceStateChanged();
    }

    [RelayCommand]
    private async Task RotateLeftAsync(CancellationToken cancellationToken) {
        int[] pages = GetSelectedPages();
        if (_workspace is null || pages.Length == 0) return;
        await RunMutationAsync(token => _workspace.RotateAsync(pages, -90, token, CreateProgress()), cancellationToken, pages).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task RotateRightAsync(CancellationToken cancellationToken) {
        int[] pages = GetSelectedPages();
        if (_workspace is null || pages.Length == 0) return;
        await RunMutationAsync(token => _workspace.RotateAsync(pages, 90, token, CreateProgress()), cancellationToken, pages).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task DuplicateSelectedAsync(CancellationToken cancellationToken) {
        int[] pages = GetSelectedPages();
        if (_workspace is null || pages.Length == 0) return;
        int[] duplicatePositions = pages.Select((pageNumber, index) => pageNumber + index + 1).ToArray();
        await RunMutationAsync(
            token => _workspace.DuplicateAsync(pages, token, CreateProgress()),
            cancellationToken,
            duplicatePositions).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task DeleteSelectedAsync(CancellationToken cancellationToken) {
        int[] pages = GetSelectedPages();
        if (_workspace is null || pages.Length == 0) return;
        if (pages.Length >= _workspace.Pages.Count) {
            ErrorMessage = "A PDF must keep at least one page.";
            return;
        }
        if (!await _confirmPageDeletion(pages.Length).ConfigureAwait(true)) {
            OperationStatus = "Delete cancelled";
            return;
        }
        await RunMutationAsync(token => _workspace.DeleteAsync(pages, token, CreateProgress()), cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task MoveSelectedUpAsync(CancellationToken cancellationToken) {
        int[] order = BuildMovedOrder(moveUp: true);
        if (_workspace is null || order.Length == 0) return;
        int[] selectionAfter = MapSelectedPagesToReorderedPositions(order, GetSelectedPages());
        await RunMutationAsync(token => _workspace.ReorderAsync(order, token, CreateProgress()), cancellationToken, selectionAfter).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task MoveSelectedDownAsync(CancellationToken cancellationToken) {
        int[] order = BuildMovedOrder(moveUp: false);
        if (_workspace is null || order.Length == 0) return;
        int[] selectionAfter = MapSelectedPagesToReorderedPositions(order, GetSelectedPages());
        await RunMutationAsync(token => _workspace.ReorderAsync(order, token, CreateProgress()), cancellationToken, selectionAfter).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task CropSelectedAsync(CancellationToken cancellationToken) {
        int[] pages = GetSelectedPages();
        if (_workspace is null || pages.Length == 0) return;
        await RunMutationAsync(token => _workspace.CropByMarginAsync(pages, CropMargin, token, CreateProgress()), cancellationToken, pages).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task InsertBlankAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        int insertBefore = _organizerSelection.Count == 0
            ? _workspace.Pages.Count + 1
            : _organizerSelection.Min();
        PdfPageInfo reference = SelectedPage is null
            ? _workspace.Pages[0]
            : _workspace.Pages[SelectedPage.PageNumber - 1];
        await RunMutationAsync(
            token => _workspace.InsertBlankAsync(insertBefore, reference.Width, reference.Height, token, CreateProgress()),
            cancellationToken,
            new[] { insertBefore }).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task ImportPagesAsync(CancellationToken cancellationToken) {
        if (_workspace is null || !CanImportPages) return;
        IReadOnlyList<string> paths = await _pickImportPdfs(cancellationToken).ConfigureAwait(true);
        if (paths.Count == 0) return;
        int insertBefore = _organizerSelection.Count == 0
            ? _workspace.Pages.Count + 1
            : _organizerSelection.Min();
        int importedPageCount = 0;
        bool succeeded = await RunStandaloneAsync(
            async token => importedPageCount = await _workspace
                .ImportAsync(paths, insertBefore, token, CreateProgress())
                .ConfigureAwait(true),
            cancellationToken).ConfigureAwait(true);
        if (succeeded) {
            RefreshWorkspacePresentation(Enumerable.Range(insertBefore, importedPageCount).ToArray());
            OperationStatus = importedPageCount == 1
                ? "Imported 1 page"
                : $"Imported {importedPageCount} pages from {paths.Count} PDFs";
        }
    }

    [RelayCommand]
    private async Task ExtractSelectedAsync(CancellationToken cancellationToken) {
        int[] pages = GetSelectedPages();
        if (_workspace is null || !CanExtractPages || pages.Length == 0) return;
        string? path = await _pickSavePdf(cancellationToken).ConfigureAwait(true);
        if (string.IsNullOrWhiteSpace(path)) return;
        await RunStandaloneAsync(
            token => _workspace.ExtractAsync(pages, path, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task SplitAsync(CancellationToken cancellationToken) {
        if (_workspace is null || !CanExtractPages) return;
        string? folder = await _pickOutputFolder(cancellationToken).ConfigureAwait(true);
        if (string.IsNullOrWhiteSpace(folder)) return;
        IReadOnlyList<string> outputs = Array.Empty<string>();
        bool succeeded = await RunStandaloneAsync(
            async token => outputs = await _workspace
                .SplitAsync(folder, SplitPagesPerDocument, token, CreateProgress())
                .ConfigureAwait(true),
            cancellationToken).ConfigureAwait(true);
        if (succeeded) {
            OperationStatus = outputs.Count == 1
                ? "Created 1 split PDF"
                : $"Created {outputs.Count} split PDFs";
        }
    }

    [RelayCommand]
    private void SelectAllPages() => SetOrganizerSelection(OrganizerPages);

    [RelayCommand]
    private void ClearPageSelection() => SetOrganizerSelection(Array.Empty<PdfOrganizerPageViewModel>());

    [RelayCommand]
    private async Task SearchAsync(CancellationToken cancellationToken) {
        if (_session is null || string.IsNullOrWhiteSpace(SearchQuery)) {
            SearchResults.Clear();
            return;
        }

        OperationStatus = "Searching document";
        await RunStandaloneAsync(async token => {
            var progress = new Progress<double>(fraction =>
                OperationProgressFraction = Math.Clamp(fraction, 0D, 1D));
            IReadOnlyList<PdfSearchHit> results = await _session
                .SearchAsync(SearchQuery, token, progress)
                .ConfigureAwait(true);
            SearchResults.Clear();
            foreach (PdfSearchHit result in results) SearchResults.Add(result);
            OperationProgressFraction = 1D;
            OperationStatus = results.Count == 0 ? "No matches" : $"{results.Count} matching page(s)";
        }, cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private void CancelOperation() => CancelCurrentOperation();

    private async Task RunSaveAsync(string? path, CancellationToken cancellationToken) {
        if (_workspace is null) return;
        bool succeeded = await RunStandaloneAsync(
            token => _workspace.SaveAsync(path, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        if (succeeded) NotifyWorkspaceStateChanged();
    }

    private async Task<bool> RunMutationAsync(
        Func<CancellationToken, Task> operation,
        CancellationToken cancellationToken,
        IReadOnlyCollection<int>? organizerSelection = null) {
        bool succeeded = await RunStandaloneAsync(operation, cancellationToken).ConfigureAwait(true);
        if (succeeded && _workspace is not null) RefreshWorkspacePresentation(organizerSelection);
        return succeeded;
    }

    private async Task<bool> RunStandaloneAsync(Func<CancellationToken, Task> operation, CancellationToken cancellationToken) {
        if (IsWorkspaceBusy) return false;
        var currentCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        _operationCancellation = currentCancellation;
        IsWorkspaceBusy = true;
        OperationProgressFraction = 0D;
        ErrorMessage = null;
        try {
            await operation(currentCancellation.Token).ConfigureAwait(true);
            OperationProgressFraction = 1D;
            return true;
        } catch (OperationCanceledException) when (currentCancellation.IsCancellationRequested) {
            OperationStatus = "Operation cancelled";
            return false;
        } catch (Exception ex) {
            ErrorMessage = ex.Message;
            OperationStatus = "Operation failed";
            return false;
        } finally {
            if (ReferenceEquals(_operationCancellation, currentCancellation)) _operationCancellation = null;
            currentCancellation.Dispose();
            IsWorkspaceBusy = false;
            NotifyWorkspaceStateChanged();
            if (_disposeWhenIdle) {
                _disposeWhenIdle = false;
                ReplaceDocument(null, null, null, null, Array.Empty<PdfPageViewModel>(), Array.Empty<PdfOrganizerPageViewModel>());
            }
        }
    }

    private IProgress<PdfWorkspaceProgress> CreateProgress() =>
        new Progress<PdfWorkspaceProgress>(progress => {
            OperationStatus = progress.Stage;
            OperationProgressFraction = Math.Clamp(progress.Fraction, 0D, 1D);
        });

    internal void CancelCurrentOperation() {
        _operationCancellation?.Cancel();
        _openCancellation?.Cancel();
        if (ConversionWorkbench.CanCancel) ConversionWorkbench.CancelCommand.Execute(null);
        if (DocumentHealth.CanCancel) DocumentHealth.CancelCommand.Execute(null);
        if (CanCancelOperation) OperationStatus = "Cancelling operation";
    }

    private int[] GetSelectedPages() => _organizerSelection.OrderBy(static page => page).ToArray();

    private int[] BuildMovedOrder(bool moveUp) {
        if (_workspace is null || _organizerSelection.Count == 0) return Array.Empty<int>();
        int[] order = Enumerable.Range(1, _workspace.Pages.Count).ToArray();
        if (moveUp) {
            for (int index = 1; index < order.Length; index++) {
                if (_organizerSelection.Contains(order[index]) && !_organizerSelection.Contains(order[index - 1])) {
                    (order[index - 1], order[index]) = (order[index], order[index - 1]);
                }
            }
        } else {
            for (int index = order.Length - 2; index >= 0; index--) {
                if (_organizerSelection.Contains(order[index]) && !_organizerSelection.Contains(order[index + 1])) {
                    (order[index], order[index + 1]) = (order[index + 1], order[index]);
                }
            }
        }
        return order;
    }

    private static int[] MapSelectedPagesToReorderedPositions(IReadOnlyList<int> order, IReadOnlyCollection<int> selectedPages) {
        var selected = new HashSet<int>(selectedPages);
        return order
            .Select((originalPageNumber, index) => new { originalPageNumber, position = index + 1 })
            .Where(item => selected.Contains(item.originalPageNumber))
            .Select(static item => item.position)
            .ToArray();
    }

    private void RefreshWorkspacePresentation(IReadOnlyCollection<int>? organizerSelection = null) {
        if (_workspace is null) return;
        CancelPendingRedaction();
        ClearObjectSelection();
        int selectedPage = Math.Clamp(SelectedPage?.PageNumber ?? 1, 1, _workspace.Pages.Count);
        PdfDocumentSession session = PdfDocumentSession.FromWorkspace(_workspace);
        var sceneCoordinator = new PageSceneCoordinator(session.LoadPageSceneAsync);
        var renderCoordinator = new PageRenderCoordinator(session.RenderPageAsync);
        PdfPageViewModel[] pages = session.Pages.Select(page => new PdfPageViewModel(
            page.PageNumber,
            page.Width,
            page.Height,
            page.RotationDegrees,
            Zoom,
            sceneCoordinator,
            renderCoordinator)).ToArray();
        PdfOrganizerPageViewModel[] organizerPages = session.Pages.Select(page => new PdfOrganizerPageViewModel(
            page.PageNumber,
            page.Width,
            page.Height,
            page.RotationDegrees,
            sceneCoordinator,
            renderCoordinator)).ToArray();

        ReplaceDocument(_workspace, session, sceneCoordinator, renderCoordinator, pages, organizerPages, organizerSelection);
        SelectedPage = Pages[selectedPage - 1];
    }

    private void NotifyWorkspaceStateChanged() {
        OnPropertyChanged(nameof(IsDirty));
        OnPropertyChanged(nameof(CanUndo));
        OnPropertyChanged(nameof(CanRedo));
        OnPropertyChanged(nameof(HasRecovery));
        OnPropertyChanged(nameof(CanMutatePages));
        OnPropertyChanged(nameof(CanExtractPages));
        OnPropertyChanged(nameof(CanImportPages));
        OnPropertyChanged(nameof(CanMutateSelection));
        OnPropertyChanged(nameof(CanExtractSelection));
        OnPropertyChanged(nameof(CanDeleteSelection));
        OnPropertyChanged(nameof(CanEditAnnotations));
        OnPropertyChanged(nameof(CanEditPageContent));
        OnPropertyChanged(nameof(CanReplaceSelectedText));
        OnPropertyChanged(nameof(CanReplaceSelectedImage));
        OnPropertyChanged(nameof(CanResizeSelectedAnnotation));
        OnPropertyChanged(nameof(CanRedact));
        OnPropertyChanged(nameof(CanFillForms));
        OnPropertyChanged(nameof(CanFlattenForms));
        OnPropertyChanged(nameof(CanFillAndFlattenForms));
        OnPropertyChanged(nameof(SecurityWarning));
        OnPropertyChanged(nameof(HasSecurityWarning));
        if (_workspace is not null) {
            DocumentName = _workspace.FileName + (_workspace.IsDirty ? " *" : string.Empty);
            DocumentDescription = $"{_workspace.Pages.Count:N0} {(_workspace.Pages.Count == 1 ? "page" : "pages")} · {FormatByteSize(_workspace.FileSize)}";
            if (_workspace.HasRecovery) OperationStatus = "Recovered edits are available for this document.";
        }
        RebuildBookmarks();
        RebuildFormFields();
    }

    private void RebuildBookmarks() {
        Bookmarks.Clear();
        if (_workspace is null) return;
        foreach (PdfOutlineItem item in _workspace.DocumentInfo.Outlines) AddBookmark(item);
    }

    private void AddBookmark(PdfOutlineItem item) {
        Bookmarks.Add(new PdfBookmarkViewModel(item.Title, item.Level, item.PageNumber));
        foreach (PdfOutlineItem child in item.Children) AddBookmark(child);
    }

    private void NavigateToPage(int pageNumber) {
        if (pageNumber < 1 || pageNumber > Pages.Count) return;
        SelectedPage = Pages[pageNumber - 1];
    }

    private async void OnPageLinkActivated(string target) => await ActivatePageLinkAsync(target).ConfigureAwait(true);

    internal async Task ActivatePageLinkAsync(string target) {
        if (target.StartsWith("page:", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(target.AsSpan(5), out int pageNumber)) {
            NavigateToPage(pageNumber);
            return;
        }

        PdfNamedDestination? destination = _workspace?.DocumentInfo.NamedDestinations
            .FirstOrDefault(item => string.Equals(item.Name, target, StringComparison.Ordinal));
        if (destination?.PageNumber is int destinationPage) {
            NavigateToPage(destinationPage);
            return;
        }

        switch (target.ToUpperInvariant()) {
            case "NEXTPAGE":
                NextPageCommand.Execute(null);
                return;
            case "PREVPAGE":
            case "PREVIOUSPAGE":
                PreviousPageCommand.Execute(null);
                return;
            case "FIRSTPAGE":
                NavigateToPage(1);
                return;
            case "LASTPAGE":
                NavigateToPage(Pages.Count);
                return;
        }

        if (Uri.TryCreate(target, UriKind.Absolute, out Uri? uri) &&
            (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeMailto)) {
            try {
                await _openUri(uri).ConfigureAwait(true);
            } catch (Exception ex) {
                ErrorMessage = ex.Message;
            }
            return;
        }

        OperationStatus = $"This link target is not supported: {target}";
    }
}
