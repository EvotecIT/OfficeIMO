using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private void OnRecoveryMaintenanceCompleted(object? sender, EventArgs args) => Avalonia.Threading.Dispatcher.UIThread.Post(() => {
        if (!_disposed) OnPropertyChanged(nameof(HasRecovery));
    });

    private readonly HashSet<int> _organizerSelection = new();
    private CancellationTokenSource? _operationCancellation;
    private bool _disposeWhenIdle;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanStartDocumentTransition))]
    [NotifyPropertyChangedFor(nameof(CanCancelOperation))]
    [NotifyPropertyChangedFor(nameof(CanReviewComment))]
    [NotifyPropertyChangedFor(nameof(CanReplyToComment))]
    [NotifyPropertyChangedFor(nameof(CanResolveComment))]
    [NotifyPropertyChangedFor(nameof(CanReopenComment))]
    [NotifyCanExecuteChangedFor(nameof(ReplyToCommentCommand))]
    [NotifyCanExecuteChangedFor(nameof(ResolveCommentCommand))]
    [NotifyCanExecuteChangedFor(nameof(ReopenCommentCommand))]
    private bool _isWorkspaceBusy;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOperationStatus))]
    private string? _operationStatus;

    [ObservableProperty]
    private double _operationProgressFraction;

    [ObservableProperty]
    private PdfBookmarkViewModel? _selectedBookmark;

    [ObservableProperty]
    private double _cropMargin = 12D;

    [ObservableProperty]
    private int _splitPagesPerDocument = 1;

    public ObservableCollection<PdfBookmarkViewModel> Bookmarks { get; } = new();

    public bool HasOperationStatus => !string.IsNullOrWhiteSpace(OperationStatus);

    public bool HasOrganizerSelection => _organizerSelection.Count > 0;

    public bool CanMutateSelection => HasOrganizerSelection && CanMutatePages;

    public bool CanExtractSelection => HasOrganizerSelection && CanExtractPages;

    public bool CanDeleteSelection => CanMutateSelection && _workspace is not null && _organizerSelection.Count < _workspace.Pages.Count;

    public string OrganizerSelectionLabel => _organizerSelection.Count == 0
        ? UiText("Workspace.SelectPages")
        : UiFormat("Workspace.SelectedPageCount", _organizerSelection.Count, OrganizerPages.Count);

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
        if (draggedPageNumber < 1 || draggedPageNumber > _workspace.Pages.Count ||
            targetPageNumber < 1 || targetPageNumber > _workspace.Pages.Count) return;
        var plan = PdfPageReorderPlan.Move(_workspace.Pages.Count, targetPageNumber, moved);
        await ApplyPageReorderAsync(plan, moved, CancellationToken.None).ConfigureAwait(true);
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
        if (string.IsNullOrWhiteSpace(path)) return;
        string fullPath = OfficeIMO.Internal.OfficeStorageIdentity.Normalize(path);
        if (!_canSaveAsPath(fullPath)) {
            OperationStatus = UiText("Workspace.SaveAsAlreadyOpen");
            return;
        }
        await RunSaveAsync(fullPath, cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task UndoAsync(CancellationToken cancellationToken) {
        if (_workspace?.CanUndo != true) return;
        bool succeeded = await RunMutationAsync(token => _workspace.UndoAsync(token), cancellationToken).ConfigureAwait(true);
        if (succeeded) OperationStatus = UiText("Workspace.UndoComplete");
    }

    [RelayCommand]
    private async Task RedoAsync(CancellationToken cancellationToken) {
        if (_workspace?.CanRedo != true) return;
        bool succeeded = await RunMutationAsync(token => _workspace.RedoAsync(token), cancellationToken).ConfigureAwait(true);
        if (succeeded) OperationStatus = UiText("Workspace.RedoComplete");
    }

    [RelayCommand]
    private async Task RestoreRecoveryAsync(CancellationToken cancellationToken) {
        if (_workspace?.HasRecovery != true) return;
        await RunMutationAsync(token => _workspace.RestoreRecoveryAsync(token), cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task DiscardRecoveryAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        bool succeeded = await RunStandaloneAsync(token => _workspace.DiscardRecoveryAsync(token), cancellationToken).ConfigureAwait(true);
        if (succeeded) OperationStatus = UiText("Workspace.RecoveryDiscarded");
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
            ErrorMessage = UiText("Workspace.AtLeastOnePage");
            return;
        }
        if (!await _confirmPageDeletion(pages.Length).ConfigureAwait(true)) {
            OperationStatus = UiText("Workspace.DeleteCancelled");
            return;
        }
        await RunMutationAsync(token => _workspace.DeleteAsync(pages, token, CreateProgress()), cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task MoveSelectedUpAsync(CancellationToken cancellationToken) {
        if (_workspace is null || !CanMutateSelection) return;
        int[] selected = GetSelectedPages();
        await ApplyPageReorderAsync(PdfPageReorderPlan.Shift(_workspace.Pages.Count, true, selected), selected, cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task MoveSelectedDownAsync(CancellationToken cancellationToken) {
        if (_workspace is null || !CanMutateSelection) return;
        int[] selected = GetSelectedPages();
        await ApplyPageReorderAsync(PdfPageReorderPlan.Shift(_workspace.Pages.Count, false, selected), selected, cancellationToken).ConfigureAwait(true);
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
    private void SelectAllPages() => SetOrganizerSelection(OrganizerPages);

    private bool IsPageWorkflowCurrent(PdfWorkspace workspace, long revision) {
        if (!_disposed && ReferenceEquals(workspace, _workspace) && workspace.Revision == revision) return true;
        if (!_disposed) ErrorMessage = UiText("Organizer.StalePreview");
        return false;
    }

    private bool IsReviewedCopyCurrent(PdfWorkspace workspace, long revision) {
        if (!IsPageWorkflowCurrent(workspace, revision)) return false;
        if (!HasFormDrafts) return true;
        ErrorMessage = UiText("Workspace.CopyHasFormDrafts");
        return false;
    }

    [RelayCommand]
    private void ClearPageSelection() => SetOrganizerSelection(Array.Empty<PdfOrganizerPageViewModel>());

    [RelayCommand]
    private void CancelOperation() => CancelCurrentOperation();

    private async Task<bool> RunSaveAsync(string? path, CancellationToken cancellationToken) {
        if (_workspace is null) return false;
        PdfWorkspace workspace = _workspace;
        if (path is null && _services.Storage.IsRecoveryLocation(workspace.Path)) {
            path = await _pickSavePdf(cancellationToken).ConfigureAwait(true);
            if (string.IsNullOrWhiteSpace(path)) return false;
            if (!_canSaveAsPath(path)) { OperationStatus = UiText("Workspace.SaveAsAlreadyOpen"); return false; }
        }
        if (path is null && workspace.UsesProviderPublication() && !await _confirmProviderWrite(workspace.Path)) return false;
        if (!ReferenceEquals(workspace, _workspace) || _disposed) return false;
        var formValues = CaptureFormDrafts();
        if (formValues is null) return false;
        bool formValuesApplied = false;
        bool succeeded = await RunStandaloneAsync(
            async token => {
                if (formValues.Count > 0) {
                    await ApplyCapturedFormValuesAsync(formValues,
                        () => workspace.FillFormFieldsAsync(formValues, token, CreateProgress())).ConfigureAwait(true);
                    formValuesApplied = true;
                }
                await workspace.SaveAsync(path, token, CreateProgress()).ConfigureAwait(true);
            },
            cancellationToken).ConfigureAwait(true);
        if (formValuesApplied && ReferenceEquals(workspace, _workspace) && !_disposed) {
            ClearSignatureValidation();
            RefreshWorkspacePresentation();
        }
        if (succeeded) NotifyWorkspaceStateChanged();
        return succeeded;
    }

    private async Task<bool> RunMutationAsync(
        Func<CancellationToken, Task> operation,
        CancellationToken cancellationToken,
        IReadOnlyCollection<int>? organizerSelection = null) {
        bool succeeded = await RunStandaloneAsync(operation, cancellationToken).ConfigureAwait(true);
        if (succeeded && _workspace is not null) {
            ClearSignatureValidation();
            RefreshWorkspacePresentation(organizerSelection);
        }
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
            OperationStatus = UiText("Workspace.OperationCompleted");
            return true;
        } catch (OperationCanceledException) when (currentCancellation.IsCancellationRequested) {
            OperationStatus = UiText("Workspace.OperationCancelled");
            return false;
        } catch (Exception ex) {
            ErrorMessage = ex.Message;
            OperationStatus = UiText("Workspace.OperationFailed");
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

    private IProgress<PdfWorkspaceProgress> CreateProgress() {
        CancellationTokenSource? attempt = _operationCancellation;
        return new Progress<PdfWorkspaceProgress>(progress => {
            if (!IsWorkspaceBusy || !ReferenceEquals(attempt, _operationCancellation)) return;
            OperationStatus = progress.Stage;
            OperationProgressFraction = Math.Clamp(progress.Fraction, 0D, 1D);
        });
    }

    internal void CancelCurrentOperation() {
        _operationCancellation?.Cancel();
        _openCancellation?.Cancel();
        CancelComparisonOpen();
        if (ConversionWorkbench.CanCancel) ConversionWorkbench.CancelCommand.Execute(null);
        if (OutputWorkbench.CanCancel) OutputWorkbench.CancelCommand.Execute(null);
        if (DocumentHealth.CanCancel) DocumentHealth.CancelCommand.Execute(null);
        if (OcrWorkbench.CanCancel) OcrWorkbench.CancelCommand.Execute(null);
        if (OcrSession.IsBusy) OcrSession.CancelCommand.Execute(null);
        if (CanCancelOperation) OperationStatus = UiText("Workspace.CancellingOperation");
    }

    private int[] GetSelectedPages() => _organizerSelection.OrderBy(static page => page).ToArray();

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
            renderCoordinator,
            _localizer)).ToArray();
        PdfOrganizerPageViewModel[] organizerPages = session.Pages.Select(page => new PdfOrganizerPageViewModel(
            page.PageNumber,
            page.Width,
            page.Height,
            page.RotationDegrees,
            sceneCoordinator,
            renderCoordinator,
            _localizer)).ToArray();

        ReplaceDocument(_workspace, session, sceneCoordinator, renderCoordinator, pages, organizerPages, organizerSelection);
        SelectedPage = Pages[selectedPage - 1];
    }

    private void NotifyWorkspaceStateChanged() {
        _assistant?.CheckSource();
        OnPropertyChanged(nameof(CanSearchDocument));
        OnPropertyChanged(nameof(ReaderHint));
        SearchCommand.NotifyCanExecuteChanged();
        NotifyCommentActions();
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
        OnPropertyChanged(nameof(CanAuthorForms));
        OnPropertyChanged(nameof(CanFlattenSelectedFormField));
        OnPropertyChanged(nameof(IsDocumentEncrypted));
        OnPropertyChanged(nameof(HasDocumentSignatures));
        OnPropertyChanged(nameof(CanChangeProtection));
        OnPropertyChanged(nameof(CanRemoveProtection));
        OnPropertyChanged(nameof(CanApplyCertificateSignature));
        OnPropertyChanged(nameof(SecurityWarning));
        OnPropertyChanged(nameof(HasSecurityWarning));
        if (_workspace is not null) {
            DocumentName = _workspace.FileName + (_workspace.IsDirty ? " *" : string.Empty);
            string pageLabel = _workspace.Pages.Count == 1 ? UiText("Document.Page") : UiText("Document.Pages");
            DocumentDescription = UiFormat("Document.Summary", _workspace.Pages.Count, pageLabel, FormatByteSize(_workspace.FileSize));
            if (_workspace.HasRecovery) OperationStatus = UiText("Workspace.RecoveryAvailable");
        }
        RebuildBookmarks();
        RebuildFormFields();
    }

    private void RebuildBookmarks() {
        Bookmarks.Clear();
        if (_workspace is null) return;
        foreach (PdfOutlineItem item in (_workspace.DocumentInfo?.Outlines ?? [])) AddBookmark(item);
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

    private async void OnComparisonPageLinkActivated(string target) =>
        await ActivateComparisonPageLinkAsync(target).ConfigureAwait(true);

    internal Task ActivatePageLinkAsync(string target) =>
        ActivatePageLinkAsync(target, _session?.DocumentInfo?.NamedDestinations ?? [], Pages, NavigateToPage);

    internal Task ActivateComparisonPageLinkAsync(string target) =>
        ActivatePageLinkAsync(
            target,
            _comparisonSession?.DocumentInfo?.NamedDestinations ?? [],
            ComparisonPages,
            NavigateToComparisonPage);

    private async Task ActivatePageLinkAsync(
        string target,
        IReadOnlyList<PdfNamedDestination> namedDestinations,
        IReadOnlyList<PdfPageViewModel> pages,
        Action<int> navigateToPage) {
        if (target.StartsWith("page:", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(target.AsSpan(5), out int pageNumber)) {
            navigateToPage(pageNumber);
            return;
        }

        PdfNamedDestination? destination = namedDestinations
            .FirstOrDefault(item => string.Equals(item.Name, target, StringComparison.Ordinal));
        if (destination?.PageNumber is int destinationPage) {
            navigateToPage(destinationPage);
            return;
        }

        switch (target.ToUpperInvariant()) {
            case "NEXTPAGE":
                navigateToPage(Math.Min(pages.Count, GetSelectedPageNumber(pages) + 1));
                return;
            case "PREVPAGE":
            case "PREVIOUSPAGE":
                navigateToPage(Math.Max(1, GetSelectedPageNumber(pages) - 1));
                return;
            case "FIRSTPAGE":
                navigateToPage(1);
                return;
            case "LASTPAGE":
                navigateToPage(pages.Count);
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

        OperationStatus = UiFormat("Workspace.UnsupportedLinkTarget", target);
    }

    private int GetSelectedPageNumber(IReadOnlyList<PdfPageViewModel> pages) =>
        ReferenceEquals(pages, ComparisonPages)
            ? ComparisonSelectedPage?.PageNumber ?? 1
            : SelectedPage?.PageNumber ?? 1;

    private void NavigateToComparisonPage(int pageNumber) {
        if (pageNumber < 1 || pageNumber > ComparisonPages.Count) return;
        ComparisonSelectedPage = ComparisonPages[pageNumber - 1];
    }
}
