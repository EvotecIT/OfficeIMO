using System.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>
/// Pages report rendering notes as they render. Pages that can look noticeably different from the original (skipped
/// content, replaced embedded fonts) raise one notice for the document; routine notes such as an installed font
/// standing in for an unembedded one stay in the note under each page.
/// </summary>
public sealed partial class MainWindowViewModel {
    private const int RenderingNoticeListedPages = 5;
    private readonly SortedSet<int> _renderingNoticePages = new();
    private bool _renderingNoticeDismissed;
    private long _renderingNoticeRevision = -1;

    public bool HasRenderingNotice =>
        _renderingNoticePages.Count > 0 && !_renderingNoticeDismissed && WorkspaceMode == StudioWorkspaceMode.PdfWorkspace;

    public string? RenderingNotice {
        get {
            if (_renderingNoticePages.Count == 0) return null;
            if (_renderingNoticePages.Count == 1) return _localizer.Format("Notices.RenderingLimitation", _renderingNoticePages.Min);
            string pages = string.Join(", ", _renderingNoticePages.Take(RenderingNoticeListedPages));
            if (_renderingNoticePages.Count > RenderingNoticeListedPages) pages += ", …";
            return _localizer.Format("Notices.RenderingLimitationPages", pages);
        }
    }

    public string? RenderingNoticeAction =>
        _renderingNoticePages.Count == 0 ? null : _localizer.Format("Notices.ShowPage", _renderingNoticePages.Min);

    [RelayCommand]
    private void ShowRenderingNoticePage() {
        if (_renderingNoticePages.Count == 0) return;
        if (WorkspaceMode != StudioWorkspaceMode.PdfWorkspace) WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
        NavigateToPage(_renderingNoticePages.Min);
    }

    // Dismissing hides the notice for this document; opening another document offers it again.
    [RelayCommand]
    private void DismissRenderingNotice() {
        _renderingNoticeDismissed = true;
        OnPropertyChanged(nameof(HasRenderingNotice));
    }

    private void TrackRenderingNotices(PdfPageViewModel page) {
        page.PropertyChanged += OnPageRenderingNoticeChanged;
        UpdateRenderingNotice(page);
    }

    private void ResetRenderingNotices(bool newDocument, PdfWorkspace? workspace) {
        foreach (PdfPageViewModel page in Pages) page.PropertyChanged -= OnPageRenderingNoticeChanged;
        long revision = workspace?.Revision ?? -1;
        bool pagesChanged = revision != _renderingNoticeRevision && workspace?.Journal.LastOrDefault()?.Kind is
            PdfWorkspaceOperationKind.Reorder or PdfWorkspaceOperationKind.Delete or PdfWorkspaceOperationKind.Duplicate or
            PdfWorkspaceOperationKind.Import or PdfWorkspaceOperationKind.InsertBlank or PdfWorkspaceOperationKind.Undo or
            PdfWorkspaceOperationKind.Redo or PdfWorkspaceOperationKind.RecoveryRestore;
        // Page numbers no longer identify the same content after organization or history navigation. Reassess those
        // pages as they render; retain known notes across ordinary presentation rebuilds and page-local edits.
        if (newDocument || pagesChanged) {
            _renderingNoticePages.Clear();
            if (newDocument) _renderingNoticeDismissed = false;
        }
        _renderingNoticeRevision = revision;
        NotifyRenderingNotice();
    }

    private void OnPageRenderingNoticeChanged(object? sender, PropertyChangedEventArgs e) {
        if (sender is PdfPageViewModel page && e.PropertyName == nameof(PdfPageViewModel.DiagnosticsSummary)) UpdateRenderingNotice(page);
    }

    // A page that re-renders clears its notes first; what was already learned is kept rather than flickering.
    private void UpdateRenderingNotice(PdfPageViewModel page) {
        if (PdfPageRenderingNotes.AffectsAppearance(page.RenderingNote) && _renderingNoticePages.Add(page.PageNumber)) NotifyRenderingNotice();
    }

    private void NotifyRenderingNotice() {
        OnPropertyChanged(nameof(HasRenderingNotice));
        OnPropertyChanged(nameof(RenderingNotice));
        OnPropertyChanged(nameof(RenderingNoticeAction));
    }
}
