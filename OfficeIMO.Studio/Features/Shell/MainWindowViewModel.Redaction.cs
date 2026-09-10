using System.Collections.ObjectModel;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    public ObservableCollection<PdfRedactionMarkViewModel> RedactionMarks { get; } = new();

    [ObservableProperty] private string _redactionSearchText = string.Empty;
    [ObservableProperty] private bool _redactionSearchExpanded = true;
    [ObservableProperty] private bool _redactionSearchRegex;
    [ObservableProperty] private bool _redactionSearchMatchCase;
    [ObservableProperty] private bool _redactionSearchSelectedPagesOnly;
    [ObservableProperty] private PdfRedactionMarkViewModel? _selectedRedactionMark;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasRedactionEvidence))]
    private PdfRedactionShareableSummary? _lastRedactionSummary;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanExportRedactionEvidence))]
    private string? _lastRedactionCopyPath;
    [ObservableProperty] private bool _sanitizeAfterRedaction;

    partial void OnSanitizeAfterRedactionChanged(bool value) => InvalidateReviewedRedactions();

    public bool CanReviewRedactions => RedactionMarks.Any(mark => mark.IsIncluded);
    public bool CanApplyReviewedRedactions => _pendingRedactionPlan is { IsReviewable: true } && CanReviewRedactions;
    public bool HasRedactionEvidence => LastRedactionSummary is not null;
    public bool CanExportRedactionEvidence => LastRedactionCopyPath is not null && HasRedactionEvidence;

    [RelayCommand]
    private async Task SaveVerifiedRedactionCopyAsync(CancellationToken cancellationToken) {
        if (_workspace is not { } workspace || LastRedactionSummary is not { } summary) return;
        string? destination = null;
        bool succeeded = await RunStandaloneAsync(async token => {
            destination = await _pickSavePdf(token).ConfigureAwait(true);
            if (destination is null) return;
            if (!ReferenceEquals(_workspace, workspace) || !ReferenceEquals(LastRedactionSummary, summary)) return;
            await workspace.SaveVerifiedRedactionCopyAsync(destination, summary, token).ConfigureAwait(true);
        }, cancellationToken).ConfigureAwait(true);
        if (succeeded && destination is not null && ReferenceEquals(_workspace, workspace) && ReferenceEquals(LastRedactionSummary, summary)) {
            LastRedactionCopyPath = destination;
            OperationStatus = _localizer.GetOrDefault("Redaction.CopySaved", "Saved the verified PDF copy. Its evidence report is ready to export.");
        }
    }

    [RelayCommand]
    private async Task ExportRedactionEvidenceAsync(CancellationToken cancellationToken) {
        if (_workspace is not { } workspace || LastRedactionSummary is not { } summary || LastRedactionCopyPath is not { } copy) return;
        await RunStandaloneAsync(async token => {
            string? destination = await _pickSaveRedactionReport(token).ConfigureAwait(true);
            if (destination is null || !ReferenceEquals(_workspace, workspace) || !ReferenceEquals(LastRedactionSummary, summary)) return;
            await workspace.ExportRedactionEvidenceAsync(destination, copy, summary, token).ConfigureAwait(true);
        }, cancellationToken).ConfigureAwait(true);
    }

    partial void OnSelectedRedactionMarkChanged(PdfRedactionMarkViewModel? value) {
        if (value is not null) SelectedPage = Pages.FirstOrDefault(page => page.PageNumber == value.PageNumber);
    }

    [RelayCommand]
    private async Task SearchRedactionsAsync(CancellationToken cancellationToken) {
        if (_workspace is null || !CanRedact || string.IsNullOrWhiteSpace(RedactionSearchText)) return;
        PdfWorkspace workspace = _workspace;
        long revision = workspace.Revision;
        long generation = _redactionPlanGeneration;
        string text = RedactionSearchText;
        bool regex = RedactionSearchRegex;
        bool matchCase = RedactionSearchMatchCase;
        int[]? pages = RedactionSearchSelectedPagesOnly
            ? OrganizerPages.Where(page => page.IsSelected).Select(page => page.PageNumber).ToArray() : null;
        if (pages is { Length: 0 }) {
            ErrorMessage = _localizer.GetOrDefault("Redaction.SelectPages", "Select at least one page in the page organizer.");
            return;
        }
        IReadOnlyList<PdfRedactionMarkViewModel>? marks = null;
        bool succeeded = await RunStandaloneAsync(async token => {
            marks = await workspace.SearchRedactionMarksAsync(text, regex, matchCase, pages, token).ConfigureAwait(true);
        }, cancellationToken).ConfigureAwait(true);
        if (!succeeded || marks is null || generation != _redactionPlanGeneration ||
            !ReferenceEquals(_workspace, workspace) || revision != workspace.Revision) return;
        if (RedactionMarks.Count + marks.Count > 2000) {
            ErrorMessage = _localizer.GetOrDefault("Redaction.TooManyMarks", "A review can contain at most 2,000 marks. Narrow the search or remove some marks.");
            return;
        }
        _pendingRedactionWorkspace = workspace;
        _pendingRedactionRevision = revision;
        foreach (PdfRedactionMarkViewModel mark in marks) AddRedactionMark(mark, update: false);
        InvalidateReviewedRedactions();
        if (marks.Count > 0) RedactionSearchExpanded = false;
        OperationStatus = _localizer.FormatOrDefault("Redaction.SearchResult", "Found {0:N0} matching line(s). Review the marked areas before applying.", marks.Count);
    }

    private void AddRedactionMark(PdfRedactionMarkViewModel mark, bool update = true) {
        if (RedactionMarks.Count >= 2000) {
            ErrorMessage = _localizer.GetOrDefault("Redaction.TooManyMarks", "A review can contain at most 2,000 marks. Narrow the search or remove some marks.");
            return;
        }
        if (RedactionMarks.Any(existing => existing.PageNumber == mark.PageNumber && existing.Bounds == mark.Bounds)) return;
        mark.PropertyChanged += OnRedactionMarkChanged;
        RedactionMarks.Add(mark);
        SelectedRedactionMark ??= mark;
        if (update) InvalidateReviewedRedactions();
    }

    private void OnRedactionMarkChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName is nameof(PdfRedactionMarkViewModel.IsIncluded) or nameof(PdfRedactionMarkViewModel.Reason)) InvalidateReviewedRedactions();
    }

    private void InvalidateReviewedRedactions() {
        _redactionPlanGeneration++;
        _pendingRedactionPlan = null;
        PendingRedactionSummary = RedactionMarks.Count == 0 ? null : _localizer.FormatOrDefault(
            "Redaction.PendingSummary", "{0:N0} of {1:N0} marks selected. Review their impact before applying.",
            RedactionMarks.Count(mark => mark.IsIncluded), RedactionMarks.Count);
        UpdateRedactionOverlays();
        OnPropertyChanged(nameof(CanReviewRedactions));
        OnPropertyChanged(nameof(CanApplyReviewedRedactions));
    }

    [RelayCommand]
    private void RemoveRedactionMark() {
        if (SelectedRedactionMark is not { } mark) return;
        mark.PropertyChanged -= OnRedactionMarkChanged;
        RedactionMarks.Remove(mark);
        SelectedRedactionMark = RedactionMarks.FirstOrDefault();
        InvalidateReviewedRedactions();
    }

    [RelayCommand]
    private async Task ReviewRedactionsAsync(CancellationToken cancellationToken) {
        if (_workspace is not { } workspace || !CanReviewRedactions) return;
        long generation = _redactionPlanGeneration;
        long revision = workspace.Revision;
        if (!ReferenceEquals(workspace, _pendingRedactionWorkspace) || revision != _pendingRedactionRevision) {
            CancelPendingRedaction();
            ErrorMessage = UiText("Editor.RedactionReviewStale");
            return;
        }
        PdfRedactionArea[] areas = RedactionMarks.Where(mark => mark.IsIncluded)
            .Select(mark => new PdfRedactionArea(mark.Area.PageNumber, mark.Area.X, mark.Area.Y,
                mark.Area.Width, mark.Area.Height, mark.Reason.Trim(), mark.Area.ContentScope, mark.Area.AppearanceMode)).ToArray();
        PdfRedactionPlan? plan = null;
        bool succeeded = await RunStandaloneAsync(async token => {
            plan = await workspace.PlanRedactionsAsync(areas, token).ConfigureAwait(true);
        }, cancellationToken).ConfigureAwait(true);
        if (!succeeded || plan is null || generation != _redactionPlanGeneration ||
            !ReferenceEquals(workspace, _workspace) || workspace.Revision != revision) return;
        _pendingRedactionPlan = plan;
        PendingRedactionSummary = _localizer.FormatOrDefault("Redaction.ReviewedSummary",
            "Reviewed {0:N0} areas across {1:N0} pages: {2:N0} text, {3:N0} image, {4:N0} annotation and {5:N0} vector matches.",
            plan.Areas.Count, plan.Areas.Select(area => area.PageNumber).Distinct().Count(),
            plan.Matches.Count(match => match.Kind == PdfRedactionMatchKind.TextBlock),
            plan.Matches.Count(match => match.Kind == PdfRedactionMatchKind.ImagePlacement),
            plan.Matches.Count(match => match.Kind == PdfRedactionMatchKind.Annotation),
            plan.Matches.Count(match => match.Kind == PdfRedactionMatchKind.VectorPath));
        if (plan.Findings.Count > 0) PendingRedactionSummary += " " + string.Join(" ", plan.Findings.Select(finding => finding.Message));
        OnPropertyChanged(nameof(CanApplyReviewedRedactions));
    }
}
