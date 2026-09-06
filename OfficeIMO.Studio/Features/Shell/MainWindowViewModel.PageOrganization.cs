using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Organizer;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private readonly Func<PageMovePreviewViewModel, Task<bool>> _reviewPageMove;
    private bool _reviewingPageMove;
    [ObservableProperty] private string _organizerPageRange = string.Empty;
    [ObservableProperty] private bool _isOrganizerRangeExpanded;
    [ObservableProperty] private string? _organizerRangeError;

    [RelayCommand]
    private void SelectPageRange() {
        if (_workspace is null || IsWorkspaceBusy) return;
        try {
            var selected = PdfPageSelection.Parse(OrganizerPageRange).Resolve(_workspace.Pages.Count).ToHashSet();
            SetOrganizerSelection(OrganizerPages.Where(page => selected.Contains(page.PageNumber)).ToArray());
            OrganizerRangeError = null;
            IsOrganizerRangeExpanded = false;
        } catch (ArgumentException exception) {
            OrganizerRangeError = exception.Message;
        } catch (FormatException exception) {
            OrganizerRangeError = exception.Message;
        } catch (OverflowException exception) {
            OrganizerRangeError = exception.Message;
        }
    }

    [RelayCommand]
    private async Task MoveSelectedToAsync(CancellationToken token) {
        if (_workspace is null || !CanMutateSelection || _reviewingPageMove) return;
        var workspace = _workspace;
        long revision = workspace.Revision;
        int[] selected = GetSelectedPages();
        var preview = new PageMovePreviewViewModel(workspace.Pages.Count, selected, _localizer);
        _reviewingPageMove = true;
        try {
            if (!await _reviewPageMove(preview).ConfigureAwait(true) || preview.Plan is null) return;
            if (!ReferenceEquals(_workspace, workspace) || workspace.Revision != revision || !CanMutatePages) {
                ErrorMessage = UiText("Organizer.StalePreview");
                return;
            }
            await ApplyPageReorderAsync(preview.Plan, selected, token).ConfigureAwait(true);
        } finally { _reviewingPageMove = false; }
    }

    private async Task ApplyPageReorderAsync(PdfPageReorderPlan plan, int[] selected, CancellationToken token) {
        if (_workspace is null || !CanMutatePages || !plan.HasChanges) return;
        var workspace = _workspace;
        int readingPage = plan.GetOutputPageNumber(SelectedPage?.PageNumber ?? 1);
        int[] selectionAfter = selected.Select(plan.GetOutputPageNumber).OrderBy(page => page).ToArray();
        if (await RunMutationAsync(cancellation => workspace.ReorderAsync(plan.SourcePageNumbers, cancellation, CreateProgress()),
            token, selectionAfter).ConfigureAwait(true) && ReferenceEquals(_workspace, workspace)) NavigateToPage(readingPage);
    }
}
