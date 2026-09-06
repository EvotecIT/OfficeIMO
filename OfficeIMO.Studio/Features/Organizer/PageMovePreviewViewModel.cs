using CommunityToolkit.Mvvm.ComponentModel;
using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PageMovePreviewViewModel : ObservableObject {
    private readonly int[] _selectedPages;
    private readonly IStudioLocalizer _localizer;
    [ObservableProperty] private string _insertBeforePage;
    [ObservableProperty] private string? _errorMessage;
    [ObservableProperty] private IReadOnlyList<PageMovePreviewRow> _rows = [];

    internal PageMovePreviewViewModel(int pageCount, int[] selectedPages, IStudioLocalizer localizer) {
        PageCount = pageCount;
        _selectedPages = (int[])selectedPages.Clone();
        _localizer = localizer;
        _insertBeforePage = (pageCount + 1).ToString(CultureInfo.InvariantCulture);
        UpdatePlan();
    }

    public int PageCount { get; }
    public int MaximumDestination => PageCount + 1;
    public string EndHint => _localizer.Format("Organizer.EndHint", MaximumDestination);
    public string SelectedPagesLabel => string.Join(", ", _selectedPages);
    public bool CanApply => Plan?.HasChanges == true;
    internal PdfPageReorderPlan? Plan { get; private set; }

    partial void OnInsertBeforePageChanged(string value) => UpdatePlan();

    private void UpdatePlan() {
        try {
            if (!int.TryParse(InsertBeforePage, NumberStyles.Integer, CultureInfo.InvariantCulture, out int destination) ||
                destination < 1 || destination > MaximumDestination) {
                Plan = null;
                Rows = [];
                ErrorMessage = _localizer.Get("Organizer.WholePageRequired");
                OnPropertyChanged(nameof(CanApply));
                return;
            }
            Plan = PdfPageReorderPlan.Move(PageCount, destination, _selectedPages);
            var selected = _selectedPages.ToHashSet();
            Rows = Plan.SourcePageNumbers.Select((source, index) => new PageMovePreviewRow(
                _localizer.Format("Organizer.OutputPage", index + 1, source), selected.Contains(source))).ToArray();
            ErrorMessage = Plan.HasChanges ? null : _localizer.Get("Organizer.UnchangedOrder");
        } catch (ArgumentException exception) {
            Plan = null;
            Rows = [];
            ErrorMessage = exception.Message;
        }
        OnPropertyChanged(nameof(CanApply));
    }
}

public sealed record PageMovePreviewRow(string Label, bool IsMoved);
