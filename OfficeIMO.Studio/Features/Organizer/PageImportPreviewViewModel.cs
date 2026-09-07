using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PageImportPreviewViewModel : ObservableObject {
    private readonly IStudioLocalizer _localizer;
    [ObservableProperty] private string _insertBeforePage;
    [ObservableProperty] private string? _errorMessage;
    [ObservableProperty] private int _importedPageCount;
    [ObservableProperty] private PageImportSourceViewModel? _selectedSource;
    internal PageImportPreviewViewModel(PdfImportPreparation preparation, int before, IStudioLocalizer localizer) {
        _localizer = localizer; TargetPageCount = preparation.TargetPageCount;
        _insertBeforePage = before.ToString(CultureInfo.InvariantCulture);
        Sources = new(preparation.Sources.Select((source, index) => new PageImportSourceViewModel(index, source.Name, source.PageCount, localizer)));
        foreach (var source in Sources) source.PropertyChanged += (_, _) => UpdatePlan();
        SelectedSource = Sources.FirstOrDefault();
        UpdatePlan();
    }
    public ObservableCollection<PageImportSourceViewModel> Sources { get; }
    public int TargetPageCount { get; }
    public string EndHint => _localizer.Format("Organizer.ImportEndHint", TargetPageCount + 1);
    public bool CanApply => ErrorMessage is null && ImportedPageCount > 0;
    internal int InsertBefore { get; private set; }
    internal PdfImportSelection[] Selections => Sources.Where(source => source.IsIncluded)
        .Select(source => new PdfImportSelection(source.SourceIndex, (int[])source.SelectedPages.Clone())).ToArray();
    partial void OnInsertBeforePageChanged(string value) => UpdatePlan();
    private void UpdatePlan() {
        ErrorMessage = null; ImportedPageCount = 0;
        if (!int.TryParse(InsertBeforePage, NumberStyles.Integer, CultureInfo.InvariantCulture, out int before) || before < 1 || before > TargetPageCount + 1)
            ErrorMessage = _localizer.Get("Organizer.WholePageRequired");
        else if (Sources.Any(source => source.IsIncluded && source.ErrorMessage is not null))
            ErrorMessage = _localizer.Get("Organizer.ImportFixRanges");
        else {
            long count = Sources.Sum(source => (long)source.SelectedPages.Length);
            if (count is < 1 or > PdfImportPreparation.MaximumImportedPages)
                ErrorMessage = _localizer.Get("Organizer.ImportPageLimit");
            else { InsertBefore = before; ImportedPageCount = (int)count; }
        }
        OnPropertyChanged(nameof(CanApply));
    }
    [RelayCommand]
    private void MoveSourceUp() => MoveSource(-1);
    [RelayCommand]
    private void MoveSourceDown() => MoveSource(1);
    private void MoveSource(int delta) {
        if (SelectedSource is null) return;
        int current = Sources.IndexOf(SelectedSource);
        int target = current + delta;
        if (current >= 0 && target >= 0 && target < Sources.Count) Sources.Move(current, target);
    }
}

public sealed partial class PageImportSourceViewModel : ObservableObject {
    [ObservableProperty] private bool _isIncluded = true;
    [ObservableProperty] private string _pageRange;
    [ObservableProperty] private string? _errorMessage;
    internal PageImportSourceViewModel(int index, string name, int count, IStudioLocalizer localizer) {
        SourceIndex = index; Name = name; PageCount = count; _pageRange = count == 1 ? "1" : "1-" + count.ToString(CultureInfo.InvariantCulture);
        RangeLabel = localizer.Format("Organizer.ImportRangeFor", name);
        UpdateSelection();
    }
    internal int SourceIndex { get; }
    public string Name { get; }
    public string RangeLabel { get; }
    public int PageCount { get; }
    internal int[] SelectedPages { get; private set; } = [];
    partial void OnPageRangeChanged(string value) => UpdateSelection();
    partial void OnIsIncludedChanged(bool value) => UpdateSelection();
    private void UpdateSelection() {
        SelectedPages = []; ErrorMessage = null;
        if (!IsIncluded) return;
        try {
            PdfPageSelection selection = PdfPageSelection.Parse(PageRange);
            if (selection.Ranges.Sum(range => (long)range.PageCount) > PdfImportPreparation.MaximumImportedPages)
                throw new ArgumentException("An import cannot exceed 100,000 selected pages.");
            SelectedPages = selection.Resolve(PageCount).ToArray();
        } catch (Exception error) when (error is ArgumentException or FormatException or OverflowException) { ErrorMessage = error.Message; }
    }
}
