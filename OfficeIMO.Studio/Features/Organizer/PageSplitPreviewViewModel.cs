using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PageSplitPreviewViewModel : ObservableObject {
    private readonly IStudioLocalizer _localizer;
    private readonly Func<string, Task> _open;
    private readonly bool _provider;
    [ObservableProperty] private string _destination;
    [ObservableProperty] private string _pagesPerPart;
    [ObservableProperty] private string? _errorMessage;
    [ObservableProperty] private IReadOnlyList<PdfSplitPart> _parts = [];
    [ObservableProperty] private IReadOnlyList<PdfSplitFile> _files = [];
    [ObservableProperty] private IReadOnlyList<PageSplitOutputRow> _outputRows = [];
    [ObservableProperty] private PageSplitOutputRow? _selectedOutput;
    [ObservableProperty] private bool _hasResult;
    [ObservableProperty] private string? _summary;
    internal PageSplitPreviewViewModel(int pageCount, int pagesPerPart, string destination, bool provider,
        IStudioLocalizer localizer, Func<string, Task> open) {
        PageCount = pageCount; _destination = destination; _provider = provider; _localizer = localizer; _open = open;
        DestinationHint = localizer.Get(provider ? "Organizer.SplitProviderHint" : "Organizer.SplitLocalHint");
        _pagesPerPart = pagesPerPart.ToString(CultureInfo.InvariantCulture);
        UpdatePlan();
    }
    public int PageCount { get; }
    public string DestinationHint { get; }
    public bool IsPreview => !HasResult;
    public bool CanApply => IsPreview && Parts.Count > 0 && ErrorMessage is null;
    internal int PartSize { get; private set; }
    partial void OnPagesPerPartChanged(string value) => UpdatePlan();
    partial void OnHasResultChanged(bool value) { OnPropertyChanged(nameof(IsPreview)); OnPropertyChanged(nameof(CanApply)); }
    private void UpdatePlan() {
        Parts = []; PartSize = 0;
        try {
            if (!int.TryParse(PagesPerPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) || count < 1)
                throw new ArgumentException(_localizer.Get("Organizer.SplitWholeCount"));
            Parts = PdfSplitPlan.Create(PageCount, count).Parts;
            PartSize = count; ErrorMessage = null;
        } catch (Exception error) when (error is ArgumentException or InvalidOperationException) { ErrorMessage = error.Message; }
        OnPropertyChanged(nameof(CanApply));
    }
    internal void Complete(PdfSplitWorkflowResult result) {
        Files = result.Files; Summary = result.Summary; HasResult = true;
        OutputRows = result.Files.Select(file => new PageSplitOutputRow(
            Parts.FirstOrDefault(part => part.FirstSourcePage == file.FirstSourcePage)?.Name ?? file.Path,
            file.Path, file.PageCount)).ToArray();
        if (!_provider && result.Files.Count > 0) Destination = Path.GetDirectoryName(result.Files[0].Path) ?? Destination;
        ErrorMessage = string.Join(Environment.NewLine, result.Diagnostics.Where(item => item.Details.ContainsKey("recoveryPaths"))
            .Select(item => item.Details["recoveryPaths"]));
    }
    [RelayCommand]
    private async Task OpenFileAsync(string? path) {
        if (string.IsNullOrWhiteSpace(path) || !Files.Any(file => file.Path == path)) return;
        try { await _open(path).ConfigureAwait(true); }
        catch (Exception error) { ErrorMessage = error.Message; }
    }
}

public sealed record PageSplitOutputRow(string Name, string Path, int PageCount);
