using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PageExtractionPreviewViewModel : ObservableObject {
    private readonly Func<string, Task> _open;
    private readonly Func<string, Task> _reveal;
    private readonly IStudioLocalizer _localizer;
    [ObservableProperty] private string _pageRange;
    [ObservableProperty] private string? _errorMessage;
    [ObservableProperty] private IReadOnlyList<PageExtractionRow> _pages = [];
    [ObservableProperty] private bool _hasResult;
    [ObservableProperty] private string? _summary;
    [ObservableProperty] private string? _outputPath;
    [ObservableProperty] private bool _hasRecovery;
    internal PageExtractionPreviewViewModel(int pageCount, int[] pages, string destination, bool provider,
        IStudioLocalizer localizer, Func<string, Task> open, Func<string, Task> reveal) {
        PageCount = pageCount; Destination = destination; _localizer = localizer; _open = open; _reveal = reveal;
        DestinationHint = localizer.Get(provider ? "Organizer.ExtractProviderHint" : "Organizer.ExtractLocalHint");
        _pageRange = string.Join(',', pages); UpdatePlan();
    }
    public int PageCount { get; }
    public string Destination { get; }
    public string DestinationHint { get; }
    public bool IsPreview => !HasResult;
    public bool CanApply => IsPreview && ErrorMessage is null && Pages.Count > 0;
    public bool CanOpenOutput => HasResult && OutputPath is not null;
    public bool CanRevealOutput => CanOpenOutput && OfficeStorageIdentity.GetLocalPath(OutputPath!) is not null;
    public string PageCountLabel => _localizer.Format("Organizer.ExtractPageCount", PageCount, Pages.Count);
    internal int[] SelectedPages => Pages.Select(page => page.SourcePage).ToArray();
    partial void OnPageRangeChanged(string value) { if (IsPreview) UpdatePlan(); }
    private void UpdatePlan() {
        Pages = []; ErrorMessage = null;
        try {
            PdfPageSelection selection = PdfPageSelection.Parse(PageRange);
            if (selection.Ranges.Sum(range => (long)range.PageCount) > 100000)
                ErrorMessage = _localizer.Get("Organizer.ExtractPageLimit");
            else Pages = selection.Resolve(PageCount).Select((page, index) => new PageExtractionRow(index + 1, page)).ToArray();
        } catch (Exception error) when (error is ArgumentException or FormatException or OverflowException) { ErrorMessage = _localizer.Format("Organizer.ExtractInvalidRange", PageCount); }
        OnPropertyChanged(nameof(CanApply)); OnPropertyChanged(nameof(PageCountLabel));
    }
    internal void Complete(OfficeWorkflowResult result) {
        Summary = result.Summary; OutputPath = result.Succeeded ? result.OutputPath : null;
        HasRecovery = result.Recovery is not null; HasResult = true;
        ErrorMessage = result.Succeeded ? null : string.Join(Environment.NewLine, result.Diagnostics
            .Where(item => item.Severity != OfficeWorkflowDiagnosticSeverity.Information && item.Message != result.Summary)
            .Select(item => item.Message).Distinct());
        OnPropertyChanged(nameof(IsPreview)); OnPropertyChanged(nameof(CanApply));
        OnPropertyChanged(nameof(CanOpenOutput)); OnPropertyChanged(nameof(CanRevealOutput));
        OpenOutputCommand.NotifyCanExecuteChanged(); RevealOutputCommand.NotifyCanExecuteChanged();
    }
    [RelayCommand(CanExecute = nameof(CanOpenOutput))]
    private async Task OpenOutputAsync() {
        if (!CanOpenOutput) return;
        try { await _open(OutputPath!).ConfigureAwait(true); } catch (Exception error) { ErrorMessage = error.Message; }
    }
    [RelayCommand(CanExecute = nameof(CanRevealOutput))]
    private async Task RevealOutputAsync() {
        if (!CanRevealOutput) return;
        try { await _reveal(OfficeStorageIdentity.GetLocalPath(OutputPath!)!).ConfigureAwait(true); }
        catch (Exception error) { ErrorMessage = error.Message; }
    }
}

public sealed record PageExtractionRow(int OutputPage, int SourcePage);
