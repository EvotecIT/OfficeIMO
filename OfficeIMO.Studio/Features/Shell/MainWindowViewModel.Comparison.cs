using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>One changed, resized, added, or removed page in a reviewed comparison.</summary>
public sealed record ComparisonDifference(int PageNumber, string Label, PdfVisualPageComparison? Comparison, int? ActualPageNumber = null, bool ExpectedMissing = false) {
    /// <inheritdoc />
    public override string ToString() => Label;
}

public sealed partial class MainWindowViewModel {
    private const int MaximumComparisonPages = 100;
    private const long MaximumComparisonRasterPixels = 4_000_000;
    private CancellationTokenSource? _comparisonReportCancellation;
    private PdfVisualComparisonReport? _comparisonReport;
    private long _comparisonReportRevision;
    [ObservableProperty] private string _comparisonExpectedRange = string.Empty;
    [ObservableProperty] private string _comparisonActualRange = string.Empty;
    public bool CanExportComparisonReport => _comparisonReport is not null && !IsComparingPages && !IsWorkspaceBusy && !HasFormDrafts;
    partial void OnComparisonExpectedRangeChanged(string value) => ClearComparisonDifferences();
    partial void OnComparisonActualRangeChanged(string value) => ClearComparisonDifferences();
    [ObservableProperty] private bool _isComparingPages;
    [ObservableProperty] private string _comparisonSummary = string.Empty;
    [ObservableProperty] private IReadOnlyList<ComparisonDifference> _comparisonDifferences = [];
    [ObservableProperty] private ComparisonDifference? _selectedComparisonDifference;
    [ObservableProperty] private Bitmap? _comparisonDifferenceImage;
    [ObservableProperty] private bool _showComparisonDifferenceImage;
    public string ComparisonExpectedRangeLabel => ComparisonText("ExpectedRange", "Current pages");
    public string ComparisonActualRangeLabel => ComparisonText("ActualRange", "Comparison pages");
    public string ComparisonRangeHint => ComparisonText("RangeHint", "All pages, or e.g. 5-8,last");
    public string ComparisonRangeDescription => ComparisonText("RangeDescription", "Up to 100 selected pages per document, paired in order. No semantic or moved-page detection.");
    public string ComparisonExportReportLabel => ComparisonText("ExportReport", "Export report");
    public bool HasComparisonDifferences => ComparisonDifferences.Count > 0;
    public bool HasComparisonDifferenceImage => ComparisonDifferenceImage is not null;
    public double ComparisonDifferenceWidth => (ComparisonDifferenceImage?.PixelSize.Width ?? 0) * Zoom;
    public double ComparisonDifferenceHeight => (ComparisonDifferenceImage?.PixelSize.Height ?? 0) * Zoom;

    partial void OnComparisonDifferenceImageChanged(Bitmap? value) {
        OnPropertyChanged(nameof(HasComparisonDifferenceImage));
        OnPropertyChanged(nameof(ComparisonDifferenceWidth));
        OnPropertyChanged(nameof(ComparisonDifferenceHeight));
    }

    partial void OnComparisonDifferencesChanged(IReadOnlyList<ComparisonDifference> value) =>
        OnPropertyChanged(nameof(HasComparisonDifferences));

    partial void OnSelectedComparisonDifferenceChanged(ComparisonDifference? value) {
        ComparisonDifferenceImage?.Dispose();
        ComparisonDifferenceImage = null;
        ShowComparisonDifferenceImage = false;
        if (value is null) return;
        _synchronizingComparison = true;
        try {
            SelectedPage = value.ExpectedMissing ? null : Pages.ElementAtOrDefault(value.PageNumber - 1);
            ComparisonSelectedPage = value.ActualPageNumber is { } actual ? ComparisonPages.ElementAtOrDefault(actual - 1) : null;
        } finally { _synchronizingComparison = false; }
        if (value.Comparison is { } page) {
            using var stream = new MemoryStream(page.DiffPng, writable: false);
            ComparisonDifferenceImage = new Bitmap(stream);
            ShowComparisonDifferenceImage = true;
        }
    }

    [RelayCommand]
    private async Task ComparePagesAsync(CancellationToken token) {
        if (_disposed || IsWorkspaceBusy || IsComparingPages || _workspace is not { } workspace ||
            _session is not { } primary || _comparisonSession is not { } actual) return;
        ClearComparisonDifferences();
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _comparisonReportCancellation = operation;
        OnPropertyChanged(nameof(CanExportComparisonReport));
        long revision = workspace.Revision;
        IsComparingPages = true;
        ComparisonSummary = ComparisonText("Comparing", "Comparing page appearance…");
        try {
            if (HasFormDrafts) throw new InvalidOperationException(ComparisonText("FormDrafts", "Apply pending form values before comparing page appearance."));
            PdfPageSelector? expectedPages = ParseComparisonRange(ComparisonExpectedRange);
            PdfPageSelector? actualPages = ParseComparisonRange(ComparisonActualRange);
            PdfVisualComparisonReport report = await workspace.RunNonDetachableCpuWorkAsync(() => primary.CompareTo(actual,
                new PdfVisualComparisonOptions {
                    ExpectedPages = expectedPages, ActualPages = actualPages,
                    Scale = 1, MaxPages = MaximumComparisonPages, MaxPixelsPerImage = MaximumComparisonRasterPixels,
                    MaxTotalPixels = 200_000_000,
                    MaxTotalOutputBytes = 64 * 1024 * 1024
                }, operation.Token), operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            if (_disposed || !ReferenceEquals(_session, primary) || !ReferenceEquals(_comparisonSession, actual) ||
                !ReferenceEquals(_workspace, workspace) || workspace.Revision != revision) return;
            var differences = new List<ComparisonDifference>();
            foreach (PdfVisualPageComparison page in report.Pages.Where(page => !page.IsMatch || page.HasSizeDifference)) {
                string kind = page.HasSizeDifference ? ComparisonText("Resized", "Page size changed") : ComparisonText("Changed", "Appearance changed");
                string pair = page.PageNumber == page.ActualPageNumber ? page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : $"{page.PageNumber} / {page.ActualPageNumber}";
                differences.Add(new(page.PageNumber, $"{pair} · {kind}", page, page.ActualPageNumber));
            }
            foreach (int page in report.UnmatchedExpectedPageNumbers)
                differences.Add(new(page, $"{page} · {(report.IsSelectedScope ? ComparisonText("UnmatchedExpected", "Unmatched current selection") : ComparisonText("Removed", "Only in current document"))}", null));
            foreach (int page in report.UnmatchedActualPageNumbers)
                differences.Add(new(page, $"{page} · {(report.IsSelectedScope ? ComparisonText("UnmatchedActual", "Unmatched comparison selection") : ComparisonText("Added", "Only in comparison"))}", null, page, ExpectedMissing: true));
            _comparisonReport = report;
            _comparisonReportRevision = revision;
            ComparisonDifferences = differences.ToArray();
            SelectedComparisonDifference = ComparisonDifferences.FirstOrDefault();
            ComparisonSummary = report.IsSelectedScope
                ? _localizer.FormatOrDefault("Comparison.SelectedSummary", "{0:N0} paired pages; {1:N0} differing pairs; {2:N0} unmatched. Selections pair in order.",
                    report.Pages.Count, report.Pages.Count(page => !page.IsMatch), report.UnmatchedExpectedPageNumbers.Count + report.UnmatchedActualPageNumbers.Count)
                : report.IsMatch ? ComparisonText("Match", "No rendered differences found.")
                : _localizer.FormatOrDefault("Comparison.ChangedCount", "{0:N0} page(s) differ. Red pixels show appearance changes; pages are paired by page number.", differences.Count);
            OnPropertyChanged(nameof(CanExportComparisonReport));
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            if (ReferenceEquals(_comparisonReportCancellation, operation)) ComparisonSummary = ComparisonText("Cancelled", "Comparison cancelled.");
        } catch (Exception error) {
            if (ReferenceEquals(_comparisonReportCancellation, operation)) ComparisonSummary = error.Message;
        } finally {
            if (ReferenceEquals(_comparisonReportCancellation, operation)) {
                _comparisonReportCancellation = null;
                IsComparingPages = false;
                OnPropertyChanged(nameof(CanExportComparisonReport));
            }
        }
    }

    [RelayCommand] private void CancelPageComparison() => _comparisonReportCancellation?.Cancel();
    [RelayCommand] private void NextComparisonDifference() => MoveComparisonDifference(1);
    [RelayCommand] private void PreviousComparisonDifference() => MoveComparisonDifference(-1);
    private static PdfPageSelector? ParseComparisonRange(string value) => string.IsNullOrWhiteSpace(value)
        ? null : value.Length <= 4096 ? PdfPageSelector.Parse(value) : throw new ArgumentException("Comparison page selection cannot exceed 4096 characters.");

    private int? GetComparisonPartner(int page, bool actualSide) {
        if (_comparisonReport is null) return page;
        PdfVisualPageComparison? pair = _comparisonReport.Pages.FirstOrDefault(item =>
            actualSide ? item.ActualPageNumber == page : item.PageNumber == page);
        return pair is null ? null : actualSide ? pair.PageNumber : pair.ActualPageNumber;
    }

    private void SynchronizeDifferenceToPage(int? pageNumber) {
        if (_synchronizingComparison) return;
        SelectedComparisonDifference = ComparisonDifferences.FirstOrDefault(difference => !difference.ExpectedMissing && difference.PageNumber == pageNumber);
    }

    private void MoveComparisonDifference(int delta) {
        if (ComparisonDifferences.Count == 0) return;
        if (SelectedComparisonDifference is null) {
            int page = SelectedPage?.PageNumber ?? ComparisonSelectedPage?.PageNumber ?? 0;
            SelectedComparisonDifference = delta > 0
                ? ComparisonDifferences.FirstOrDefault(difference => difference.PageNumber > page) ?? ComparisonDifferences[^1]
                : ComparisonDifferences.LastOrDefault(difference => difference.PageNumber < page) ?? ComparisonDifferences[0];
        } else {
            int current = ComparisonDifferences.ToList().IndexOf(SelectedComparisonDifference);
            SelectedComparisonDifference = ComparisonDifferences[Math.Clamp(current + delta, 0, ComparisonDifferences.Count - 1)];
        }
    }

    private void ClearComparisonDifferences() {
        _comparisonReportCancellation?.Cancel();
        _comparisonReportCancellation = null;
        IsComparingPages = false;
        _comparisonReport = null;
        OnPropertyChanged(nameof(CanExportComparisonReport));
        SelectedComparisonDifference = null;
        ComparisonDifferences = [];
        ComparisonSummary = ComparisonText("Ready", "Compare the current document with this PDF. Pages are paired by page number.");
    }
    private string ComparisonText(string key, string fallback) => _localizer.GetOrDefault("Comparison." + key, fallback);
}
