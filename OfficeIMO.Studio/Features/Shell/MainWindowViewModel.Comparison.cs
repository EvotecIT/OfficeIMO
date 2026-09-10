using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>One changed, resized, added, or removed page in a reviewed comparison.</summary>
public sealed record ComparisonDifference(int PageNumber, string Label, PdfVisualPageComparison? Comparison) {
    /// <inheritdoc />
    public override string ToString() => Label;
}

public sealed partial class MainWindowViewModel {
    private const int MaximumComparisonPages = 100;
    private const long MaximumComparisonRasterPixels = 4_000_000;
    private CancellationTokenSource? _comparisonReportCancellation;
    [ObservableProperty] private bool _isComparingPages;
    [ObservableProperty] private string _comparisonSummary = string.Empty;
    [ObservableProperty] private IReadOnlyList<ComparisonDifference> _comparisonDifferences = [];
    [ObservableProperty] private ComparisonDifference? _selectedComparisonDifference;
    [ObservableProperty] private Bitmap? _comparisonDifferenceImage;
    [ObservableProperty] private bool _showComparisonDifferenceImage;
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
            SelectedPage = Pages.ElementAtOrDefault(value.PageNumber - 1);
            ComparisonSelectedPage = ComparisonPages.ElementAtOrDefault(value.PageNumber - 1);
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
        long revision = workspace.Revision;
        IsComparingPages = true;
        ComparisonSummary = ComparisonText("Comparing", "Comparing page appearance…");
        try {
            if (HasFormDrafts) throw new InvalidOperationException(ComparisonText("FormDrafts", "Apply pending form values before comparing page appearance."));
            if (primary.Pages.Count > MaximumComparisonPages || actual.Pages.Count > MaximumComparisonPages) {
                throw new InvalidOperationException(ComparisonText("PageLimit", "Comparison is limited to 100 pages per document. Extract a smaller range for review."));
            }
            PdfVisualComparisonReport report = await workspace.RunNonDetachableCpuWorkAsync(() => primary.CompareTo(actual,
                new PdfVisualComparisonOptions {
                    Scale = 1, MaxPages = MaximumComparisonPages, MaxPixelsPerImage = MaximumComparisonRasterPixels,
                    // Each pair charges the two source rasters and the difference raster.
                    MaxTotalPixels = 3L * MaximumComparisonPages * MaximumComparisonRasterPixels,
                    MaxTotalOutputBytes = 64 * 1024 * 1024
                }, operation.Token), operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            if (_disposed || !ReferenceEquals(_session, primary) || !ReferenceEquals(_comparisonSession, actual) ||
                !ReferenceEquals(_workspace, workspace) || workspace.Revision != revision) return;
            var differences = new List<ComparisonDifference>();
            foreach (PdfVisualPageComparison page in report.Pages.Where(page => !page.IsMatch || page.HasSizeDifference)) {
                string kind = page.HasSizeDifference ? ComparisonText("Resized", "Page size changed") : ComparisonText("Changed", "Appearance changed");
                differences.Add(new(page.PageNumber, $"{page.PageNumber} · {kind}", page));
            }
            for (int page = Math.Min(report.ExpectedPageCount, report.ActualPageCount) + 1;
                page <= Math.Max(report.ExpectedPageCount, report.ActualPageCount); page++) {
                string kind = page > report.ExpectedPageCount ? ComparisonText("Added", "Only in comparison") : ComparisonText("Removed", "Only in current document");
                differences.Add(new(page, $"{page} · {kind}", null));
            }
            ComparisonDifferences = differences.OrderBy(difference => difference.PageNumber).ToArray();
            SelectedComparisonDifference = ComparisonDifferences.FirstOrDefault();
            ComparisonSummary = report.IsMatch
                ? ComparisonText("Match", "No rendered differences found.")
                : _localizer.FormatOrDefault("Comparison.ChangedCount", "{0:N0} page(s) differ. Red pixels show appearance changes; pages are paired by page number.", differences.Count);
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            if (ReferenceEquals(_comparisonReportCancellation, operation)) ComparisonSummary = ComparisonText("Cancelled", "Comparison cancelled.");
        } catch (Exception error) {
            if (ReferenceEquals(_comparisonReportCancellation, operation)) ComparisonSummary = error.Message;
        } finally {
            if (ReferenceEquals(_comparisonReportCancellation, operation)) {
                _comparisonReportCancellation = null;
                IsComparingPages = false;
            }
        }
    }

    [RelayCommand] private void CancelPageComparison() => _comparisonReportCancellation?.Cancel();
    [RelayCommand] private void NextComparisonDifference() => MoveComparisonDifference(1);
    [RelayCommand] private void PreviousComparisonDifference() => MoveComparisonDifference(-1);
    private void SynchronizeDifferenceToPage(int? pageNumber) {
        if (_synchronizingComparison) return;
        SelectedComparisonDifference = ComparisonDifferences.FirstOrDefault(difference => difference.PageNumber == pageNumber);
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
        SelectedComparisonDifference = null;
        ComparisonDifferences = [];
        ComparisonSummary = ComparisonText("Ready", "Compare the current document with this PDF. Pages are paired by page number.");
    }
    private string ComparisonText(string key, string fallback) => _localizer.GetOrDefault("Comparison." + key, fallback);
}
