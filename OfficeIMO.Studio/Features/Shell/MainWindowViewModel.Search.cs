using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private long _searchGeneration;
    private bool _searchCompleted;
    private string? _searchStatus;

    [ObservableProperty]
    private string _searchQuery = string.Empty;

    [ObservableProperty]
    private PdfSearchHit? _selectedSearchResult;

    public ObservableCollection<PdfSearchHit> SearchResults { get; } = new();

    public bool HasSearchResults => SearchResults.Count > 0;

    public string SearchPosition => !CanSearchDocument ? UiText("Capability.SearchRestricted") : HasSearchResults
        ? UiFormat("Search.Position", SelectedSearchResult is null ? 0 : SearchResults.IndexOf(SelectedSearchResult) + 1, SearchResults.Count)
        : UiText(_searchCompleted ? "Search.NoResults" : "Search.Prompt");

    partial void OnSearchQueryChanged(string value) => ClearSearchResults();

    partial void OnSelectedSearchResultChanged(PdfSearchHit? value) {
        foreach (var page in Pages) page.ActiveSearchHighlight = page.PageNumber == value?.PageNumber ? value.Bounds : null;
        OnPropertyChanged(nameof(SearchPosition));
        if (value is not null) NavigateToPage(value.PageNumber);
    }

    private void ClearSearchResults() {
        if (_searchStatus is not null && OperationStatus == _searchStatus) OperationStatus = null;
        _searchStatus = null;
        _searchGeneration++;
        _searchCompleted = false;
        SelectedSearchResult = null;
        SearchResults.Clear();
        foreach (var page in Pages) {
            page.SearchHighlights = Array.Empty<Avalonia.Rect>();
            page.ActiveSearchHighlight = null;
        }
        NotifySearchResultsChanged();
    }

    private void NotifySearchResultsChanged() {
        OnPropertyChanged(nameof(HasSearchResults));
        OnPropertyChanged(nameof(SearchPosition));
        NextSearchResultCommand.NotifyCanExecuteChanged();
        PreviousSearchResultCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(HasSearchResults))]
    private void NextSearchResult() => MoveSearchResult(1);

    [RelayCommand(CanExecute = nameof(HasSearchResults))]
    private void PreviousSearchResult() => MoveSearchResult(-1);

    private void MoveSearchResult(int direction) {
        if (!HasSearchResults) return;
        int index = SelectedSearchResult is null ? (direction > 0 ? -1 : 0) : SearchResults.IndexOf(SelectedSearchResult);
        SelectedSearchResult = SearchResults[(index + direction + SearchResults.Count) % SearchResults.Count];
    }

    [RelayCommand]
    private void ClearSearch() {
        SearchQuery = string.Empty;
        ClearSearchResults();
    }

    [RelayCommand(CanExecute = nameof(CanSearchDocument))]
    private async Task SearchAsync(CancellationToken cancellationToken) {
        if (IsWorkspaceBusy) return;
        ClearSearchResults();
        PdfDocumentSession? session = _session;
        string query = SearchQuery;
        long generation = _searchGeneration;
        if (session is null || string.IsNullOrWhiteSpace(query)) return;
        if (!CanSearchDocument) { ErrorMessage = UiText("Capability.SearchRestricted"); return; }
        _searchStatus = OperationStatus = UiText("Workspace.SearchingDocument");
        bool succeeded = await RunStandaloneAsync(async token => {
            CancellationTokenSource? attempt = _operationCancellation;
            var progress = new Progress<double>(fraction => {
                if (IsWorkspaceBusy && ReferenceEquals(attempt, _operationCancellation)) OperationProgressFraction = Math.Clamp(fraction, 0D, 1D);
            });
            var results = await session.SearchAsync(query, token, progress).ConfigureAwait(true);
            if (generation != _searchGeneration || !ReferenceEquals(session, _session)) return;
            foreach (var result in results) SearchResults.Add(result.WithLocalizer(_localizer));
            var pageMatches = SearchResults.GroupBy(hit => hit.PageNumber).ToDictionary(group => group.Key, group => group.Select(hit => hit.Bounds).ToArray());
            foreach (var page in Pages) page.SearchHighlights = pageMatches.TryGetValue(page.PageNumber, out var highlights) ? highlights : Array.Empty<Avalonia.Rect>();
            _searchCompleted = true;
            NotifySearchResultsChanged();
            SelectedSearchResult = SearchResults.FirstOrDefault();
            OperationProgressFraction = 1D;
        }, cancellationToken).ConfigureAwait(true);
        if (succeeded && generation == _searchGeneration && ReferenceEquals(session, _session))
            _searchStatus = OperationStatus = HasSearchResults ? UiFormat("Search.MatchCount", SearchResults.Count) : UiText("Workspace.NoMatches");
    }
}
