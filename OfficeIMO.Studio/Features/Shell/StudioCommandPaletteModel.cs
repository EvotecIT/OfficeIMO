using CommunityToolkit.Mvvm.ComponentModel;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed partial class StudioCommandPaletteModel : ObservableObject {
    private readonly StudioCommandCatalog _catalog;
    [ObservableProperty] private string _query = string.Empty;
    [ObservableProperty] private StudioCommandItem? _selectedCommand;

    internal StudioCommandPaletteModel(StudioCommandCatalog catalog) {
        _catalog = catalog;
        SelectedCommand = Results.FirstOrDefault();
    }

    /// <summary>
    /// Matching commands ranked for action: recently run and available commands first, commands that
    /// cannot run yet last so their reason stays discoverable without crowding the useful choices.
    /// </summary>
    public IReadOnlyList<StudioCommandItem> Results {
        get {
            IReadOnlyList<StudioCommandItem> matches = _catalog.Search(Query);
            return matches
                .Select((item, index) => (item, index))
                .OrderBy(entry => entry.item.IsAvailable ? 0 : 1)
                .ThenBy(entry => _catalog.RecentRank(entry.item.Id))
                .ThenBy(entry => entry.index)
                .Select(entry => entry.item)
                .ToArray();
        }
    }

    public bool HasResults => Results.Count > 0;

    partial void OnQueryChanged(string value) {
        OnPropertyChanged(nameof(Results));
        OnPropertyChanged(nameof(HasResults));
        SelectedCommand = Results.FirstOrDefault();
    }

    internal void Refresh() {
        OnPropertyChanged(nameof(Results));
        OnPropertyChanged(nameof(HasResults));
        SelectedCommand = Results.FirstOrDefault();
    }
}
