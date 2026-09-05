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

    public IReadOnlyList<StudioCommandItem> Results => _catalog.Search(Query);
    public bool HasResults => Results.Count > 0;

    partial void OnQueryChanged(string value) {
        OnPropertyChanged(nameof(Results));
        OnPropertyChanged(nameof(HasResults));
        SelectedCommand = Results.FirstOrDefault();
    }
}
