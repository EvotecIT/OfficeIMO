using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionWorkbenchViewModel {
    private IReadOnlyList<string> _unmatchedInputs = [];

    public bool HasUnmatchedInputs => _unmatchedInputs.Count > 0;

    public string UnmatchedInputMessage => MatchingInputRoutes.Count == 0
        ? _localizer.FormatOrDefault("Conversion.Intake.Unsupported", "No available conversion accepts these selected file types. Add a supported document or dismiss this selection.")
        : _localizer.FormatOrDefault(
        "Conversion.Intake.Mismatch",
        "{0:N0} selected file(s) cannot use {1}. Choose a conversion below to add them, or dismiss this selection.",
        _unmatchedInputs.Count, SelectedRoute.Label);

    public IReadOnlyList<ConversionRouteChoice> MatchingInputRoutes => Routes.Where(route =>
        _unmatchedInputs.Any(path => route.Route.SourceExtensions.Any(extension =>
            string.Equals(NormalizeExtension(extension), Path.GetExtension(_storage?.Describe(path).Name ?? path),
                StringComparison.OrdinalIgnoreCase)))).ToArray();

    private void SetUnmatchedInputs(IReadOnlyList<string> paths) {
        _unmatchedInputs = paths;
        OnPropertyChanged(nameof(HasUnmatchedInputs));
        OnPropertyChanged(nameof(UnmatchedInputMessage));
        OnPropertyChanged(nameof(MatchingInputRoutes));
    }

    [RelayCommand(CanExecute = nameof(CanEditQueue))]
    private void UseInputRoute(ConversionRouteChoice? route) {
        if (route is null || !Routes.Contains(route) || !HasUnmatchedInputs) return;
        SelectedRoute = route;
        AddPaths(_unmatchedInputs);
    }

    [RelayCommand(CanExecute = nameof(CanEditQueue))]
    private void DismissUnmatchedInputs() => SetUnmatchedInputs([]);
}
