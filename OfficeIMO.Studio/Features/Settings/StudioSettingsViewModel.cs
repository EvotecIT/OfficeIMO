using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Settings;

internal sealed record StudioCultureChoice(string Name, string Label);

internal sealed record StudioThemeChoice(StudioThemePreference Value, string Label, string Description);

/// <summary>Presents application-wide preferences and privacy-bounded support information.</summary>
internal sealed partial class StudioSettingsViewModel : ObservableObject, IDisposable {
    private readonly StudioPreferencesService _preferences;
    private readonly IStudioLocalizer _localizer;
    private bool _synchronizing;

    internal StudioSettingsViewModel(
        StudioPreferencesService preferences,
        IStudioLocalizer localizer,
        IStudioDiagnostics diagnostics) {
        _preferences = preferences ?? throw new ArgumentNullException(nameof(preferences));
        _localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
        ArgumentNullException.ThrowIfNull(diagnostics);

        Cultures = StudioCultureCatalog.Available
            .Select(culture => new StudioCultureChoice(
                culture.Name,
                culture.Name == StudioCultureCatalog.PseudoCulture
                    ? _localizer.Get("Settings.PseudolocalizedEnglish")
                    : culture.NativeName))
            .ToArray();
        Themes = [
            new(StudioThemePreference.System, _localizer.Get("Settings.ThemeSystem"), _localizer.Get("Settings.ThemeSystemDescription")),
            new(StudioThemePreference.Light, _localizer.Get("Settings.ThemeLight"), _localizer.Get("Settings.ThemeLightDescription")),
            new(StudioThemePreference.Dark, _localizer.Get("Settings.ThemeDark"), _localizer.Get("Settings.ThemeDarkDescription")),
            new(StudioThemePreference.HighContrast, _localizer.Get("Settings.ThemeHighContrast"), _localizer.Get("Settings.ThemeHighContrastDescription"))
        ];

        StudioSupportSnapshot support = diagnostics.CreateSupportSnapshot();
        ProductVersion = support.Version;
        RuntimeDescription = $"{support.Runtime} · {support.Architecture}";
        OperatingSystem = support.OperatingSystem;
        DiagnosticsDirectory = diagnostics.DirectoryPath;
        PrivacyNotice = _localizer.Get("Settings.DiagnosticsPrivacyNotice");
        SynchronizeFromPreferences();
        _preferences.Changed += OnPreferencesChanged;
    }

    internal IReadOnlyList<StudioCultureChoice> Cultures { get; }

    internal IReadOnlyList<StudioThemeChoice> Themes { get; }

    internal string ProductVersion { get; }

    internal string RuntimeDescription { get; }

    internal string OperatingSystem { get; }

    internal string DiagnosticsDirectory { get; }

    internal string PrivacyNotice { get; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(RestartRequired))]
    private StudioCultureChoice _selectedCulture = null!;

    [ObservableProperty]
    private StudioThemeChoice _selectedTheme = null!;

    internal bool RestartRequired =>
        !string.Equals(SelectedCulture.Name, _localizer.Culture.Name, StringComparison.OrdinalIgnoreCase);

    partial void OnSelectedCultureChanged(StudioCultureChoice value) {
        if (_synchronizing || value is null) return;
        _preferences.Update(current => current with { UiCulture = value.Name });
        OnPropertyChanged(nameof(RestartRequired));
    }

    partial void OnSelectedThemeChanged(StudioThemeChoice value) {
        if (_synchronizing || value is null) return;
        _preferences.Update(current => current with { Theme = value.Value });
    }

    public void Dispose() => _preferences.Changed -= OnPreferencesChanged;

    private void OnPreferencesChanged(object? sender, EventArgs eventArgs) => SynchronizeFromPreferences();

    private void SynchronizeFromPreferences() {
        _synchronizing = true;
        try {
            SelectedCulture = Cultures.FirstOrDefault(choice =>
                string.Equals(choice.Name, _preferences.Current.UiCulture, StringComparison.OrdinalIgnoreCase)) ?? Cultures[0];
            SelectedTheme = Themes.First(choice => choice.Value == _preferences.Current.Theme);
        } finally {
            _synchronizing = false;
        }
        OnPropertyChanged(nameof(RestartRequired));
    }
}
