using System.Globalization;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Application composition root for transport-neutral Studio presentation services.</summary>
internal sealed class StudioApplicationServices {
    private StudioApplicationServices(
        StudioDataPaths paths,
        StudioPreferencesService preferences,
        StudioCultureService cultures,
        IStudioLocalizer localizer,
        IStudioDiagnostics diagnostics) {
        Paths = paths;
        Preferences = preferences;
        Cultures = cultures;
        Localizer = localizer;
        Diagnostics = diagnostics;
    }

    internal StudioDataPaths Paths { get; }

    internal StudioPreferencesService Preferences { get; }

    internal StudioCultureService Cultures { get; }

    internal IStudioLocalizer Localizer { get; }

    internal IStudioDiagnostics Diagnostics { get; }

    internal static StudioApplicationServices CreateDefault() => Create(StudioDataPaths.CreateDefault());

    internal static StudioApplicationServices Create(StudioDataPaths paths) {
        ArgumentNullException.ThrowIfNull(paths);
        var preferences = new StudioPreferencesService(new JsonStudioPreferencesStore(paths.PreferencesPath));
        var cultures = new StudioCultureService();
        CultureInfo culture = cultures.Apply(preferences.Current.UiCulture);
        var localizer = new StudioLocalizer(culture);
        var diagnostics = new StudioDiagnostics(paths.DiagnosticsRoot);
        return new StudioApplicationServices(paths, preferences, cultures, localizer, diagnostics);
    }
}
