namespace OfficeIMO.Studio.Infrastructure.Preferences;

internal enum StudioThemePreference {
    System,
    Light,
    Dark
}

/// <summary>Versioned, user-scoped presentation preferences for OfficeIMO Studio.</summary>
internal sealed record StudioPreferences {
    internal const int CurrentSchemaVersion = 1;

    public int SchemaVersion { get; init; } = CurrentSchemaVersion;

    public string UiCulture { get; init; } = "en";

    public StudioThemePreference Theme { get; init; } = StudioThemePreference.System;

    internal StudioPreferences Normalize() {
        string culture = Infrastructure.Localization.StudioCultureCatalog.NormalizeOrDefault(UiCulture);
        StudioThemePreference theme = Enum.IsDefined(Theme) ? Theme : StudioThemePreference.System;
        return this with {
            SchemaVersion = CurrentSchemaVersion,
            UiCulture = culture,
            Theme = theme
        };
    }
}
