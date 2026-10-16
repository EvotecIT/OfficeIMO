namespace OfficeIMO.Studio.Infrastructure.Preferences;

internal enum StudioThemePreference {
    System,
    Light,
    Dark,
    HighContrast
}

internal enum StudioDensityPreference {
    Comfortable,
    Compact
}

/// <summary>Versioned, user-scoped presentation preferences for OfficeIMO Studio.</summary>
internal sealed record StudioPreferences {
    internal const int CurrentSchemaVersion = 1;

    public int SchemaVersion { get; init; } = CurrentSchemaVersion;

    public string UiCulture { get; init; } = "en";

    public string? OcrExecutablePath { get; init; }

    public StudioThemePreference Theme { get; init; } = StudioThemePreference.System;

    public StudioDensityPreference Density { get; init; } = StudioDensityPreference.Comfortable;

    public bool RememberSession { get; init; } = true;

    public bool CreateRecoverySnapshots { get; init; } = true;

    public bool RememberDocumentHistory { get; init; } = true;

    /// <summary>Shows first-page previews of local recent documents on Home; turn off for private work.</summary>
    public bool ShowRecentPreviews { get; init; } = true;

    public StudioAiConnectionPreference? AiConnection { get; init; }

    public StudioWindowPlacement? MainWindow { get; init; }

    internal StudioPreferences Normalize() {
        string culture = Infrastructure.Localization.StudioCultureCatalog.NormalizeOrDefault(UiCulture);
        StudioThemePreference theme = Enum.IsDefined(Theme) ? Theme : StudioThemePreference.System;
        return this with {
            SchemaVersion = CurrentSchemaVersion,
            UiCulture = culture,
            OcrExecutablePath = string.IsNullOrWhiteSpace(OcrExecutablePath) ? null : OcrExecutablePath.Trim(),
            Theme = theme,
            Density = Enum.IsDefined(Density) ? Density : StudioDensityPreference.Comfortable,
            MainWindow = MainWindow?.IsUsable == true ? MainWindow : null
        };
    }
}

/// <summary>The last normal (restored) main-window bounds; position in device pixels, size in device-independent units.</summary>
internal sealed record StudioWindowPlacement {
    public int X { get; init; }

    public int Y { get; init; }

    public double Width { get; init; }

    public double Height { get; init; }

    public bool Maximized { get; init; }

    internal bool IsUsable => double.IsFinite(Width) && double.IsFinite(Height) &&
        Width is >= 320 and <= 32768 && Height is >= 240 and <= 32768;
}
