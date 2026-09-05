using Avalonia.Styling;

namespace OfficeIMO.Studio.Infrastructure.Preferences;

/// <summary>Application theme variants beyond the platform-provided light and dark variants.</summary>
public static class StudioThemeVariants {
    /// <summary>A dark, high-contrast palette for users who need stronger visual separation.</summary>
    public static ThemeVariant HighContrast { get; } = new("HighContrast", ThemeVariant.Dark);
}
