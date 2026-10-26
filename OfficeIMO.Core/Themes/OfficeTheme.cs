using System;

namespace OfficeIMO;

/// <summary>An immutable visual identity shared by document formats. Loading a document does not apply a theme.</summary>
public sealed partial class OfficeTheme {
    /// <summary>Creates a theme from immutable palette, typography, table, chart and spacing components.</summary>
    public OfficeTheme(string name, OfficeThemePalette palette, OfficeThemeFonts fonts, OfficeThemeTypeScale typeScale,
        OfficeThemeTableStyle tableStyle, OfficeThemeChartStyle chartStyle, OfficeThemeSpacing spacing) {
        Name = ThemeValues.FontName(name, nameof(name));
        Palette = palette ?? throw new ArgumentNullException(nameof(palette));
        Fonts = fonts ?? throw new ArgumentNullException(nameof(fonts));
        TypeScale = typeScale ?? throw new ArgumentNullException(nameof(typeScale));
        TableStyle = tableStyle ?? throw new ArgumentNullException(nameof(tableStyle));
        ChartStyle = chartStyle ?? throw new ArgumentNullException(nameof(chartStyle));
        Spacing = spacing ?? throw new ArgumentNullException(nameof(spacing));
    }
    /// <summary>Gets the theme name.</summary>
    public string Name { get; }
    /// <summary>Gets literal palette colors.</summary>
    public OfficeThemePalette Palette { get; }
    /// <summary>Gets heading and body fonts.</summary>
    public OfficeThemeFonts Fonts { get; }
    /// <summary>Gets document type sizes.</summary>
    public OfficeThemeTypeScale TypeScale { get; }
    /// <summary>Gets table styling tokens.</summary>
    public OfficeThemeTableStyle TableStyle { get; }
    /// <summary>Gets chart styling tokens.</summary>
    public OfficeThemeChartStyle ChartStyle { get; }
    /// <summary>Gets paragraph, list and slide spacing.</summary>
    public OfficeThemeSpacing Spacing { get; }
    /// <summary>Gets the default visual identity, <see cref="Modern"/>.</summary>
    public static OfficeTheme Default => Modern;
    /// <summary>Returns a theme with a new name.</summary>
    public OfficeTheme WithName(string name) => Copy(name: name ?? throw new ArgumentNullException(nameof(name)));
    /// <summary>Returns a theme with a replacement palette.</summary>
    public OfficeTheme WithPalette(OfficeThemePalette palette) => Copy(palette: palette ?? throw new ArgumentNullException(nameof(palette)));
    /// <summary>Returns a theme with replacement heading and body fonts.</summary>
    public OfficeTheme WithFonts(OfficeThemeFonts fonts) => Copy(fonts: fonts ?? throw new ArgumentNullException(nameof(fonts)));
    /// <summary>Returns a theme with replacement type sizes.</summary>
    public OfficeTheme WithTypeScale(OfficeThemeTypeScale typeScale) => Copy(typeScale: typeScale ?? throw new ArgumentNullException(nameof(typeScale)));
    /// <summary>Returns a theme with replacement table styling.</summary>
    public OfficeTheme WithTableStyle(OfficeThemeTableStyle tableStyle) => Copy(tableStyle: tableStyle ?? throw new ArgumentNullException(nameof(tableStyle)));
    /// <summary>Returns a theme with replacement chart styling.</summary>
    public OfficeTheme WithChartStyle(OfficeThemeChartStyle chartStyle) => Copy(chartStyle: chartStyle ?? throw new ArgumentNullException(nameof(chartStyle)));
    /// <summary>Returns a theme with replacement spacing.</summary>
    public OfficeTheme WithSpacing(OfficeThemeSpacing spacing) => Copy(spacing: spacing ?? throw new ArgumentNullException(nameof(spacing)));
    private OfficeTheme Copy(string? name = null, OfficeThemePalette? palette = null, OfficeThemeFonts? fonts = null,
        OfficeThemeTypeScale? typeScale = null, OfficeThemeTableStyle? tableStyle = null,
        OfficeThemeChartStyle? chartStyle = null, OfficeThemeSpacing? spacing = null) =>
        new OfficeTheme(name ?? Name, palette ?? Palette, fonts ?? Fonts, typeScale ?? TypeScale,
            tableStyle ?? TableStyle, chartStyle ?? ChartStyle, spacing ?? Spacing);
}
