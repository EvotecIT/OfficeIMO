using System;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>An immutable visual identity shared by document formats. Loading a document does not apply a theme.</summary>
public sealed partial class OfficeTheme {
    /// <summary>
    /// Creates a theme from immutable palette, typography, table, chart and spacing components. Text style, surfaces,
    /// page margins and footer take the <see cref="Modern"/> treatment, with tints derived from <paramref name="palette"/>.
    /// </summary>
    public OfficeTheme(string name, OfficeThemePalette palette, OfficeThemeFonts fonts, OfficeThemeTypeScale typeScale,
        OfficeThemeTableStyle tableStyle, OfficeThemeChartStyle chartStyle, OfficeThemeSpacing spacing)
        : this(name, palette, fonts, typeScale, tableStyle, chartStyle, spacing, DefaultTextStyle(),
            OfficeThemeSurfaces.FromPalette(palette ?? throw new ArgumentNullException(nameof(palette)),
                (tableStyle ?? throw new ArgumentNullException(nameof(tableStyle))).BandFill, tableStyle.BorderColor),
            DefaultPage(), DefaultFooter()) {
    }

    /// <summary>Creates a theme from every immutable component.</summary>
    public OfficeTheme(string name, OfficeThemePalette palette, OfficeThemeFonts fonts, OfficeThemeTypeScale typeScale,
        OfficeThemeTableStyle tableStyle, OfficeThemeChartStyle chartStyle, OfficeThemeSpacing spacing,
        OfficeThemeTextStyle textStyle, OfficeThemeSurfaces surfaces, OfficeThemePage page, OfficeThemeFooter footer) {
        Name = ThemeValues.FontName(name, nameof(name));
        Palette = palette ?? throw new ArgumentNullException(nameof(palette));
        Fonts = fonts ?? throw new ArgumentNullException(nameof(fonts));
        TypeScale = typeScale ?? throw new ArgumentNullException(nameof(typeScale));
        TableStyle = tableStyle ?? throw new ArgumentNullException(nameof(tableStyle));
        ChartStyle = chartStyle ?? throw new ArgumentNullException(nameof(chartStyle));
        Spacing = spacing ?? throw new ArgumentNullException(nameof(spacing));
        TextStyle = textStyle ?? throw new ArgumentNullException(nameof(textStyle));
        Surfaces = surfaces ?? throw new ArgumentNullException(nameof(surfaces));
        Page = page ?? throw new ArgumentNullException(nameof(page));
        Footer = footer ?? throw new ArgumentNullException(nameof(footer));
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
    /// <summary>Gets line spacing, title, heading and link colours and heading spacing.</summary>
    public OfficeThemeTextStyle TextStyle { get; }
    /// <summary>Gets panel, hairline, accent-tint and status fills.</summary>
    public OfficeThemeSurfaces Surfaces { get; }
    /// <summary>Gets page margins for Word and PDF output.</summary>
    public OfficeThemePage Page { get; }
    /// <summary>Gets footer text styling.</summary>
    public OfficeThemeFooter Footer { get; }
    /// <summary>Gets whether the theme paints a dark background (used by screen formats; Word and Excel keep a light page).</summary>
    public bool IsDark => OfficeThemeSurfaces.IsDark(Palette.Background);
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
    /// <summary>Returns a theme with replacement text style.</summary>
    public OfficeTheme WithTextStyle(OfficeThemeTextStyle textStyle) => Copy(textStyle: textStyle ?? throw new ArgumentNullException(nameof(textStyle)));
    /// <summary>Returns a theme with replacement surface colours.</summary>
    public OfficeTheme WithSurfaces(OfficeThemeSurfaces surfaces) => Copy(surfaces: surfaces ?? throw new ArgumentNullException(nameof(surfaces)));
    /// <summary>Returns a theme with replacement page margins.</summary>
    public OfficeTheme WithPage(OfficeThemePage page) => Copy(page: page ?? throw new ArgumentNullException(nameof(page)));
    /// <summary>Returns a theme with replacement footer styling.</summary>
    public OfficeTheme WithFooter(OfficeThemeFooter footer) => Copy(footer: footer ?? throw new ArgumentNullException(nameof(footer)));
    private OfficeTheme Copy(string? name = null, OfficeThemePalette? palette = null, OfficeThemeFonts? fonts = null,
        OfficeThemeTypeScale? typeScale = null, OfficeThemeTableStyle? tableStyle = null,
        OfficeThemeChartStyle? chartStyle = null, OfficeThemeSpacing? spacing = null, OfficeThemeTextStyle? textStyle = null,
        OfficeThemeSurfaces? surfaces = null, OfficeThemePage? page = null, OfficeThemeFooter? footer = null) =>
        new OfficeTheme(name ?? Name, palette ?? Palette, fonts ?? Fonts, typeScale ?? TypeScale,
            tableStyle ?? TableStyle, chartStyle ?? ChartStyle, spacing ?? Spacing, textStyle ?? TextStyle,
            surfaces ?? Surfaces, page ?? Page, footer ?? Footer);

    // The seven-component constructor keeps the default theme's text, page and footer treatment.
    private static OfficeThemeTextStyle DefaultTextStyle() => Modern.TextStyle;
    private static OfficeThemePage DefaultPage() => Modern.Page;
    private static OfficeThemeFooter DefaultFooter() => Modern.Footer;
}
