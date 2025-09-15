namespace OfficeIMO.Markdown;

/// <summary>
/// Optional color overrides for HTML rendering. Values are CSS color strings (e.g., #RRGGBB).
/// Leave null to keep defaults from the selected HtmlStyle.
/// </summary>
public sealed class ThemeColors {
    /// <summary>
    /// Accent color for light mode. Used for links, active TOC items, and accented borders.
    /// Example: "#0969da".
    /// </summary>
    public string? AccentLight { get; set; } = null;
    /// <summary>
    /// Accent color for dark mode. Used for links, active TOC items, and accented borders.
    /// Example: "#2f81f7".
    /// </summary>
    public string? AccentDark { get; set; } = null;
    /// <summary>
    /// Heading color (h1..h6) in light mode.
    /// </summary>
    public string? HeadingLight { get; set; } = null;
    /// <summary>
    /// Heading color (h1..h6) in dark mode.
    /// </summary>
    public string? HeadingDark { get; set; } = null;
    /// <summary>
    /// Background color of the TOC panel in light mode.
    /// </summary>
    public string? TocBgLight { get; set; } = null;
    /// <summary>
    /// Background color of the TOC panel in dark mode.
    /// </summary>
    public string? TocBgDark { get; set; } = null;
    /// <summary>
    /// Border color of the TOC in light mode.
    /// </summary>
    public string? TocBorderLight { get; set; } = null;
    /// <summary>
    /// Border color of the TOC in dark mode.
    /// </summary>
    public string? TocBorderDark { get; set; } = null;
    /// <summary>
    /// Active TOC link color in light mode (overrides <see cref="AccentLight"/> when set).
    /// </summary>
    public string? ActiveLinkLight { get; set; } = null;
    /// <summary>
    /// Active TOC link color in dark mode (overrides <see cref="AccentDark"/> when set).
    /// </summary>
    public string? ActiveLinkDark { get; set; } = null;
}
