namespace OfficeIMO.Markdown;

/// <summary>
/// Describes a built-in visual theme preset that can be offered by callers as a stable theme choice.
/// </summary>
public sealed class MarkdownVisualThemePreset {
    internal MarkdownVisualThemePreset(
        OfficeVisualThemeKind kind,
        string name,
        string description,
        HtmlStyle htmlStyle,
        params string[] aliases) {
        Kind = kind;
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Description = description ?? throw new ArgumentNullException(nameof(description));
        HtmlStyle = htmlStyle;
        Aliases = Array.AsReadOnly(aliases ?? Array.Empty<string>());
    }

    /// <summary>Built-in theme kind used to create the preset.</summary>
    public OfficeVisualThemeKind Kind { get; }

    /// <summary>Stable display/API name for the preset.</summary>
    public string Name { get; }

    /// <summary>Short description callers can show in pickers or documentation.</summary>
    public string Description { get; }

    /// <summary>Preferred HTML style used when rendering the preset to HTML.</summary>
    public HtmlStyle HtmlStyle { get; }

    /// <summary>Additional accepted names for lookup and front matter.</summary>
    public IReadOnlyList<string> Aliases { get; }

    /// <summary>Creates a fresh theme instance for this preset.</summary>
    public MarkdownVisualTheme CreateTheme() => MarkdownVisualTheme.Create(Kind);

    /// <summary>Creates a fresh theme instance for this preset and applies a built-in color scheme.</summary>
    public MarkdownVisualTheme CreateTheme(MarkdownColorSchemeKind colorScheme) => MarkdownVisualTheme.Create(Kind, colorScheme);
}
