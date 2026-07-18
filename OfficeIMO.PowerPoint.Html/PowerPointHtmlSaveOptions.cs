using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Options for exporting PowerPoint presentations to HTML.
/// </summary>
public sealed class PowerPointHtmlSaveOptions {
    /// <summary>PowerPoint-to-HTML lane to export. Defaults to semantic slides.</summary>
    public OfficeHtmlConversionProfile Profile { get; set; } = OfficeHtmlConversionProfile.PowerPointSemanticSlides;

    /// <summary>Theme used by the shared OfficeIMO HTML shell.</summary>
    public OfficeVisualThemeKind Theme { get; set; } = OfficeVisualThemeKind.WordLike;

    /// <summary>When true, emits the shared OfficeIMO CSS shell.</summary>
    public bool IncludeDefaultStyles { get; set; } = true;

    /// <summary>Optional document title.</summary>
    public string? Title { get; set; }

    /// <summary>When true, hidden slides are included.</summary>
    public bool IncludeHiddenSlides { get; set; }

    /// <summary>When true, notes are included in the extraction proof block.</summary>
    public bool IncludeNotes { get; set; } = true;

    /// <summary>When true, tables are exported.</summary>
    public bool IncludeTables { get; set; } = true;

    /// <summary>When true, hidden shapes are included in semantic and positioned review output.</summary>
    public bool IncludeHiddenShapes { get; set; }

    /// <summary>When true, emits slide-aligned extraction markdown as proof text.</summary>
    public bool IncludeExtractionProof { get; set; } = true;

    /// <summary>Creates a reusable copy of these conversion settings.</summary>
    public PowerPointHtmlSaveOptions Clone() => new PowerPointHtmlSaveOptions {
        Profile = Profile,
        Theme = Theme,
        IncludeDefaultStyles = IncludeDefaultStyles,
        Title = Title,
        IncludeHiddenSlides = IncludeHiddenSlides,
        IncludeNotes = IncludeNotes,
        IncludeTables = IncludeTables,
        IncludeHiddenShapes = IncludeHiddenShapes,
        IncludeExtractionProof = IncludeExtractionProof
    };
}
