using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Options for saving a PowerPoint-to-HTML capability gallery scenario.
/// </summary>
public sealed class PowerPointHtmlCapabilityGalleryOptions {
    /// <summary>Stable scenario identifier used for file names and manifests.</summary>
    public string ScenarioId { get; set; } = "powerpoint-rich-presentation";

    /// <summary>Human-readable scenario title.</summary>
    public string Title { get; set; } = "PowerPoint Rich Presentation";

    /// <summary>Theme applied to generated HTML artifacts.</summary>
    public OfficeHtmlDocumentThemeKind Theme { get; set; } = OfficeHtmlDocumentThemeKind.Report;

    /// <summary>When true, hidden slides are included.</summary>
    public bool IncludeHiddenSlides { get; set; }

    /// <summary>When true, slide notes are included in extraction proof.</summary>
    public bool IncludeNotes { get; set; } = true;

    /// <summary>When true, table shapes are exported.</summary>
    public bool IncludeTables { get; set; } = true;
}
