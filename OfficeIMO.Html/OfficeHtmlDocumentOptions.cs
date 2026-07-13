namespace OfficeIMO.Html;

/// <summary>
/// Options used by shared OfficeIMO HTML document shell helpers.
/// </summary>
public sealed class OfficeHtmlDocumentOptions {
    /// <summary>HTML document title.</summary>
    public string Title { get; set; } = "OfficeIMO HTML";

    /// <summary>Theme applied when default shell styles are included.</summary>
    public OfficeVisualThemeKind Theme { get; set; } = OfficeVisualThemeKind.WordLike;

    /// <summary>When true, emits the shared OfficeIMO CSS shell.</summary>
    public bool IncludeDefaultStyles { get; set; } = true;

    /// <summary>Optional CSS class assigned to the generated body.</summary>
    public string BodyClass { get; set; } = "officeimo-html";

    /// <summary>Line ending used by generated HTML.</summary>
    public string NewLine { get; set; } = "\n";
}
