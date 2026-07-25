namespace OfficeIMO.Html;

/// <summary>Identifies the standards surface described by a renderer capability.</summary>
public enum HtmlRenderCapabilityKind {
    /// <summary>One or more CSS properties or value families.</summary>
    Css,
    /// <summary>An HTML element, attribute, or semantic behavior.</summary>
    Html,
    /// <summary>A CSS at-rule or paged-media behavior.</summary>
    PagedMedia,
    /// <summary>A resource, font, image, or SVG behavior.</summary>
    Resource,
    /// <summary>An output-artifact behavior such as metadata or accessibility.</summary>
    Output
}
