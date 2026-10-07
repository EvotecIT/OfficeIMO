namespace OfficeIMO.Html;

/// <summary>Identifies the standards surface described by a renderer capability.</summary>
public enum HtmlRenderCapabilityKind {
    /// <summary>One or more CSS properties or value families.</summary>
    Css = 0,
    /// <summary>An HTML element, attribute, or semantic behavior.</summary>
    Html = 1,
    /// <summary>A CSS at-rule or paged-media behavior.</summary>
    PagedMedia = 2,
    /// <summary>A resource, font, image, or SVG behavior.</summary>
    Resource = 3,
    /// <summary>An output-artifact behavior such as metadata or accessibility.</summary>
    Output = 4,
    /// <summary>Source bytes, character encodings, and decoding policy.</summary>
    Encoding = 5,
    /// <summary>Owned DOM nodes, queries, mutation, and snapshots.</summary>
    Dom = 6
}
