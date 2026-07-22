namespace OfficeIMO.Html;

/// <summary>Document capabilities used by shared target contracts and preflight analysis.</summary>
public enum HtmlSemanticFeature {
    /// <summary>Document-level title, language, author, and related metadata.</summary>
    Metadata,
    /// <summary>Document sections, articles, pages, sheets, or slides.</summary>
    Sections,
    /// <summary>Heading hierarchy.</summary>
    Headings,
    /// <summary>Paragraph and block text.</summary>
    Paragraphs,
    /// <summary>Inline emphasis, decoration, vertical position, and related run styling.</summary>
    RichText,
    /// <summary>Hyperlinks and named destinations.</summary>
    Links,
    /// <summary>Ordered, unordered, and nested lists.</summary>
    Lists,
    /// <summary>Tabular structure, spans, headers, and captions.</summary>
    Tables,
    /// <summary>Raster or vector images.</summary>
    Images,
    /// <summary>Audio, video, object, and other media elements.</summary>
    Media,
    /// <summary>HTML forms and controls.</summary>
    Forms,
    /// <summary>Footnotes, endnotes, presenter notes, and note-like content.</summary>
    Notes,
    /// <summary>Review comments and discussion annotations.</summary>
    Comments,
    /// <summary>Bookmarks, tracked changes, labels, and other semantic annotations.</summary>
    Annotations,
    /// <summary>Spreadsheet formulas and calculated expressions.</summary>
    Formulas,
    /// <summary>Native or visual charts.</summary>
    Charts,
    /// <summary>Authored positions, sizes, transforms, and drawing order.</summary>
    Geometry,
    /// <summary>CSS cascade and computed presentation.</summary>
    Css,
    /// <summary>External and embedded resource resolution.</summary>
    Resources,
    /// <summary>Paged layout, fragmentation, page boxes, and running content.</summary>
    PagedLayout
}
