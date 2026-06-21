namespace OfficeIMO.Html;

/// <summary>
/// Normalized logical HTML node categories used by OfficeIMO conversion tooling.
/// </summary>
public enum HtmlLogicalNodeKind {
    /// <summary>Root document node.</summary>
    Document,
    /// <summary>Sectioning or body-level container.</summary>
    Section,
    /// <summary>Heading element.</summary>
    Heading,
    /// <summary>Paragraph-like block.</summary>
    Paragraph,
    /// <summary>Ordered or unordered list.</summary>
    List,
    /// <summary>List item.</summary>
    ListItem,
    /// <summary>Table element.</summary>
    Table,
    /// <summary>Table row.</summary>
    TableRow,
    /// <summary>Table header or data cell.</summary>
    TableCell,
    /// <summary>Figure element.</summary>
    Figure,
    /// <summary>Image, picture, or inline SVG.</summary>
    Image,
    /// <summary>Audio, video, source, or track media element.</summary>
    Media,
    /// <summary>Hyperlink element.</summary>
    Link,
    /// <summary>Form container.</summary>
    Form,
    /// <summary>Input, select, textarea, button, or option control.</summary>
    FormControl,
    /// <summary>Text node or text-like leaf element.</summary>
    Text,
    /// <summary>Inline semantic or formatting element.</summary>
    Inline,
    /// <summary>Document metadata element.</summary>
    Metadata,
    /// <summary>Node that does not map to a known logical kind.</summary>
    Unknown,
    /// <summary>Table caption.</summary>
    TableCaption
}
