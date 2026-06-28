namespace OfficeIMO.Markdown;

/// <summary>
/// Native inline projection kind for UI/read-model consumers.
/// </summary>
public enum MarkdownNativeInlineKind {
    /// <summary>Plain text content.</summary>
    Text,
    /// <summary>Inline code span.</summary>
    Code,
    /// <summary>Hyperlink inline.</summary>
    Link,
    /// <summary>Standalone inline image.</summary>
    Image,
    /// <summary>Linked inline image.</summary>
    ImageLink,
    /// <summary>Strong/bold inline content.</summary>
    Strong,
    /// <summary>Emphasis/italic inline content.</summary>
    Emphasis,
    /// <summary>Combined strong and emphasis inline content.</summary>
    StrongEmphasis,
    /// <summary>Strikethrough inline content.</summary>
    Strikethrough,
    /// <summary>Highlighted inline content.</summary>
    Highlight,
    /// <summary>Inserted inline content.</summary>
    Inserted,
    /// <summary>Underline inline content.</summary>
    Underline,
    /// <summary>Hard line break.</summary>
    HardBreak,
    /// <summary>Inline HTML tag wrapper.</summary>
    HtmlTag,
    /// <summary>Raw inline HTML.</summary>
    HtmlRaw,
    /// <summary>Footnote reference.</summary>
    FootnoteRef,
    /// <summary>Inline node without a specialized native projection.</summary>
    Other
}
