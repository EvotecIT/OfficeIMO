namespace OfficeIMO.Markdown;

/// <summary>
/// Syntax node kinds produced by <see cref="MarkdownReader.ParseWithSyntaxTree(string, MarkdownReaderOptions?)"/>.
/// </summary>
public enum MarkdownSyntaxKind {
    /// <summary>Root document node.</summary>
    Document,
    /// <summary>ATX or Setext heading block.</summary>
    Heading,
    /// <summary>Heading level metadata.</summary>
    HeadingLevel,
    /// <summary>Heading text payload.</summary>
    HeadingText,
    /// <summary>Plain text inline node.</summary>
    InlineText,
    /// <summary>Inline code span node.</summary>
    InlineCodeSpan,
    /// <summary>Hyperlink inline node.</summary>
    InlineLink,
    /// <summary>Hyperlink destination URL metadata.</summary>
    InlineLinkTarget,
    /// <summary>Optional hyperlink title metadata.</summary>
    InlineLinkTitle,
    /// <summary>Optional hyperlink target attribute preserved from richer HTML sources.</summary>
    InlineLinkHtmlTarget,
    /// <summary>Optional hyperlink rel attribute preserved from richer HTML sources.</summary>
    InlineLinkHtmlRel,
    /// <summary>Standalone inline image node.</summary>
    InlineImage,
    /// <summary>Linked inline image node.</summary>
    InlineImageLink,
    /// <summary>Strong/bold inline node.</summary>
    InlineStrong,
    /// <summary>Emphasis/italic inline node.</summary>
    InlineEmphasis,
    /// <summary>Combined strong+emphasis inline node.</summary>
    InlineStrongEmphasis,
    /// <summary>Strikethrough inline node.</summary>
    InlineStrikethrough,
    /// <summary>Highlight/mark inline node.</summary>
    InlineHighlight,
    /// <summary>Underline inline node.</summary>
    InlineUnderline,
    /// <summary>Hard line break inline node.</summary>
    InlineHardBreak,
    /// <summary>Inline HTML tag wrapper node.</summary>
    InlineHtmlTag,
    /// <summary>Raw inline HTML node.</summary>
    InlineHtmlRaw,
    /// <summary>Footnote reference inline node.</summary>
    InlineFootnoteRef,
    /// <summary>Paragraph block.</summary>
    Paragraph,
    /// <summary>Blockquote block.</summary>
    Quote,
    /// <summary>Unordered list block.</summary>
    UnorderedList,
    /// <summary>Ordered list block.</summary>
    OrderedList,
    /// <summary>List item node.</summary>
    ListItem,
    /// <summary>Fenced or indented code block.</summary>
    CodeBlock,
    /// <summary>Host-defined semantic fenced block.</summary>
    SemanticFencedBlock,
    /// <summary>Semantic kind metadata for a semantic fenced block.</summary>
    FenceSemanticKind,
    /// <summary>Fenced code block info string / language hint.</summary>
    CodeFenceInfo,
    /// <summary>Code block content payload.</summary>
    CodeContent,
    /// <summary>Markdown table block.</summary>
    Table,
    /// <summary>Header row inside a markdown table.</summary>
    TableHeader,
    /// <summary>Body row inside a markdown table.</summary>
    TableRow,
    /// <summary>Single cell inside a markdown table row/header.</summary>
    TableCell,
    /// <summary>Horizontal rule block.</summary>
    HorizontalRule,
    /// <summary>Image block.</summary>
    Image,
    /// <summary>Image alternative text.</summary>
    ImageAlt,
    /// <summary>Image source path or URL.</summary>
    ImageSource,
    /// <summary>Optional hyperlink target wrapping an image block.</summary>
    ImageLinkTarget,
    /// <summary>Optional hyperlink title wrapping an image block.</summary>
    ImageLinkTitle,
    /// <summary>Optional hyperlink target wrapping an image block.</summary>
    ImageLinkHtmlTarget,
    /// <summary>Optional hyperlink rel wrapping an image block.</summary>
    ImageLinkHtmlRel,
    /// <summary>Image title attribute.</summary>
    ImageTitle,
    /// <summary>Callout or admonition block.</summary>
    Callout,
    /// <summary>Callout/admonition marker kind such as note or tip.</summary>
    CalloutKind,
    /// <summary>Inline title/header content for a callout block.</summary>
    CalloutTitle,
    /// <summary>Definition list block.</summary>
    DefinitionList,
    /// <summary>Semantic definition-list group with shared terms and definitions.</summary>
    DefinitionGroup,
    /// <summary>Single definition list item.</summary>
    DefinitionItem,
    /// <summary>Definition list term node.</summary>
    DefinitionTerm,
    /// <summary>Definition list definition/content node.</summary>
    DefinitionValue,
    /// <summary>Footnote definition block.</summary>
    FootnoteDefinition,
    /// <summary>Footnote definition label/identifier.</summary>
    FootnoteLabel,
    /// <summary>Reference-style link definition consumed during parsing.</summary>
    ReferenceLinkDefinition,
    /// <summary>Reference-style link definition label/identifier.</summary>
    ReferenceLinkLabel,
    /// <summary>Reference-style link definition URL/destination.</summary>
    ReferenceLinkUrl,
    /// <summary>Reference-style link definition optional title.</summary>
    ReferenceLinkTitle,
    /// <summary>Details/disclosure block.</summary>
    Details,
    /// <summary>Summary node inside a details block.</summary>
    Summary,
    /// <summary>Front matter block.</summary>
    FrontMatter,
    /// <summary>Raw HTML block.</summary>
    HtmlRaw,
    /// <summary>HTML comment block.</summary>
    HtmlComment,
    /// <summary>Generated table of contents block.</summary>
    Toc,
    /// <summary>Placeholder table of contents block.</summary>
    TocPlaceholder,
    /// <summary>Fallback for blocks without a dedicated mapping yet.</summary>
    Unknown
}
