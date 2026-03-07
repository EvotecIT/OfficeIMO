namespace OfficeIMO.Markdown;

/// <summary>
/// Syntax node kinds produced by <see cref="MarkdownReader.ParseWithSyntaxTree(string, MarkdownReaderOptions?)"/>.
/// </summary>
public enum MarkdownSyntaxKind {
    /// <summary>Root document node.</summary>
    Document,
    /// <summary>ATX or Setext heading block.</summary>
    Heading,
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
    /// <summary>Horizontal rule block.</summary>
    HorizontalRule,
    /// <summary>Image block.</summary>
    Image,
    /// <summary>Image alternative text.</summary>
    ImageAlt,
    /// <summary>Image source path or URL.</summary>
    ImageSource,
    /// <summary>Image title attribute.</summary>
    ImageTitle,
    /// <summary>Callout or admonition block.</summary>
    Callout,
    /// <summary>Definition list block.</summary>
    DefinitionList,
    /// <summary>Single definition list item.</summary>
    DefinitionItem,
    /// <summary>Definition list term node.</summary>
    DefinitionTerm,
    /// <summary>Definition list definition/content node.</summary>
    DefinitionValue,
    /// <summary>Footnote definition block.</summary>
    FootnoteDefinition,
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
