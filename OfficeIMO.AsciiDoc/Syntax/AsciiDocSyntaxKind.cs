namespace OfficeIMO.AsciiDoc;

/// <summary>Kinds produced by the lossless AsciiDoc syntax tree.</summary>
public enum AsciiDocSyntaxKind {
    /// <summary>Document root.</summary>
    Document = 0,
    /// <summary>Blank source line.</summary>
    BlankLine,
    /// <summary>Single-line comment.</summary>
    CommentLine,
    /// <summary>Delimited comment block.</summary>
    CommentBlock,
    /// <summary>Document title or section title.</summary>
    Heading,
    /// <summary>Heading marker token.</summary>
    HeadingMarker,
    /// <summary>Document attribute entry.</summary>
    AttributeEntry,
    /// <summary>Attribute punctuation marker.</summary>
    AttributeMarker,
    /// <summary>Attribute name token.</summary>
    AttributeName,
    /// <summary>Attribute value token.</summary>
    AttributeValue,
    /// <summary>Block element attribute list.</summary>
    BlockAttributeList,
    /// <summary>Content inside a block element attribute list.</summary>
    BlockAttributeListContent,
    /// <summary>Block title metadata line.</summary>
    BlockTitle,
    /// <summary>Block anchor metadata line.</summary>
    BlockAnchor,
    /// <summary>Paragraph block.</summary>
    Paragraph,
    /// <summary>Unordered list block.</summary>
    UnorderedList,
    /// <summary>Ordered list block.</summary>
    OrderedList,
    /// <summary>Description list block.</summary>
    DescriptionList,
    /// <summary>Description list item.</summary>
    DescriptionListItem,
    /// <summary>Description list marker.</summary>
    DescriptionListMarker,
    /// <summary>List continuation marker.</summary>
    ListContinuation,
    /// <summary>List item.</summary>
    ListItem,
    /// <summary>List marker token.</summary>
    ListMarker,
    /// <summary>Admonition paragraph.</summary>
    Admonition,
    /// <summary>Admonition label marker.</summary>
    AdmonitionMarker,
    /// <summary>Typed table block.</summary>
    Table,
    /// <summary>Typed table cell.</summary>
    TableCell,
    /// <summary>Table cell specifier.</summary>
    TableCellSpecifier,
    /// <summary>Table cell separator.</summary>
    TableCellSeparator,
    /// <summary>Table cell content.</summary>
    TableCellContent,
    /// <summary>Delimited block.</summary>
    DelimitedBlock,
    /// <summary>Opening or closing block delimiter.</summary>
    BlockDelimiter,
    /// <summary>Delimited block content.</summary>
    BlockContent,
    /// <summary>Block macro invocation.</summary>
    BlockMacro,
    /// <summary>Macro name token.</summary>
    MacroName,
    /// <summary>Macro separator token.</summary>
    MacroSeparator,
    /// <summary>Macro target token.</summary>
    MacroTarget,
    /// <summary>Macro attribute-list token.</summary>
    MacroAttributeList,
    /// <summary>Sequence of inline nodes inside block text.</summary>
    InlineSequence,
    /// <summary>Unformatted inline text.</summary>
    InlineText,
    /// <summary>Inline formatting span.</summary>
    InlineFormatted,
    /// <summary>Opening or closing inline formatting marker.</summary>
    InlineFormattingMarker,
    /// <summary>Inline document attribute reference.</summary>
    InlineAttributeReference,
    /// <summary>Inline macro invocation.</summary>
    InlineMacro,
    /// <summary>Inline cross-reference.</summary>
    InlineCrossReference,
    /// <summary>Inline anchor.</summary>
    InlineAnchor,
    /// <summary>Inline STEM expression.</summary>
    InlineStem,
    /// <summary>Inline passthrough.</summary>
    InlinePassthrough,
    /// <summary>Text token.</summary>
    Text,
    /// <summary>Line-ending token.</summary>
    LineEnding,
    /// <summary>Exact whitespace or punctuation trivia between recognized tokens.</summary>
    Trivia,
    /// <summary>Source-preserved construct without current profile semantics.</summary>
    Raw
}
