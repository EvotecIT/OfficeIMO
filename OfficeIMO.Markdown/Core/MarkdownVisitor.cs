namespace OfficeIMO.Markdown;


/// <summary>
/// Depth-first visitor over the OfficeIMO.Markdown object tree.
/// </summary>
public abstract class MarkdownVisitor {
    /// <summary>Visits a node when it is not <c>null</c>.</summary>
    public virtual void Visit(MarkdownObject? node) {
        if (node == null) {
            return;
        }

        switch (node) {
            case MarkdownDoc document:
                VisitDocument(document);
                break;

            case HeadingBlock heading:
                VisitHeadingBlock(heading);
                break;
            case ParagraphBlock paragraph:
                VisitParagraphBlock(paragraph);
                break;
            case QuoteBlock quote:
                VisitQuoteBlock(quote);
                break;
            case CalloutBlock callout:
                VisitCalloutBlock(callout);
                break;
            case DetailsBlock details:
                VisitDetailsBlock(details);
                break;
            case SummaryBlock summary:
                VisitSummaryBlock(summary);
                break;
            case OrderedListBlock orderedList:
                VisitOrderedListBlock(orderedList);
                break;
            case UnorderedListBlock unorderedList:
                VisitUnorderedListBlock(unorderedList);
                break;
            case TableBlock table:
                VisitTableBlock(table);
                break;
            case DefinitionListBlock definitionList:
                VisitDefinitionListBlock(definitionList);
                break;
            case FootnoteDefinitionBlock footnote:
                VisitFootnoteDefinitionBlock(footnote);
                break;
            case CodeBlock codeBlock:
                VisitCodeBlock(codeBlock);
                break;
            case SemanticFencedBlock semanticFencedBlock:
                VisitSemanticFencedBlock(semanticFencedBlock);
                break;
            case ImageBlock imageBlock:
                VisitImageBlock(imageBlock);
                break;
            case FrontMatterBlock frontMatter:
                VisitFrontMatterBlock(frontMatter);
                break;
            case HtmlCommentBlock htmlComment:
                VisitHtmlCommentBlock(htmlComment);
                break;
            case HtmlRawBlock htmlRawBlock:
                VisitHtmlRawBlock(htmlRawBlock);
                break;
            case HorizontalRuleBlock horizontalRule:
                VisitHorizontalRuleBlock(horizontalRule);
                break;
            case TocBlock toc:
                VisitTocBlock(toc);
                break;
            case TocMarkerBlock tocMarker:
                VisitTocMarkerBlock(tocMarker);
                break;
            case MarkdownBlock block:
                VisitBlock(block);
                break;

            case ListItem listItem:
                VisitListItem(listItem);
                break;
            case TableCell tableCell:
                VisitTableCell(tableCell);
                break;
            case DefinitionListGroup definitionGroup:
                VisitDefinitionListGroup(definitionGroup);
                break;
            case DefinitionListEntry definitionEntry:
                VisitDefinitionListEntry(definitionEntry);
                break;
            case DefinitionListDefinition definition:
                VisitDefinitionListDefinition(definition);
                break;

            case InlineSequence inlineSequence:
                VisitInlineSequence(inlineSequence);
                break;
            case LinkInline link:
                VisitLinkInline(link);
                break;
            case ImageLinkInline imageLink:
                VisitImageLinkInline(imageLink);
                break;
            case ImageInline imageInline:
                VisitImageInline(imageInline);
                break;
            case AbbreviationInline abbreviation:
                VisitAbbreviationInline(abbreviation);
                break;
            case HtmlTagSequenceInline htmlTag:
                VisitHtmlTagSequenceInline(htmlTag);
                break;
            case BoldSequenceInline boldSequence:
                VisitBoldSequenceInline(boldSequence);
                break;
            case ItalicSequenceInline italicSequence:
                VisitItalicSequenceInline(italicSequence);
                break;
            case BoldItalicSequenceInline boldItalicSequence:
                VisitBoldItalicSequenceInline(boldItalicSequence);
                break;
            case StrikethroughSequenceInline strikethroughSequence:
                VisitStrikethroughSequenceInline(strikethroughSequence);
                break;
            case HighlightSequenceInline highlightSequence:
                VisitHighlightSequenceInline(highlightSequence);
                break;
            case InsertedSequenceInline insertedSequence:
                VisitInsertedSequenceInline(insertedSequence);
                break;
            case SuperscriptSequenceInline superscriptSequence:
                VisitSuperscriptSequenceInline(superscriptSequence);
                break;
            case SubscriptSequenceInline subscriptSequence:
                VisitSubscriptSequenceInline(subscriptSequence);
                break;
            case MarkdownTextRun textRun:
                VisitTextRun(textRun);
                break;
            case CodeSpanInline codeSpan:
                VisitCodeSpanInline(codeSpan);
                break;
            case FootnoteRefInline footnoteRef:
                VisitFootnoteRefInline(footnoteRef);
                break;
            case HardBreakInline hardBreak:
                VisitHardBreakInline(hardBreak);
                break;
            case SoftBreakInline softBreak:
                VisitSoftBreakInline(softBreak);
                break;
            case BoldInline bold:
                VisitBoldInline(bold);
                break;
            case ItalicInline italic:
                VisitItalicInline(italic);
                break;
            case BoldItalicInline boldItalic:
                VisitBoldItalicInline(boldItalic);
                break;
            case StrikethroughInline strikethrough:
                VisitStrikethroughInline(strikethrough);
                break;
            case HighlightInline highlight:
                VisitHighlightInline(highlight);
                break;
            case InsertedInline inserted:
                VisitInsertedInline(inserted);
                break;
            case SuperscriptInline superscript:
                VisitSuperscriptInline(superscript);
                break;
            case SubscriptInline subscript:
                VisitSubscriptInline(subscript);
                break;
            case UnderlineInline underline:
                VisitUnderlineInline(underline);
                break;
            case HtmlRawInline htmlRawInline:
                VisitHtmlRawInline(htmlRawInline);
                break;
            case MarkdownInline inline:
                VisitInline(inline);
                break;

            default:
                DefaultVisit(node);
                break;
        }
    }

    /// <summary>Visits a sequence of nodes in order.</summary>
    public virtual void Visit(IEnumerable<MarkdownObject>? nodes) {
        if (nodes == null) {
            return;
        }

        foreach (var node in nodes) {
            Visit(node);
        }
    }

    /// <summary>Visits all direct children of the node in order.</summary>
    protected void VisitChildren(MarkdownObject node) {
        var children = node.ChildObjects;
        for (int i = 0; i < children.Count; i++) {
            Visit(children[i]);
        }
    }

    /// <summary>Default visit behavior for nodes without a more specific override.</summary>
    protected virtual void DefaultVisit(MarkdownObject node) => VisitChildren(node);

    /// <summary>Visits a document and, by default, its children.</summary>
    protected virtual void VisitDocument(MarkdownDoc document) => DefaultVisit(document);

    /// <summary>Visits a generic block and, by default, its children.</summary>
    protected virtual void VisitBlock(MarkdownBlock block) => DefaultVisit(block);
    /// <summary>Visits a heading block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitHeadingBlock(HeadingBlock block) => VisitBlock(block);
    /// <summary>Visits a paragraph block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitParagraphBlock(ParagraphBlock block) => VisitBlock(block);
    /// <summary>Visits a quote block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitQuoteBlock(QuoteBlock block) => VisitBlock(block);
    /// <summary>Visits a callout block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitCalloutBlock(CalloutBlock block) => VisitBlock(block);
    /// <summary>Visits a details block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitDetailsBlock(DetailsBlock block) => VisitBlock(block);
    /// <summary>Visits a summary block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitSummaryBlock(SummaryBlock block) => VisitBlock(block);
    /// <summary>Visits an ordered-list block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitOrderedListBlock(OrderedListBlock block) => VisitBlock(block);
    /// <summary>Visits an unordered-list block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitUnorderedListBlock(UnorderedListBlock block) => VisitBlock(block);
    /// <summary>Visits a table block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitTableBlock(TableBlock block) => VisitBlock(block);
    /// <summary>Visits a definition-list block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitDefinitionListBlock(DefinitionListBlock block) => VisitBlock(block);
    /// <summary>Visits a footnote definition through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitFootnoteDefinitionBlock(FootnoteDefinitionBlock block) => VisitBlock(block);
    /// <summary>Visits a code block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitCodeBlock(CodeBlock block) => VisitBlock(block);
    /// <summary>Visits a semantic fenced block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitSemanticFencedBlock(SemanticFencedBlock block) => VisitBlock(block);
    /// <summary>Visits an image block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitImageBlock(ImageBlock block) => VisitBlock(block);
    /// <summary>Visits a front-matter block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitFrontMatterBlock(FrontMatterBlock block) => VisitBlock(block);
    /// <summary>Visits an HTML comment block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitHtmlCommentBlock(HtmlCommentBlock block) => VisitBlock(block);
    /// <summary>Visits a raw HTML block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitHtmlRawBlock(HtmlRawBlock block) => VisitBlock(block);
    /// <summary>Visits a horizontal-rule block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitHorizontalRuleBlock(HorizontalRuleBlock block) => VisitBlock(block);
    /// <summary>Visits a generated table-of-contents block through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitTocBlock(TocBlock block) => VisitBlock(block);
    /// <summary>Visits a table-of-contents marker through <see cref="VisitBlock"/>.</summary>
    protected virtual void VisitTocMarkerBlock(TocMarkerBlock block) => VisitBlock(block);

    /// <summary>Visits a list item and, by default, its children.</summary>
    protected virtual void VisitListItem(ListItem item) => DefaultVisit(item);
    /// <summary>Visits a table cell and, by default, its children.</summary>
    protected virtual void VisitTableCell(TableCell cell) => DefaultVisit(cell);
    /// <summary>Visits a definition-list group and, by default, its children.</summary>
    protected virtual void VisitDefinitionListGroup(DefinitionListGroup group) => DefaultVisit(group);
    /// <summary>Visits a definition-list entry and, by default, its children.</summary>
    protected virtual void VisitDefinitionListEntry(DefinitionListEntry entry) => DefaultVisit(entry);
    /// <summary>Visits a definition-list definition and, by default, its children.</summary>
    protected virtual void VisitDefinitionListDefinition(DefinitionListDefinition definition) => DefaultVisit(definition);

    /// <summary>Visits a generic inline and, by default, its children.</summary>
    protected virtual void VisitInline(MarkdownInline inline) => DefaultVisit(inline);
    /// <summary>Visits an inline sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitInlineSequence(InlineSequence sequence) => VisitInline(sequence);
    /// <summary>Visits a link inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitLinkInline(LinkInline inline) => VisitInline(inline);
    /// <summary>Visits an image-link inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitImageLinkInline(ImageLinkInline inline) => VisitInline(inline);
    /// <summary>Visits an image inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitImageInline(ImageInline inline) => VisitInline(inline);
    /// <summary>Visits an abbreviation inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitAbbreviationInline(AbbreviationInline inline) => VisitInline(inline);
    /// <summary>Visits an HTML-tag sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitHtmlTagSequenceInline(HtmlTagSequenceInline inline) => VisitInline(inline);
    /// <summary>Visits a bold sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitBoldSequenceInline(BoldSequenceInline inline) => VisitInline(inline);
    /// <summary>Visits an italic sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitItalicSequenceInline(ItalicSequenceInline inline) => VisitInline(inline);
    /// <summary>Visits a bold-italic sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitBoldItalicSequenceInline(BoldItalicSequenceInline inline) => VisitInline(inline);
    /// <summary>Visits a strikethrough sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitStrikethroughSequenceInline(StrikethroughSequenceInline inline) => VisitInline(inline);
    /// <summary>Visits a highlighted sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitHighlightSequenceInline(HighlightSequenceInline inline) => VisitInline(inline);
    /// <summary>Visits an inserted-text sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitInsertedSequenceInline(InsertedSequenceInline inline) => VisitInline(inline);
    /// <summary>Visits a superscript sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitSuperscriptSequenceInline(SuperscriptSequenceInline inline) => VisitInline(inline);
    /// <summary>Visits a subscript sequence through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitSubscriptSequenceInline(SubscriptSequenceInline inline) => VisitInline(inline);
    /// <summary>Visits a text run through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitTextRun(MarkdownTextRun inline) => VisitInline(inline);
    /// <summary>Visits a code-span inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitCodeSpanInline(CodeSpanInline inline) => VisitInline(inline);
    /// <summary>Visits a footnote reference through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitFootnoteRefInline(FootnoteRefInline inline) => VisitInline(inline);
    /// <summary>Visits a hard line break through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitHardBreakInline(HardBreakInline inline) => VisitInline(inline);
    /// <summary>Visits a soft line break through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitSoftBreakInline(SoftBreakInline inline) => VisitInline(inline);
    /// <summary>Visits a bold inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitBoldInline(BoldInline inline) => VisitInline(inline);
    /// <summary>Visits an italic inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitItalicInline(ItalicInline inline) => VisitInline(inline);
    /// <summary>Visits a bold-italic inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitBoldItalicInline(BoldItalicInline inline) => VisitInline(inline);
    /// <summary>Visits a strikethrough inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitStrikethroughInline(StrikethroughInline inline) => VisitInline(inline);
    /// <summary>Visits a highlighted inline through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitHighlightInline(HighlightInline inline) => VisitInline(inline);
    /// <summary>Visits inserted text through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitInsertedInline(InsertedInline inline) => VisitInline(inline);
    /// <summary>Visits superscript text through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitSuperscriptInline(SuperscriptInline inline) => VisitInline(inline);
    /// <summary>Visits subscript text through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitSubscriptInline(SubscriptInline inline) => VisitInline(inline);
    /// <summary>Visits underlined text through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitUnderlineInline(UnderlineInline inline) => VisitInline(inline);
    /// <summary>Visits raw HTML inline content through <see cref="VisitInline"/>.</summary>
    protected virtual void VisitHtmlRawInline(HtmlRawInline inline) => VisitInline(inline);
}
