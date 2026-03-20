namespace OfficeIMO.Markdown;

#pragma warning disable CS1591

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
            case TextRun textRun:
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

    protected virtual void VisitDocument(MarkdownDoc document) => DefaultVisit(document);

    protected virtual void VisitBlock(MarkdownBlock block) => DefaultVisit(block);
    protected virtual void VisitHeadingBlock(HeadingBlock block) => VisitBlock(block);
    protected virtual void VisitParagraphBlock(ParagraphBlock block) => VisitBlock(block);
    protected virtual void VisitQuoteBlock(QuoteBlock block) => VisitBlock(block);
    protected virtual void VisitCalloutBlock(CalloutBlock block) => VisitBlock(block);
    protected virtual void VisitDetailsBlock(DetailsBlock block) => VisitBlock(block);
    protected virtual void VisitSummaryBlock(SummaryBlock block) => VisitBlock(block);
    protected virtual void VisitOrderedListBlock(OrderedListBlock block) => VisitBlock(block);
    protected virtual void VisitUnorderedListBlock(UnorderedListBlock block) => VisitBlock(block);
    protected virtual void VisitTableBlock(TableBlock block) => VisitBlock(block);
    protected virtual void VisitDefinitionListBlock(DefinitionListBlock block) => VisitBlock(block);
    protected virtual void VisitFootnoteDefinitionBlock(FootnoteDefinitionBlock block) => VisitBlock(block);
    protected virtual void VisitCodeBlock(CodeBlock block) => VisitBlock(block);
    protected virtual void VisitSemanticFencedBlock(SemanticFencedBlock block) => VisitBlock(block);
    protected virtual void VisitImageBlock(ImageBlock block) => VisitBlock(block);
    protected virtual void VisitFrontMatterBlock(FrontMatterBlock block) => VisitBlock(block);
    protected virtual void VisitHtmlCommentBlock(HtmlCommentBlock block) => VisitBlock(block);
    protected virtual void VisitHtmlRawBlock(HtmlRawBlock block) => VisitBlock(block);
    protected virtual void VisitHorizontalRuleBlock(HorizontalRuleBlock block) => VisitBlock(block);
    protected virtual void VisitTocBlock(TocBlock block) => VisitBlock(block);

    protected virtual void VisitListItem(ListItem item) => DefaultVisit(item);
    protected virtual void VisitTableCell(TableCell cell) => DefaultVisit(cell);
    protected virtual void VisitDefinitionListGroup(DefinitionListGroup group) => DefaultVisit(group);
    protected virtual void VisitDefinitionListEntry(DefinitionListEntry entry) => DefaultVisit(entry);
    protected virtual void VisitDefinitionListDefinition(DefinitionListDefinition definition) => DefaultVisit(definition);

    protected virtual void VisitInline(MarkdownInline inline) => DefaultVisit(inline);
    protected virtual void VisitInlineSequence(InlineSequence sequence) => VisitInline(sequence);
    protected virtual void VisitLinkInline(LinkInline inline) => VisitInline(inline);
    protected virtual void VisitImageLinkInline(ImageLinkInline inline) => VisitInline(inline);
    protected virtual void VisitImageInline(ImageInline inline) => VisitInline(inline);
    protected virtual void VisitHtmlTagSequenceInline(HtmlTagSequenceInline inline) => VisitInline(inline);
    protected virtual void VisitBoldSequenceInline(BoldSequenceInline inline) => VisitInline(inline);
    protected virtual void VisitItalicSequenceInline(ItalicSequenceInline inline) => VisitInline(inline);
    protected virtual void VisitBoldItalicSequenceInline(BoldItalicSequenceInline inline) => VisitInline(inline);
    protected virtual void VisitStrikethroughSequenceInline(StrikethroughSequenceInline inline) => VisitInline(inline);
    protected virtual void VisitHighlightSequenceInline(HighlightSequenceInline inline) => VisitInline(inline);
    protected virtual void VisitTextRun(TextRun inline) => VisitInline(inline);
    protected virtual void VisitCodeSpanInline(CodeSpanInline inline) => VisitInline(inline);
    protected virtual void VisitFootnoteRefInline(FootnoteRefInline inline) => VisitInline(inline);
    protected virtual void VisitHardBreakInline(HardBreakInline inline) => VisitInline(inline);
    protected virtual void VisitBoldInline(BoldInline inline) => VisitInline(inline);
    protected virtual void VisitItalicInline(ItalicInline inline) => VisitInline(inline);
    protected virtual void VisitBoldItalicInline(BoldItalicInline inline) => VisitInline(inline);
    protected virtual void VisitStrikethroughInline(StrikethroughInline inline) => VisitInline(inline);
    protected virtual void VisitHighlightInline(HighlightInline inline) => VisitInline(inline);
    protected virtual void VisitUnderlineInline(UnderlineInline inline) => VisitInline(inline);
    protected virtual void VisitHtmlRawInline(HtmlRawInline inline) => VisitInline(inline);
}

#pragma warning restore CS1591
