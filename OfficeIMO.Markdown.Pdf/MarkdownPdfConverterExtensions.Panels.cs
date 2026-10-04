using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.PdfTextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static void RenderQuoteBlock(PdfCore.PdfDocument pdf, QuoteBlock quote, MarkdownDoc document, MarkdownToPdfOptions options, MarkdownPdfStyle visualTheme) {
        if (quote.ChildBlocks.Count > 0) {
            RenderBlocksWithPanelRuns(pdf, quote.ChildBlocks, document, options, visualTheme, visualTheme.QuotePanelStyleSnapshot);
            return;
        }

        string text = string.Join(Environment.NewLine, quote.Lines);

        if (!string.IsNullOrWhiteSpace(text)) {
            pdf.PanelParagraph(builder => {
                builder.Italic(true);
                AppendTextWithLineBreaks(builder, text);
            }, visualTheme.QuotePanelStyleSnapshot);
        }
    }

    private static void RenderBlocksWithPanelRuns(
        PdfCore.PdfDocument pdf,
        IReadOnlyList<IMarkdownBlock> blocks,
        MarkdownDoc document,
        MarkdownToPdfOptions options,
        MarkdownPdfStyle visualTheme,
        PdfCore.PdfPanelStyle panelStyle,
        Action<PdfCore.PdfContentBuilder>? renderFirstPanelHeader = null) {
        var panelBlocks = new List<IMarkdownBlock>();
        bool renderedHeader = false;

        for (int i = 0; i < blocks.Count; i++) {
            IMarkdownBlock block = blocks[i];
            if (CanRenderBlockInsidePanel(block)) {
                panelBlocks.Add(block);
                continue;
            }

            FlushPanelBlocks(pdf, panelBlocks, document, options, visualTheme, panelStyle, renderFirstPanelHeader, ref renderedHeader);
            if (!renderedHeader && renderFirstPanelHeader != null) {
                pdf.Panel(panel => renderFirstPanelHeader(panel), panelStyle);
                renderedHeader = true;
            }

            RenderBlock(pdf, block, document, options, visualTheme);
        }

        FlushPanelBlocks(pdf, panelBlocks, document, options, visualTheme, panelStyle, renderFirstPanelHeader, ref renderedHeader);
    }

    private static void FlushPanelBlocks(
        PdfCore.PdfDocument pdf,
        List<IMarkdownBlock> panelBlocks,
        MarkdownDoc document,
        MarkdownToPdfOptions options,
        MarkdownPdfStyle visualTheme,
        PdfCore.PdfPanelStyle panelStyle,
        Action<PdfCore.PdfContentBuilder>? renderFirstPanelHeader,
        ref bool renderedHeader) {
        if (panelBlocks.Count == 0) {
            return;
        }

        IMarkdownBlock[] batch = panelBlocks.ToArray();
        panelBlocks.Clear();
        Action<PdfCore.PdfContentBuilder>? header = !renderedHeader ? renderFirstPanelHeader : null;

        pdf.Panel(panel => {
            if (header != null) {
                header(panel);
            }

            RenderBlocks(pdf, batch, document, options, visualTheme);
        }, panelStyle);
        renderedHeader = true;
    }

    private static bool CanRenderBlocksInsidePanel(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks.Count == 0) {
            return false;
        }

        for (int i = 0; i < blocks.Count; i++) {
            if (!CanRenderBlockInsidePanel(blocks[i])) {
                return false;
            }
        }

        return true;
    }

    private static bool CanRenderBlockInsidePanel(IMarkdownBlock block) {
        switch (block) {
            case ParagraphBlock paragraph:
                return !IsImageOnlyParagraph(paragraph);
            case HeadingBlock:
            case CodeBlock:
            case TableBlock:
            case HorizontalRuleBlock:
            case DefinitionListBlock:
                return true;
            case SemanticFencedBlock semantic:
                return !IsChartSemanticFence(semantic);
            case UnorderedListBlock unordered:
                return CanRenderListItemsInsidePanel(unordered.Items);
            case OrderedListBlock ordered:
                return CanRenderListItemsInsidePanel(ordered.Items);
            case QuoteBlock quote:
                return quote.ChildBlocks.Count == 0 || CanRenderBlocksInsidePanel(quote.ChildBlocks);
            case DetailsBlock details:
                return details.ChildBlocks.Count == 0 || CanRenderBlocksInsidePanel(details.ChildBlocks);
            default:
                return false;
        }
    }

    private static bool CanRenderListItemsInsidePanel(IReadOnlyList<ListItem> items) {
        for (int i = 0; i < items.Count; i++) {
            for (int paragraphIndex = 0; paragraphIndex < items[i].AdditionalParagraphs.Count; paragraphIndex++) {
                if (IsImageOnlyInlines(items[i].AdditionalParagraphs[paragraphIndex])) {
                    return false;
                }
            }

            if (items[i].NestedBlocks.Count > 0 && !CanRenderBlocksInsidePanel(items[i].NestedBlocks)) {
                return false;
            }
        }

        return true;
    }

    private static bool IsImageOnlyParagraph(ParagraphBlock paragraph) {
        return IsImageOnlyInlines(paragraph.Inlines);
    }

    private static bool IsImageOnlyInlines(InlineSequence inlines) {
        IReadOnlyList<IMarkdownInline> nodes = inlines.Nodes;
        return nodes.Count == 1 && (nodes[0] is ImageInline || nodes[0] is ImageLinkInline);
    }

}
