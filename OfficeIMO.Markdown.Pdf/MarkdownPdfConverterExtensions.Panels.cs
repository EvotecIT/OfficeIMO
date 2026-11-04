using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.PdfTextRun;

namespace OfficeIMO.Markdown.Pdf;

// First-party Markdown to PDF conversion helpers.
public static partial class MarkdownPdfConverterExtensions {
    private static void RenderQuoteBlock(PdfCore.PdfDocument pdf, QuoteBlock quote, MarkdownDocument document, MarkdownPdfExportOptions options, MarkdownPdfStyle visualTheme) {
        if (quote.ChildBlocks.Count > 0) {
            RenderBlocksWithPanelRuns(pdf, quote.ChildBlocks, document, options, visualTheme, visualTheme.QuotePanelStyleSnapshot);
            return;
        }

        string text = string.Join(Environment.NewLine, quote.Lines);

        if (!string.IsNullOrWhiteSpace(text)) {
            pdf.PanelParagraph(builder => {
                builder.RunDefaults.Italic = true;
                AppendTextWithLineBreaks(builder, text);
            }, FitPanelStyle(visualTheme.QuotePanelStyleSnapshot, options, out _));
        }
    }

    private static void RenderBlocksWithPanelRuns(
        PdfCore.PdfDocument pdf,
        IReadOnlyList<IMarkdownBlock> blocks,
        MarkdownDocument document,
        MarkdownPdfExportOptions options,
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
                RenderPanel(pdf, panelStyle, options, panel => renderFirstPanelHeader(panel));
                renderedHeader = true;
            }

            RenderBlock(pdf, block, document, options, visualTheme);
        }

        FlushPanelBlocks(pdf, panelBlocks, document, options, visualTheme, panelStyle, renderFirstPanelHeader, ref renderedHeader);
    }

    private static void FlushPanelBlocks(
        PdfCore.PdfDocument pdf,
        List<IMarkdownBlock> panelBlocks,
        MarkdownDocument document,
        MarkdownPdfExportOptions options,
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

        RenderPanel(pdf, panelStyle, options, panel => {
            if (header != null) {
                header(panel);
            }

            RenderBlocks(pdf, batch, document, options, visualTheme);
        });
        renderedHeader = true;
    }

    // Nested panels each take width from the frame. Valid Markdown can nest deeper than a page is wide,
    // so indentation stops at this text width instead of failing the export.
    private const double MinimumPanelTextWidth = 72D;

    private static void RenderPanel(
        PdfCore.PdfDocument pdf,
        PdfCore.PdfPanelStyle style,
        MarkdownPdfExportOptions options,
        Action<PdfCore.PdfContentBuilder> content) {
        PdfCore.PdfPanelStyle fitted = FitPanelStyle(style, options, out double reservedWidth);
        options.PanelReservedWidth += reservedWidth;
        try {
            pdf.Panel(content, fitted);
        } finally {
            options.PanelReservedWidth -= reservedWidth;
        }
    }

    private static double GetPanelFrameWidth(MarkdownPdfExportOptions options) =>
        TryGetAvailablePdfContentWidth(options, out double pageWidth) ? pageWidth - options.PanelReservedWidth : 0D;

    private static PdfCore.PdfPanelStyle FitPanelStyle(PdfCore.PdfPanelStyle style, MarkdownPdfExportOptions options, out double reservedWidth) {
        double requested = style.LeftInset.ToPoints() + style.RightInset.ToPoints() + 2D * style.PaddingX.ToPoints();
        double spare = Math.Max(0D, GetPanelFrameWidth(options) - MinimumPanelTextWidth);
        if (requested <= spare) {
            reservedWidth = requested;
            return style;
        }

        double scale = spare / requested;
        style.LeftInset = OfficeLength.Points(style.LeftInset.ToPoints() * scale);
        style.RightInset = OfficeLength.Points(style.RightInset.ToPoints() * scale);
        style.PaddingX = OfficeLength.Points(style.PaddingX.ToPoints() * scale);
        reservedWidth = spare;
        AddLayoutClampWarning(options, "panel-indent-reduced", "Nested content", "Nested content is deeper than the page width allows, so its indentation was reduced to keep the text readable.");
        return style;
    }

    private static void FitTableColumns(PdfCore.PdfTableStyle style, int columnCount, MarkdownPdfExportOptions options) {
        style.DropPaddingInNarrowColumns = true;
        double padding = (style.CellPaddingLeft ?? style.CellPaddingX).ToPoints() + (style.CellPaddingRight ?? style.CellPaddingX).ToPoints();
        if (GetPanelFrameWidth(options) / columnCount <= padding) {
            AddLayoutClampWarning(options, "table-cell-padding-dropped", "Table", "The table has more columns than the page width allows with cell padding, so the padding was removed from its narrow columns.");
        }
    }

    private static void AddLayoutClampWarning(MarkdownPdfExportOptions options, string code, string source, string message) {
        if (options.Warnings.All(warning => warning.Code != code)) {
            AddWarning(options, code, source, message);
        }
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
