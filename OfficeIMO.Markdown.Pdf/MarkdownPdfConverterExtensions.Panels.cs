using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static void RenderQuoteBlock(PdfCore.PdfDocument pdf, QuoteBlock quote, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        if (quote.Children.Count > 0) {
            RenderBlocksWithPanelRuns(pdf, quote.Children, document, options, visualTheme, visualTheme.QuotePanelStyleSnapshot);
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
        MarkdownPdfSaveOptions options,
        MarkdownPdfVisualTheme visualTheme,
        PdfCore.PanelStyle panelStyle,
        Action<PdfCore.PdfItemCompose>? renderFirstPanelHeader = null) {
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
        MarkdownPdfSaveOptions options,
        MarkdownPdfVisualTheme visualTheme,
        PdfCore.PanelStyle panelStyle,
        Action<PdfCore.PdfItemCompose>? renderFirstPanelHeader,
        ref bool renderedHeader) {
        if (panelBlocks.Count == 0) {
            return;
        }

        IMarkdownBlock[] batch = panelBlocks.ToArray();
        panelBlocks.Clear();
        Action<PdfCore.PdfItemCompose>? header = !renderedHeader ? renderFirstPanelHeader : null;

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
                return quote.Children.Count == 0 || CanRenderBlocksInsidePanel(quote.Children);
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

            if (items[i].Children.Count > 0 && !CanRenderBlocksInsidePanel(items[i].Children)) {
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

    private static bool TryAppendBlocksInsidePanel(PdfCore.PdfParagraphBuilder builder, IReadOnlyList<IMarkdownBlock> blocks, InlineStyle style, MarkdownPdfVisualTheme visualTheme, bool lineBreakBeforeFirst) {
        bool wroteContent = false;
        for (int i = 0; i < blocks.Count; i++) {
            if (!TryAppendBlockInsidePanel(builder, blocks[i], style, visualTheme, ref wroteContent, lineBreakBeforeFirst)) {
                return false;
            }
        }

        return wroteContent;
    }

    private static bool TryAppendBlockInsidePanel(PdfCore.PdfParagraphBuilder builder, IMarkdownBlock block, InlineStyle style, MarkdownPdfVisualTheme visualTheme, ref bool wroteContent, bool lineBreakBeforeFirst) {
        switch (block) {
            case ParagraphBlock paragraph:
                if (IsEmpty(paragraph.Inlines)) {
                    return true;
                }

                StartPanelBlock(builder, ref wroteContent, lineBreakBeforeFirst);
                AppendInlines(builder, paragraph.Inlines, style);
                return true;
            case HeadingBlock heading:
                if (string.IsNullOrWhiteSpace(heading.Text)) {
                    return true;
                }

                StartPanelBlock(builder, ref wroteContent, lineBreakBeforeFirst);
                AppendInlines(builder, heading.Inlines, style.With(bold: true, fontSize: GetPanelHeadingFontSize(heading.Level)));
                return true;
            case UnorderedListBlock unordered:
                AppendUnorderedListInsidePanel(builder, unordered, style, visualTheme, ref wroteContent, lineBreakBeforeFirst);
                return true;
            case OrderedListBlock ordered:
                AppendOrderedListInsidePanel(builder, ordered, style, visualTheme, ref wroteContent, lineBreakBeforeFirst);
                return true;
            case CodeBlock code:
                AppendCodeBlockInsidePanel(builder, code, visualTheme, ref wroteContent, lineBreakBeforeFirst);
                return true;
            case SemanticFencedBlock semantic:
                AppendSemanticBlockInsidePanel(builder, semantic, visualTheme, ref wroteContent, lineBreakBeforeFirst);
                return true;
            case TableBlock table:
                AppendTableInsidePanel(builder, table, style, ref wroteContent, lineBreakBeforeFirst);
                return true;
            case QuoteBlock quote:
                AppendQuoteInsidePanel(builder, quote, style, visualTheme, ref wroteContent, lineBreakBeforeFirst);
                return true;
            case DetailsBlock details:
                AppendDetailsInsidePanel(builder, details, style, visualTheme, ref wroteContent, lineBreakBeforeFirst);
                return true;
            case DefinitionListBlock definitionList:
                AppendDefinitionListInsidePanel(builder, definitionList, style, ref wroteContent, lineBreakBeforeFirst);
                return true;
            case HorizontalRuleBlock:
                StartPanelBlock(builder, ref wroteContent, lineBreakBeforeFirst);
                builder.Text(new string('-', 32));
                return true;
            default:
                return false;
        }
    }

    private static void StartPanelBlock(PdfCore.PdfParagraphBuilder builder, ref bool wroteContent, bool lineBreakBeforeFirst) {
        if (wroteContent) {
            builder.LineBreak();
            builder.LineBreak();
        } else if (lineBreakBeforeFirst) {
            builder.LineBreak();
        }

        wroteContent = true;
    }

    private static void StartPanelLine(PdfCore.PdfParagraphBuilder builder, ref bool wroteContent, bool lineBreakBeforeFirst) {
        if (wroteContent) {
            builder.LineBreak();
        } else if (lineBreakBeforeFirst) {
            builder.LineBreak();
        }

        wroteContent = true;
    }

    private static double GetPanelHeadingFontSize(int level) {
        return level <= 2 ? 12D : level == 3 ? 11D : 10D;
    }

    private static void AppendUnorderedListInsidePanel(PdfCore.PdfParagraphBuilder builder, UnorderedListBlock list, InlineStyle style, MarkdownPdfVisualTheme visualTheme, ref bool wroteContent, bool lineBreakBeforeFirst) {
        for (int i = 0; i < list.Items.Count; i++) {
            AppendListItemInsidePanel(builder, list.Items[i], "•", style, visualTheme, ref wroteContent, lineBreakBeforeFirst);
        }
    }

    private static void AppendOrderedListInsidePanel(PdfCore.PdfParagraphBuilder builder, OrderedListBlock list, InlineStyle style, MarkdownPdfVisualTheme visualTheme, ref bool wroteContent, bool lineBreakBeforeFirst) {
        for (int i = 0; i < list.Items.Count; i++) {
            string marker = (list.Start + i).ToString(CultureInfo.InvariantCulture) + ".";
            AppendListItemInsidePanel(builder, list.Items[i], marker, style, visualTheme, ref wroteContent, lineBreakBeforeFirst);
        }
    }

    private static void AppendListItemInsidePanel(PdfCore.PdfParagraphBuilder builder, ListItem item, string marker, InlineStyle style, MarkdownPdfVisualTheme visualTheme, ref bool wroteContent, bool lineBreakBeforeFirst) {
        StartPanelLine(builder, ref wroteContent, lineBreakBeforeFirst);
        string indent = item.Level <= 0 ? string.Empty : new string(' ', item.Level * 2);
        if (item.IsTask) {
            PdfCore.PdfColor stateColor = item.Checked ? visualTheme.ChecklistCheckedTextColorSnapshot : visualTheme.ChecklistUncheckedTextColorSnapshot;
            ApplyStyle(builder, style.With(bold: true, color: stateColor)).Text(indent + (item.Checked ? "Done: " : "Open: "));
            AppendInlines(builder, item.Content, style.With(color: stateColor));
        } else {
            ApplyStyle(builder, style).Text(indent + marker + " ");
            AppendInlines(builder, item.Content, style);
        }

        for (int paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
            StartPanelLine(builder, ref wroteContent, lineBreakBeforeFirst);
            ApplyStyle(builder, style).Text(indent + "  ");
            AppendInlines(builder, item.AdditionalParagraphs[paragraphIndex], style);
        }

        for (int childIndex = 0; childIndex < item.Children.Count; childIndex++) {
            TryAppendBlockInsidePanel(builder, item.Children[childIndex], style, visualTheme, ref wroteContent, lineBreakBeforeFirst);
        }
    }

    private static void AppendCodeBlockInsidePanel(PdfCore.PdfParagraphBuilder builder, CodeBlock code, MarkdownPdfVisualTheme visualTheme, ref bool wroteContent, bool lineBreakBeforeFirst) {
        StartPanelBlock(builder, ref wroteContent, lineBreakBeforeFirst);
        string language = code.Language?.Trim() ?? string.Empty;
        if (language.Length > 0) {
            builder.FontSize(visualTheme.CodeBlockLabelFontSizeSnapshot)
                .Bold(language, visualTheme.CodeBlockLabelColorSnapshot)
                .LineBreak();
        }

        builder.Font(PdfCore.PdfStandardFont.Courier)
            .FontSize(visualTheme.CodeBlockFontSizeSnapshot)
            .Color(visualTheme.CodeBlockTextColorSnapshot);
        AppendTextWithLineBreaks(builder, code.Content ?? string.Empty);
        ApplyStyle(builder, CreateInlineStyle(visualTheme));
    }

    private static void AppendSemanticBlockInsidePanel(PdfCore.PdfParagraphBuilder builder, SemanticFencedBlock semantic, MarkdownPdfVisualTheme visualTheme, ref bool wroteContent, bool lineBreakBeforeFirst) {
        StartPanelBlock(builder, ref wroteContent, lineBreakBeforeFirst);
        string label = semantic.SemanticKind;
        if (!string.IsNullOrWhiteSpace(semantic.Language)) {
            label += " / " + semantic.Language;
        }

        builder.FontSize(visualTheme.SemanticBlockLabelFontSizeSnapshot)
            .Bold(label, visualTheme.SemanticBlockLabelColorSnapshot)
            .LineBreak();
        builder.Font(PdfCore.PdfStandardFont.Courier)
            .FontSize(visualTheme.SemanticBlockFontSizeSnapshot)
            .Color(visualTheme.SemanticBlockTextColorSnapshot);
        AppendTextWithLineBreaks(builder, semantic.Content ?? string.Empty);
        ApplyStyle(builder, CreateInlineStyle(visualTheme));
    }

    private static void AppendTableInsidePanel(PdfCore.PdfParagraphBuilder builder, TableBlock table, InlineStyle style, ref bool wroteContent, bool lineBreakBeforeFirst) {
        IReadOnlyList<InlineSequence> headers = table.HeaderInlines;
        IReadOnlyList<IReadOnlyList<InlineSequence>> rows = table.RowInlines;
        if (headers.Count == 0 && rows.Count == 0) {
            return;
        }

        StartPanelBlock(builder, ref wroteContent, lineBreakBeforeFirst);
        if (headers.Count > 0) {
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                if (rowIndex > 0) {
                    builder.LineBreak();
                }

                IReadOnlyList<InlineSequence> row = rows[rowIndex];
                for (int columnIndex = 0; columnIndex < headers.Count; columnIndex++) {
                    if (columnIndex > 0) {
                        builder.Text(" | ");
                    }

                    ApplyStyle(builder, style.With(bold: true)).Text(GetPlainText(headers[columnIndex]) + ": ");
                    if (columnIndex < row.Count) {
                        AppendInlines(builder, row[columnIndex], style);
                    }
                }
            }
            return;
        }

        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            if (rowIndex > 0) {
                builder.LineBreak();
            }

            IReadOnlyList<InlineSequence> row = rows[rowIndex];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                if (columnIndex > 0) {
                    builder.Text(" | ");
                }

                AppendInlines(builder, row[columnIndex], style);
            }
        }
    }

    private static void AppendQuoteInsidePanel(PdfCore.PdfParagraphBuilder builder, QuoteBlock quote, InlineStyle style, MarkdownPdfVisualTheme visualTheme, ref bool wroteContent, bool lineBreakBeforeFirst) {
        InlineStyle quoteStyle = style.With(italic: true);
        if (quote.Children.Count > 0) {
            TryAppendBlocksInsidePanel(builder, quote.Children, quoteStyle, visualTheme, lineBreakBeforeFirst: !wroteContent && lineBreakBeforeFirst);
            wroteContent = true;
            return;
        }

        string text = string.Join(Environment.NewLine, quote.Lines);
        if (!string.IsNullOrWhiteSpace(text)) {
            StartPanelBlock(builder, ref wroteContent, lineBreakBeforeFirst);
            ApplyStyle(builder, quoteStyle);
            AppendTextWithLineBreaks(builder, text);
        }
    }

    private static void AppendDetailsInsidePanel(PdfCore.PdfParagraphBuilder builder, DetailsBlock details, InlineStyle style, MarkdownPdfVisualTheme visualTheme, ref bool wroteContent, bool lineBreakBeforeFirst) {
        if (details.Summary != null) {
            StartPanelBlock(builder, ref wroteContent, lineBreakBeforeFirst);
            ApplyStyle(builder, style.With(bold: true)).Text(details.Open ? "Details: " : "Collapsed details: ");
            AppendInlines(builder, details.Summary.Inlines, style.With(bold: true));
        }

        TryAppendBlocksInsidePanel(builder, details.ChildBlocks, style, visualTheme, lineBreakBeforeFirst: false);
        wroteContent = true;
    }

    private static void AppendDefinitionListInsidePanel(PdfCore.PdfParagraphBuilder builder, DefinitionListBlock definitionList, InlineStyle style, ref bool wroteContent, bool lineBreakBeforeFirst) {
        IReadOnlyList<DefinitionListInlineItem> items = definitionList.InlineItems;
        for (int i = 0; i < items.Count; i++) {
            StartPanelLine(builder, ref wroteContent, lineBreakBeforeFirst);
            AppendInlines(builder, items[i].Term, style.With(bold: true));
            builder.Text(": ");
            AppendInlines(builder, items[i].Definition, style);
        }
    }

    private static string GetPlainText(InlineSequence sequence) {
        var builder = new StringBuilder();
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            if (sequence.Nodes[i] is IPlainTextMarkdownInline plain) {
                plain.AppendPlainText(builder);
            }
        }

        return builder.ToString();
    }
}
