using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static void RenderBlock(PdfCore.PdfDocument pdf, IMarkdownBlock block, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        switch (block) {
            case HeadingBlock heading:
                RenderHeading(pdf, heading, document, visualTheme);
                break;
            case ParagraphBlock paragraph:
                RenderParagraph(pdf, paragraph.Inlines, options, visualTheme);
                break;
            case OrderedListBlock ordered:
                RenderOrderedList(pdf, ordered, document, options, visualTheme);
                break;
            case UnorderedListBlock unordered:
                RenderUnorderedList(pdf, unordered, document, options, visualTheme);
                break;
            case TableBlock table:
                RenderTable(pdf, table, visualTheme);
                break;
            case CodeBlock code:
                RenderCodeBlock(pdf, code, visualTheme);
                break;
            case SemanticFencedBlock semantic:
                RenderSemanticFencedBlock(pdf, semantic, options, visualTheme);
                break;
            case CalloutBlock callout:
                RenderCalloutBlock(pdf, callout, document, options, visualTheme);
                break;
            case DetailsBlock details:
                RenderDetailsBlock(pdf, details, document, options, visualTheme);
                break;
            case DefinitionListBlock definitionList:
                RenderDefinitionList(pdf, definitionList, visualTheme);
                break;
            case FootnoteDefinitionBlock footnote:
                RenderFootnoteDefinition(pdf, footnote, document, options, visualTheme);
                break;
            case TocBlock toc:
                RenderTocBlock(pdf, toc, visualTheme);
                break;
            case QuoteBlock quote:
                RenderQuoteBlock(pdf, quote, document, options, visualTheme);
                break;
            case HorizontalRuleBlock:
                pdf.HR();
                break;
            case ImageBlock image:
                RenderImageBlock(pdf, image, options, visualTheme);
                break;
            case FrontMatterBlock frontMatter:
                RenderFrontMatter(pdf, frontMatter, document, options, visualTheme);
                break;
            case HtmlRawBlock html:
                AddWarning(options, "UnsupportedBlock", "HtmlRawBlock", "Raw HTML blocks are rendered as plain text in the Markdown PDF adapter.");
                RenderPlainBlock(pdf, html.Html);
                break;
            case HtmlCommentBlock:
                AddWarning(options, "SkippedBlock", "HtmlCommentBlock", "HTML comments are skipped during Markdown PDF export.");
                break;
            default:
                AddWarning(options, "SimplifiedBlock", block.GetType().Name, "The Markdown block is rendered from its Markdown text representation.");
                RenderPlainBlock(pdf, block.RenderMarkdown());
                break;
        }
    }

    private static void RenderHeading(PdfCore.PdfDocument pdf, HeadingBlock heading, MarkdownDoc document, MarkdownPdfStyle visualTheme) {
        string text = heading.Text;
        if (string.IsNullOrWhiteSpace(text)) {
            return;
        }

        string anchor = document.GetHeadingAnchor(heading);
        if (!string.IsNullOrWhiteSpace(anchor)) {
            pdf.Bookmark(anchor);
        }

        if (heading.Level <= 1) {
            pdf.H1(text);
        } else if (heading.Level == 2) {
            pdf.H2(text);
        } else if (heading.Level == 3) {
            pdf.H3(text);
        } else {
            double fontSize = heading.Level == 4 ? 13D : 11.5D;
            PdfCore.PdfColor headingColor = ResolveManualHeadingColor(visualTheme);
            pdf.Paragraph(builder => {
                builder.Bold(true).FontSize(fontSize);
                AppendInlines(builder, heading.Inlines, CreateInlineStyle(visualTheme).With(bold: true, color: headingColor, fontSize: fontSize));
            }, style: new PdfCore.PdfParagraphStyle { SpacingBefore = 8, SpacingAfter = 4, KeepWithNext = true });
        }
    }

    private static PdfCore.PdfColor ResolveManualHeadingColor(MarkdownPdfStyle visualTheme) {
        PdfCore.PdfHeadingStyles? headingStyles = visualTheme.DocumentThemeSnapshot?.HeadingStyles;
        return headingStyles?.Level3?.Color ??
               headingStyles?.Level2?.Color ??
               headingStyles?.Level1?.Color ??
               visualTheme.DocumentHeaderTitleColorSnapshot;
    }

    private static void RenderParagraph(PdfCore.PdfDocument pdf, InlineSequence inlines, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        if (TryRenderImageOnlyParagraph(pdf, inlines, options, visualTheme)) {
            return;
        }

        if (IsEmpty(inlines)) {
            return;
        }

        pdf.Paragraph(builder => AppendInlines(builder, inlines, CreateInlineStyle(visualTheme)));
    }

    private static void RenderPlainBlock(PdfCore.PdfDocument pdf, string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return;
        }

        pdf.Paragraph(builder => builder.Text(text));
    }

    private static void RenderOrderedList(PdfCore.PdfDocument pdf, OrderedListBlock list, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        var items = new List<PdfCore.PdfListItem>();
        for (int i = 0; i < list.Items.Count; i++) {
            ListItem item = list.Items[i];
            IReadOnlyList<PdfTextRun> runs = CreateListItemRuns(item, includeTaskMarker: true, CreateInlineStyle(visualTheme));
            items.Add(PdfCore.PdfListItem.Rich(runs.Count == 0 ? new[] { PdfTextRun.Normal(string.Empty) } : runs));
        }

        if (items.Count > 0) {
            pdf.RichNumbered(items, startNumber: list.Start);
        }

        RenderListChildren(pdf, list.Items, document, options, visualTheme);
    }

    private static void RenderUnorderedList(PdfCore.PdfDocument pdf, UnorderedListBlock list, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        if (list.Items.Any(item => item.IsTask)) {
            RenderMixedUnorderedTaskList(pdf, list.Items, options, visualTheme);
            RenderListChildren(pdf, list.Items, document, options, visualTheme);
            return;
        }

        RenderUnorderedListItems(pdf, list.Items, visualTheme);
        RenderListChildren(pdf, list.Items, document, options, visualTheme);
    }

    private static void RenderMixedUnorderedTaskList(PdfCore.PdfDocument pdf, IReadOnlyList<ListItem> sourceItems, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        var bulletItems = new List<ListItem>();
        var taskItems = new List<ListItem>();

        for (int i = 0; i < sourceItems.Count; i++) {
            ListItem item = sourceItems[i];
            if (item.IsTask) {
                if (bulletItems.Count > 0) {
                    RenderUnorderedListItems(pdf, bulletItems, visualTheme);
                    bulletItems.Clear();
                }

                taskItems.Add(item);
            } else {
                if (taskItems.Count > 0) {
                    RenderChecklistTable(pdf, taskItems, visualTheme);
                    taskItems.Clear();
                }

                bulletItems.Add(item);
            }
        }

        if (bulletItems.Count > 0) {
            RenderUnorderedListItems(pdf, bulletItems, visualTheme);
        }

        if (taskItems.Count > 0) {
            RenderChecklistTable(pdf, taskItems, visualTheme);
        }
    }

    private static void RenderUnorderedListItems(PdfCore.PdfDocument pdf, IReadOnlyList<ListItem> sourceItems, MarkdownPdfStyle visualTheme) {
        var items = new List<PdfCore.PdfListItem>();
        for (int i = 0; i < sourceItems.Count; i++) {
            ListItem item = sourceItems[i];
            IReadOnlyList<PdfTextRun> runs = CreateListItemRuns(item, includeTaskMarker: true, CreateInlineStyle(visualTheme));
            items.Add(PdfCore.PdfListItem.Rich(runs.Count == 0 ? new[] { PdfTextRun.Normal(string.Empty) } : runs));
        }

        if (items.Count > 0) {
            pdf.RichBullets(items);
        }
    }

    private static void RenderChecklistTable(PdfCore.PdfDocument pdf, IReadOnlyList<ListItem> taskItems, MarkdownPdfStyle visualTheme) {
        var rows = new List<PdfCore.PdfTableCell[]>();
        var icons = new Dictionary<(int Row, int Column), PdfCore.PdfCellIcon>();
        var fills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
        PdfCore.PdfTableStyle style = visualTheme.ChecklistTableStyleSnapshot;
        double fontSize = style.FontSize ?? 10D;
        double iconSize = Math.Max(8.5D, Math.Min(10.5D, fontSize));
        double opticalBaselineOffset = (iconSize / 2D) - (fontSize * 0.4625D);
        for (int i = 0; i < taskItems.Count; i++) {
            ListItem item = taskItems[i];
            PdfCore.PdfColor iconColor = item.Checked ? visualTheme.ChecklistCheckedIconColorSnapshot : visualTheme.ChecklistUncheckedIconColorSnapshot;
            PdfCore.PdfColor textColor = item.Checked ? visualTheme.ChecklistCheckedTextColorSnapshot : visualTheme.ChecklistUncheckedTextColorSnapshot;
            PdfCore.PdfColor? fillColor = item.Checked ? visualTheme.ChecklistCheckedFillColorSnapshot : visualTheme.ChecklistUncheckedFillColorSnapshot;
            InlineStyle textStyle = CreateInlineStyle(visualTheme).With(color: textColor);
            IReadOnlyList<PdfTextRun> runs = CreateListItemRuns(item, includeTaskMarker: false, textStyle);
            icons[(i, 0)] = new PdfCore.PdfCellIcon {
                Kind = item.Checked ? PdfCore.PdfCellIconKind.CheckBoxChecked : PdfCore.PdfCellIconKind.CheckBoxUnchecked,
                Color = iconColor,
                Size = iconSize,
                OffsetY = opticalBaselineOffset
            };
            if (fillColor != null) {
                fills[(i, 0)] = fillColor.Value;
                fills[(i, 1)] = fillColor.Value;
            }

            rows.Add(new[] {
                PdfCore.PdfTableCell.TextCell(string.Empty),
                PdfCore.PdfTableCell.RichTextCell(runs.Count == 0 ? new[] { CreateRun(string.Empty, textStyle) } : runs)
            });
        }

        if (rows.Count > 0) {
            style.HeaderRowCount = 0;
            style.RepeatHeaderRowCount = 0;
            style.CellIcons = icons;
            if (fills.Count > 0) {
                style.CellFills = fills;
            }

            pdf.Table(rows, style: style);
        }
    }

    private static IReadOnlyList<PdfTextRun> CreateListItemRuns(ListItem item, bool includeTaskMarker, InlineStyle style) {
        var runs = new List<PdfTextRun>();
        if (includeTaskMarker && item.IsTask) {
            runs.Add(CreateRun(item.Checked ? "[x] " : "[ ] ", style));
        }

        runs.AddRange(ToTextRuns(item.Content, style));
        return runs;
    }

    private static void RenderListChildren(PdfCore.PdfDocument pdf, IReadOnlyList<ListItem> items, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        for (int i = 0; i < items.Count; i++) {
            ListItem item = items[i];
            for (int paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
                RenderParagraph(pdf, item.AdditionalParagraphs[paragraphIndex], options, visualTheme);
            }

            for (int childIndex = 0; childIndex < item.NestedBlocks.Count; childIndex++) {
                RenderBlock(pdf, item.NestedBlocks[childIndex], document, options, visualTheme);
            }
        }
    }
}
