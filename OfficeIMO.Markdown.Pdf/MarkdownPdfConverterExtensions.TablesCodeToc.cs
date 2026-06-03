using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static void RenderTable(PdfCore.PdfDocument pdf, TableBlock table, MarkdownPdfVisualTheme visualTheme) {
        int columnCount = Math.Max(table.Headers.Count, table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Count));
        if (columnCount == 0) {
            return;
        }

        var rows = new List<PdfCore.PdfTableCell[]>();
        if (table.Headers.Count > 0) {
            rows.Add(CreateTableRow(table.HeaderInlines, columnCount, visualTheme));
        }

        IReadOnlyList<IReadOnlyList<InlineSequence>> bodyRows = table.RowInlines;
        for (int rowIndex = 0; rowIndex < bodyRows.Count; rowIndex++) {
            rows.Add(CreateTableRow(bodyRows[rowIndex], columnCount, visualTheme));
        }

        PdfCore.PdfTableStyle style = visualTheme.TableStyleSnapshot;
        style.HeaderRowCount = table.Headers.Count > 0 ? 1 : 0;
        style.RepeatHeaderRowCount = table.Headers.Count > 0 ? 1 : 0;
        style.RightAlignNumeric = true;

        pdf.Table(rows, style: style);
    }

    private static PdfCore.PdfTableCell[] CreateTableRow(IReadOnlyList<InlineSequence> cells, int columnCount, MarkdownPdfVisualTheme visualTheme) {
        var row = new PdfCore.PdfTableCell[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            InlineSequence? sequence = columnIndex < cells.Count ? cells[columnIndex] : null;
            IReadOnlyList<PdfTextRun> runs = sequence == null ? Array.Empty<PdfTextRun>() : ToTextRuns(sequence, CreateInlineStyle(visualTheme));
            row[columnIndex] = runs.Count == 0
                ? PdfCore.PdfTableCell.TextCell(string.Empty)
                : PdfCore.PdfTableCell.RichTextCell(runs);
        }

        return row;
    }

    private static void RenderCodeBlock(PdfCore.PdfDocument pdf, CodeBlock code, MarkdownPdfVisualTheme visualTheme) {
        string content = code.Content ?? string.Empty;
        string language = code.Language?.Trim() ?? string.Empty;

        pdf.PanelParagraph(builder => {
            if (language.Length > 0) {
                builder.FontSize(visualTheme.CodeBlockLabelFontSizeSnapshot)
                    .Bold(language, visualTheme.CodeBlockLabelColorSnapshot)
                    .LineBreak();
            }

            builder.Font(PdfCore.PdfStandardFont.Courier)
                .FontSize(visualTheme.CodeBlockFontSizeSnapshot)
                .Color(visualTheme.CodeBlockTextColorSnapshot);
            AppendTextWithLineBreaks(builder, content);
        }, visualTheme.CodeBlockPanelStyleSnapshot);

        if (!string.IsNullOrWhiteSpace(code.Caption)) {
            pdf.Paragraph(builder => builder.Italic(code.Caption!), style: new PdfCore.PdfParagraphStyle { SpacingAfter = 8 });
        }
    }

    private static void RenderSemanticFencedBlock(PdfCore.PdfDocument pdf, SemanticFencedBlock semantic, MarkdownPdfVisualTheme visualTheme) {
        string label = semantic.SemanticKind;
        if (!string.IsNullOrWhiteSpace(semantic.Language)) {
            label += " / " + semantic.Language;
        }

        pdf.PanelParagraph(builder => {
            builder.FontSize(visualTheme.SemanticBlockLabelFontSizeSnapshot)
                .Bold(label, visualTheme.SemanticBlockLabelColorSnapshot);
            builder.LineBreak();
            builder.Font(PdfCore.PdfStandardFont.Courier)
                .FontSize(visualTheme.SemanticBlockFontSizeSnapshot)
                .Color(visualTheme.SemanticBlockTextColorSnapshot);
            AppendTextWithLineBreaks(builder, semantic.Content ?? string.Empty);
        }, visualTheme.SemanticBlockPanelStyleSnapshot);

        if (!string.IsNullOrWhiteSpace(semantic.Caption)) {
            pdf.Paragraph(builder => builder.Italic(semantic.Caption!), style: new PdfCore.PdfParagraphStyle { SpacingAfter = 8 });
        }
    }

    private static void RenderCalloutBlock(PdfCore.PdfDocument pdf, CalloutBlock callout, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        string title = string.IsNullOrWhiteSpace(callout.Title) ? FormatTitleFromKind(callout.Kind) : callout.Title;

        IReadOnlyList<IMarkdownBlock> children = callout.ChildBlocks;
        bool canRenderChildrenInsidePanel = children.Count > 0 && CanRenderBlocksInsidePanel(children);
        PdfCore.PanelStyle panelStyle = visualTheme.CreateCalloutPanelStyle(callout.Kind);
        if (canRenderChildrenInsidePanel) {
            pdf.Panel(panel => {
                panel.Paragraph(builder => {
                    if (IsEmpty(callout.TitleInlines)) {
                        builder.Bold(title);
                    } else {
                        AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme).With(bold: true));
                    }
                });
                RenderBlocks(pdf, children, document, options, visualTheme);
            }, panelStyle);
            return;
        }

        if (children.Count == 0 && !string.IsNullOrWhiteSpace(callout.Body)) {
            pdf.Panel(panel => {
                panel.Paragraph(builder => {
                    if (IsEmpty(callout.TitleInlines)) {
                        builder.Bold(title);
                    } else {
                        AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme).With(bold: true));
                    }

                    builder.LineBreak();
                    AppendTextWithLineBreaks(builder, callout.Body);
                });
            }, panelStyle);
            return;
        }

        bool renderedInsidePanel = false;
        pdf.PanelParagraph(builder => {
            if (IsEmpty(callout.TitleInlines)) {
                builder.Bold(title);
            } else {
                AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme).With(bold: true));
            }

            if (canRenderChildrenInsidePanel) {
                renderedInsidePanel = TryAppendBlocksInsidePanel(builder, children, CreateInlineStyle(visualTheme), visualTheme, lineBreakBeforeFirst: true);
            } else if (children.Count == 0 && !string.IsNullOrWhiteSpace(callout.Body)) {
                builder.LineBreak();
                AppendTextWithLineBreaks(builder, callout.Body);
                renderedInsidePanel = true;
            }
        }, panelStyle);

        if (children.Count > 0 && !renderedInsidePanel) {
            RenderBlocks(pdf, children, document, options, visualTheme);
        }
    }

    private static void RenderDetailsBlock(PdfCore.PdfDocument pdf, DetailsBlock details, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        if (details.Summary != null) {
            pdf.PanelParagraph(builder => {
                builder.Bold(details.Open ? "Details: " : "Collapsed details: ");
                AppendInlines(builder, details.Summary.Inlines, CreateInlineStyle(visualTheme).With(bold: true));
            }, visualTheme.DetailsPanelStyleSnapshot);
        }

        RenderBlocks(pdf, details.ChildBlocks, document, options, visualTheme);
    }

    private static void RenderDefinitionList(PdfCore.PdfDocument pdf, DefinitionListBlock definitionList, MarkdownPdfVisualTheme visualTheme) {
        IReadOnlyList<DefinitionListInlineItem> items = definitionList.InlineItems;
        if (items.Count == 0) {
            return;
        }

        var rows = new List<PdfCore.PdfTableCell[]>();
        for (int i = 0; i < items.Count; i++) {
            IReadOnlyList<PdfTextRun> termRuns = ToTextRuns(items[i].Term, CreateInlineStyle(visualTheme).With(bold: true));
            IReadOnlyList<PdfTextRun> definitionRuns = ToTextRuns(items[i].Definition, CreateInlineStyle(visualTheme));
            rows.Add(new[] {
                termRuns.Count == 0 ? PdfCore.PdfTableCell.TextCell(string.Empty) : PdfCore.PdfTableCell.RichTextCell(termRuns),
                definitionRuns.Count == 0 ? PdfCore.PdfTableCell.TextCell(string.Empty) : PdfCore.PdfTableCell.RichTextCell(definitionRuns)
            });
        }

        PdfCore.PdfTableStyle style = visualTheme.DefinitionListTableStyleSnapshot;
        style.HeaderRowCount = 0;
        pdf.Table(rows, style: style);
    }

    private static void RenderFootnoteDefinition(PdfCore.PdfDocument pdf, FootnoteDefinitionBlock footnote, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        if (string.IsNullOrWhiteSpace(footnote.Label)) {
            return;
        }

        pdf.Paragraph(builder => {
            builder.Superscript(footnote.Label);
            builder.Text(" ");
            if (footnote.Blocks.Count == 1 && footnote.Blocks[0] is ParagraphBlock paragraph) {
                AppendInlines(builder, paragraph.Inlines, CreateInlineStyle(visualTheme));
            } else {
                AppendTextWithLineBreaks(builder, footnote.Text);
            }
        }, style: new PdfCore.PdfParagraphStyle { SpacingBefore = 4, SpacingAfter = 4 });

        if (footnote.Blocks.Count > 1) {
            for (int i = 1; i < footnote.Blocks.Count; i++) {
                RenderBlock(pdf, footnote.Blocks[i], document, options, visualTheme);
            }
        }
    }

    private static void RenderTocBlock(PdfCore.PdfDocument pdf, TocBlock toc, MarkdownPdfVisualTheme visualTheme) {
        if (toc.Entries.Count == 0) {
            return;
        }

        int baseLevel = toc.NormalizeLevels ? toc.Entries.Min(entry => entry.Level) : 1;
        if (ShouldRenderTocAsPanel(toc)) {
            pdf.PanelParagraph(builder => {
                if (toc.IncludeTitle && !string.IsNullOrWhiteSpace(toc.Title)) {
                    builder.Bold(true)
                        .FontSize(GetTocTitleFontSize(toc.TitleLevel))
                        .Color(visualTheme.TocTitleColorSnapshot)
                        .Text(toc.Title.Trim())
                        .Bold(false)
                        .ResetFontSize()
                        .Color(visualTheme.TocTextColorSnapshot);

                    if (toc.Entries.Count > 0) {
                        builder.LineBreak();
                        builder.LineBreak();
                    }
                } else {
                    builder.Color(visualTheme.TocTextColorSnapshot);
                }

                for (int i = 0; i < toc.Entries.Count; i++) {
                    if (i > 0) {
                        builder.LineBreak();
                    }

                    AppendTocEntry(builder, toc.Entries[i], i, baseLevel, toc.Ordered, visualTheme);
                }
            }, visualTheme.TocPanelStyleSnapshot, defaultColor: visualTheme.TocTextColorSnapshot);
            return;
        }

        var items = new List<PdfCore.PdfListItem>();
        for (int i = 0; i < toc.Entries.Count; i++) {
            TocBlock.Entry entry = toc.Entries[i];
            string prefix = new string(' ', Math.Max(0, entry.Level - baseLevel) * 2);
            string text = prefix + entry.Text;
            items.Add(string.IsNullOrWhiteSpace(entry.Anchor)
                ? PdfCore.PdfListItem.Plain(text)
                : PdfCore.PdfListItem.Rich(new[] { PdfTextRun.LinkToBookmark(text, entry.Anchor, visualTheme.TocLinkColorSnapshot, underline: true, contents: "Table of contents: " + entry.Text) }));
        }

        if (toc.Ordered) {
            pdf.RichNumbered(items);
        } else {
            pdf.RichBullets(items);
        }
    }

    private static bool ShouldRenderTocAsPanel(TocBlock toc) =>
        toc.Layout != TocLayout.List || toc.Chrome == TocChrome.Panel || toc.Chrome == TocChrome.Outline;

    private static void AppendTocEntry(PdfCore.PdfParagraphBuilder builder, TocBlock.Entry entry, int index, int baseLevel, bool ordered, MarkdownPdfVisualTheme visualTheme) {
        int depth = Math.Max(0, entry.Level - baseLevel);
        if (depth > 0) {
            builder.Text(new string(' ', depth * 4));
        }

        string marker = ordered ? (index + 1).ToString(CultureInfo.InvariantCulture) + ". " : "• ";
        builder.Color(visualTheme.TocTextColorSnapshot).Text(marker);
        string text = string.IsNullOrWhiteSpace(entry.Text) ? "(untitled)" : entry.Text.Trim();
        if (string.IsNullOrWhiteSpace(entry.Anchor)) {
            builder.Color(visualTheme.TocLinkColorSnapshot).Text(text);
        } else {
            builder.LinkToBookmark(text, entry.Anchor, visualTheme.TocLinkColorSnapshot, underline: false, contents: "Table of contents: " + text);
        }

        builder.Color(visualTheme.TocTextColorSnapshot);
    }

    private static double GetTocTitleFontSize(int titleLevel) {
        return titleLevel <= 1 ? 15D : titleLevel == 2 ? 13D : 11.5D;
    }
}
