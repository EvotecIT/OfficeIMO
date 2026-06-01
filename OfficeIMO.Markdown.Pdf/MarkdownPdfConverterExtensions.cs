using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static class MarkdownPdfConverterExtensions {
    /// <summary>
    /// Converts Markdown text to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDoc ToPdfDocument(this string markdown, MarkdownPdfSaveOptions? options = null) {
        if (markdown == null) {
            throw new ArgumentNullException(nameof(markdown));
        }

        options ??= new MarkdownPdfSaveOptions();
        MarkdownDoc document = MarkdownReader.Parse(markdown, options.ReaderOptions);
        return document.ToPdfDocument(options);
    }

    /// <summary>
    /// Converts a Markdown file to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDoc ToPdfDocumentFromMarkdownFile(this string path, MarkdownPdfSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Markdown file path cannot be empty.", nameof(path));
        }

        options ??= new MarkdownPdfSaveOptions();
        string fullPath = Path.GetFullPath(path);
        string markdown = File.ReadAllText(fullPath, Encoding.UTF8);
        return MarkdownPdfConverter.ConvertFileMarkdown(markdown, fullPath, options);
    }

    /// <summary>
    /// Converts a Markdown document model to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDoc ToPdfDocument(this MarkdownDoc document, MarkdownPdfSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new MarkdownPdfSaveOptions();
        options.ResetExportState();

        PdfCore.PdfOptions pdfOptions = options.PdfOptions ?? new PdfCore.PdfOptions();
        if (options.CreateOutlineFromHeadings) {
            pdfOptions.CreateOutlineFromHeadings = true;
        }

        MarkdownPdfVisualTheme visualTheme = ResolveVisualTheme(document, options);
        PdfCore.PdfDoc pdf = PdfCore.PdfDoc.Create(pdfOptions);
        PdfCore.PdfTheme? documentTheme = visualTheme.DocumentThemeSnapshot;
        if (documentTheme != null) {
            pdf.Theme(documentTheme);
        }
        visualTheme.ApplyPageDecorations(pdf, pdfOptions);

        IReadOnlyList<IMarkdownBlock> topLevelBlocks = GetPdfTopLevelBlocks(document);
        ApplyMetadata(pdf, document, options);
        string? promotedFrontMatterTitle = GetPromotedFrontMatterTitle(document, options);
        RenderBlocks(pdf, topLevelBlocks, document, options, visualTheme, promotedFrontMatterTitle);
        if (topLevelBlocks.Count == 0) {
            pdf.Paragraph(paragraph => paragraph.Text(string.Empty));
        }

        return pdf;
    }

    private static IReadOnlyList<IMarkdownBlock> GetPdfTopLevelBlocks(MarkdownDoc document) {
        var (blocks, _) = document.GetBlocksAndHeadingSlugs();
        if (document.DocumentHeader == null) {
            return blocks;
        }

        var withFrontMatter = new List<IMarkdownBlock>(blocks.Count + 1) {
            document.DocumentHeader
        };
        withFrontMatter.AddRange(blocks);
        return withFrontMatter;
    }

    /// <summary>
    /// Converts Markdown text to PDF bytes.
    /// </summary>
    public static byte[] SaveAsPdf(this string markdown, MarkdownPdfSaveOptions? options = null) {
        return markdown.ToPdfDocument(options).ToBytes();
    }

    /// <summary>
    /// Saves Markdown text as a PDF file.
    /// </summary>
    public static void SaveAsPdf(this string markdown, string path, MarkdownPdfSaveOptions? options = null) {
        markdown.ToPdfDocument(options).Save(path);
    }

    /// <summary>
    /// Writes Markdown text as PDF to a stream.
    /// </summary>
    public static void SaveAsPdf(this string markdown, Stream stream, MarkdownPdfSaveOptions? options = null) {
        markdown.ToPdfDocument(options).Save(stream);
    }

    /// <summary>
    /// Converts a Markdown document model to PDF bytes.
    /// </summary>
    public static byte[] SaveAsPdf(this MarkdownDoc document, MarkdownPdfSaveOptions? options = null) {
        return document.ToPdfDocument(options).ToBytes();
    }

    /// <summary>
    /// Saves a Markdown document model as a PDF file.
    /// </summary>
    public static void SaveAsPdf(this MarkdownDoc document, string path, MarkdownPdfSaveOptions? options = null) {
        document.ToPdfDocument(options).Save(path);
    }

    /// <summary>
    /// Writes a Markdown document model as PDF to a stream.
    /// </summary>
    public static void SaveAsPdf(this MarkdownDoc document, Stream stream, MarkdownPdfSaveOptions? options = null) {
        document.ToPdfDocument(options).Save(stream);
    }

    private static MarkdownPdfVisualTheme ResolveVisualTheme(MarkdownDoc document, MarkdownPdfSaveOptions options) {
        MarkdownPdfVisualTheme? explicitTheme = options.VisualTheme;
        if (explicitTheme != null) {
            return explicitTheme;
        }

        if (options.UseFrontMatterVisualTheme && document.DocumentHeader != null) {
            string? frontMatterTheme = GetFrontMatterMetadata(document.DocumentHeader, "pdfTheme") ?? GetFrontMatterMetadata(document.DocumentHeader, "pdf-theme");
            if (frontMatterTheme != null) {
                if (MarkdownPdfVisualTheme.TryCreate(frontMatterTheme, out MarkdownPdfVisualTheme? theme)) {
                    return theme!;
                }

                AddWarning(options, "UnsupportedVisualTheme", frontMatterTheme, "The requested Markdown PDF visual theme is not recognized; the configured fallback visual profile is used.");
            }
        }

        return options.ApplyWordLikeTheme
            ? MarkdownPdfVisualTheme.WordLike()
            : MarkdownPdfVisualTheme.Plain();
    }

    private static void RenderBlocks(PdfCore.PdfDoc pdf, IEnumerable<IMarkdownBlock> blocks, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme, string? skipFirstHeadingTitle = null) {
        bool skippedPromotedHeading = false;
        var materializedBlocks = blocks as IReadOnlyList<IMarkdownBlock> ?? blocks.ToList();
        for (int i = 0; i < materializedBlocks.Count; i++) {
            IMarkdownBlock block = materializedBlocks[i];
            if (!skippedPromotedHeading && skipFirstHeadingTitle != null && block is HeadingBlock heading && heading.Level == 1 && IsSameNormalizedText(heading.Text, skipFirstHeadingTitle)) {
                skippedPromotedHeading = true;
                continue;
            }

            if (block is HeadingBlock tocTitleHeading &&
                i + 1 < materializedBlocks.Count &&
                materializedBlocks[i + 1] is TocBlock toc &&
                ShouldRenderTocAsPanel(toc) &&
                toc.IncludeTitle &&
                IsSameNormalizedText(tocTitleHeading.Text, toc.Title)) {
                continue;
            }

            RenderBlock(pdf, block, document, options, visualTheme);
        }
    }

    private static void ApplyMetadata(PdfCore.PdfDoc pdf, MarkdownDoc document, MarkdownPdfSaveOptions options) {
        string? title = NormalizeMetadata(options.Title);
        string? author = NormalizeMetadata(options.Author);
        string? subject = NormalizeMetadata(options.Subject);
        string? keywords = NormalizeMetadata(options.Keywords);

        if (options.UseFrontMatterMetadata) {
            FrontMatterBlock? frontMatter = document.DocumentHeader;
            if (frontMatter != null) {
                title ??= GetFrontMatterMetadata(frontMatter, "title");
                author ??= GetFrontMatterMetadata(frontMatter, "author");
                subject ??= GetFrontMatterMetadata(frontMatter, "subject") ?? GetFrontMatterMetadata(frontMatter, "description") ?? GetFrontMatterMetadata(frontMatter, "summary");
                keywords ??= GetFrontMatterMetadata(frontMatter, "keywords") ?? GetFrontMatterMetadata(frontMatter, "tags");
            }
        }

        if (title == null && options.UseFirstHeadingAsTitle) {
            title = NormalizeMetadata(document.Blocks.OfType<HeadingBlock>().FirstOrDefault()?.Text);
        }

        if (title != null || author != null || subject != null || keywords != null) {
            pdf.Meta(title, author, subject, keywords);
        }
    }

    private static string? GetFrontMatterMetadata(FrontMatterBlock frontMatter, string key) {
        FrontMatterBlock.Entry? entry = frontMatter.FindEntry(key);
        return entry == null ? null : NormalizeMetadata(ConvertMetadataValue(entry.Value));
    }

    private static string? GetPromotedFrontMatterTitle(MarkdownDoc document, MarkdownPdfSaveOptions options) {
        if (options.FrontMatterRenderMode != MarkdownPdfFrontMatterRenderMode.DocumentHeader || document.DocumentHeader == null) {
            return null;
        }

        return GetFrontMatterMetadata(document.DocumentHeader, "title");
    }

    private static string? ConvertMetadataValue(object? value) {
        switch (value) {
            case null:
                return null;
            case string text:
                return text;
            case IEnumerable<string> values:
                return string.Join(", ", values.Where(item => !string.IsNullOrWhiteSpace(item)).Select(item => item.Trim()));
            case System.Collections.IEnumerable values:
                var items = new List<string>();
                foreach (object? item in values) {
                    string? normalized = NormalizeMetadata(Convert.ToString(item, CultureInfo.InvariantCulture));
                    if (normalized != null) {
                        items.Add(normalized);
                    }
                }

                return items.Count == 0 ? null : string.Join(", ", items);
            default:
                return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
    }

    private static void RenderBlock(PdfCore.PdfDoc pdf, IMarkdownBlock block, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        switch (block) {
            case HeadingBlock heading:
                RenderHeading(pdf, heading, document, visualTheme);
                break;
            case ParagraphBlock paragraph:
                RenderParagraph(pdf, paragraph.Inlines, visualTheme);
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
                RenderSemanticFencedBlock(pdf, semantic, visualTheme);
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
                RenderImageBlock(pdf, image, options);
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

    private static void RenderHeading(PdfCore.PdfDoc pdf, HeadingBlock heading, MarkdownDoc document, MarkdownPdfVisualTheme visualTheme) {
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
            pdf.Paragraph(builder => {
                builder.Bold(true).FontSize(fontSize);
                AppendInlines(builder, heading.Inlines, CreateInlineStyle(visualTheme) with { Bold = true, FontSize = fontSize });
            }, style: new PdfCore.PdfParagraphStyle { SpacingBefore = 8, SpacingAfter = 4, KeepWithNext = true });
        }
    }

    private static void RenderParagraph(PdfCore.PdfDoc pdf, InlineSequence inlines, MarkdownPdfVisualTheme visualTheme) {
        if (IsEmpty(inlines)) {
            return;
        }

        pdf.Paragraph(builder => AppendInlines(builder, inlines, CreateInlineStyle(visualTheme)));
    }

    private static void RenderPlainBlock(PdfCore.PdfDoc pdf, string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return;
        }

        pdf.Paragraph(builder => builder.Text(text));
    }

    private static void RenderOrderedList(PdfCore.PdfDoc pdf, OrderedListBlock list, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
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

    private static void RenderUnorderedList(PdfCore.PdfDoc pdf, UnorderedListBlock list, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        if (list.Items.Any(item => item.IsTask)) {
            RenderMixedUnorderedTaskList(pdf, list.Items, options, visualTheme);
            RenderListChildren(pdf, list.Items, document, options, visualTheme);
            return;
        }

        RenderUnorderedListItems(pdf, list.Items, visualTheme);
        RenderListChildren(pdf, list.Items, document, options, visualTheme);
    }

    private static void RenderMixedUnorderedTaskList(PdfCore.PdfDoc pdf, IReadOnlyList<ListItem> sourceItems, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
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

    private static void RenderUnorderedListItems(PdfCore.PdfDoc pdf, IReadOnlyList<ListItem> sourceItems, MarkdownPdfVisualTheme visualTheme) {
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

    private static void RenderChecklistTable(PdfCore.PdfDoc pdf, IReadOnlyList<ListItem> taskItems, MarkdownPdfVisualTheme visualTheme) {
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
            InlineStyle textStyle = CreateInlineStyle(visualTheme) with { Color = textColor };
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

    private static void RenderListChildren(PdfCore.PdfDoc pdf, IReadOnlyList<ListItem> items, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        for (int i = 0; i < items.Count; i++) {
            ListItem item = items[i];
            for (int paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
                RenderParagraph(pdf, item.AdditionalParagraphs[paragraphIndex], visualTheme);
            }

            for (int childIndex = 0; childIndex < item.ChildBlocks.Count; childIndex++) {
                RenderBlock(pdf, item.ChildBlocks[childIndex], document, options, visualTheme);
            }
        }
    }

    private static void RenderTable(PdfCore.PdfDoc pdf, TableBlock table, MarkdownPdfVisualTheme visualTheme) {
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

    private static void RenderCodeBlock(PdfCore.PdfDoc pdf, CodeBlock code, MarkdownPdfVisualTheme visualTheme) {
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

    private static void RenderSemanticFencedBlock(PdfCore.PdfDoc pdf, SemanticFencedBlock semantic, MarkdownPdfVisualTheme visualTheme) {
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

    private static void RenderCalloutBlock(PdfCore.PdfDoc pdf, CalloutBlock callout, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
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
                        AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme) with { Bold = true });
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
                        AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme) with { Bold = true });
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
                AppendInlines(builder, callout.TitleInlines, CreateInlineStyle(visualTheme) with { Bold = true });
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

    private static void RenderDetailsBlock(PdfCore.PdfDoc pdf, DetailsBlock details, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        if (details.Summary != null) {
            pdf.PanelParagraph(builder => {
                builder.Bold(details.Open ? "Details: " : "Collapsed details: ");
                AppendInlines(builder, details.Summary.Inlines, CreateInlineStyle(visualTheme) with { Bold = true });
            }, visualTheme.DetailsPanelStyleSnapshot);
        }

        RenderBlocks(pdf, details.ChildBlocks, document, options, visualTheme);
    }

    private static void RenderDefinitionList(PdfCore.PdfDoc pdf, DefinitionListBlock definitionList, MarkdownPdfVisualTheme visualTheme) {
        IReadOnlyList<DefinitionListInlineItem> items = definitionList.InlineItems;
        if (items.Count == 0) {
            return;
        }

        var rows = new List<PdfCore.PdfTableCell[]>();
        for (int i = 0; i < items.Count; i++) {
            IReadOnlyList<PdfTextRun> termRuns = ToTextRuns(items[i].Term, CreateInlineStyle(visualTheme) with { Bold = true });
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

    private static void RenderFootnoteDefinition(PdfCore.PdfDoc pdf, FootnoteDefinitionBlock footnote, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
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

    private static void RenderTocBlock(PdfCore.PdfDoc pdf, TocBlock toc, MarkdownPdfVisualTheme visualTheme) {
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

    private static void RenderQuoteBlock(PdfCore.PdfDoc pdf, QuoteBlock quote, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        if (quote.Children.Count > 0) {
            bool canRenderChildrenInsidePanel = CanRenderBlocksInsidePanel(quote.Children);
            if (canRenderChildrenInsidePanel) {
                pdf.Panel(_ => RenderBlocks(pdf, quote.Children, document, options, visualTheme), visualTheme.QuotePanelStyleSnapshot);
                return;
            }

            bool renderedInsidePanel = false;
            pdf.PanelParagraph(builder => {
                builder.Italic(true);
                if (canRenderChildrenInsidePanel) {
                    renderedInsidePanel = TryAppendBlocksInsidePanel(builder, quote.Children, CreateInlineStyle(visualTheme) with { Italic = true }, visualTheme, lineBreakBeforeFirst: false);
                } else {
                    builder.Italic("Quote");
                }
            }, visualTheme.QuotePanelStyleSnapshot);

            if (renderedInsidePanel) {
                return;
            }

            RenderBlocks(pdf, quote.Children, document, options, visualTheme);
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
            case ParagraphBlock:
            case HeadingBlock:
            case CodeBlock:
            case SemanticFencedBlock:
            case TableBlock:
            case HorizontalRuleBlock:
            case DefinitionListBlock:
                return true;
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
            if (items[i].ChildBlocks.Count > 0 && !CanRenderBlocksInsidePanel(items[i].ChildBlocks)) {
                return false;
            }
        }

        return true;
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
                AppendInlines(builder, heading.Inlines, style with { Bold = true, FontSize = GetPanelHeadingFontSize(heading.Level) });
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
            ApplyStyle(builder, style with { Bold = true, Color = stateColor }).Text(indent + (item.Checked ? "Done: " : "Open: "));
            AppendInlines(builder, item.Content, style with { Color = stateColor });
        } else {
            ApplyStyle(builder, style).Text(indent + marker + " ");
            AppendInlines(builder, item.Content, style);
        }

        for (int paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
            StartPanelLine(builder, ref wroteContent, lineBreakBeforeFirst);
            ApplyStyle(builder, style).Text(indent + "  ");
            AppendInlines(builder, item.AdditionalParagraphs[paragraphIndex], style);
        }

        for (int childIndex = 0; childIndex < item.ChildBlocks.Count; childIndex++) {
            TryAppendBlockInsidePanel(builder, item.ChildBlocks[childIndex], style, visualTheme, ref wroteContent, lineBreakBeforeFirst);
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

                    ApplyStyle(builder, style with { Bold = true }).Text(GetPlainText(headers[columnIndex]) + ": ");
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
        InlineStyle quoteStyle = style with { Italic = true };
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
            ApplyStyle(builder, style with { Bold = true }).Text(details.Open ? "Details: " : "Collapsed details: ");
            AppendInlines(builder, details.Summary.Inlines, style with { Bold = true });
        }

        TryAppendBlocksInsidePanel(builder, details.ChildBlocks, style, visualTheme, lineBreakBeforeFirst: false);
        wroteContent = true;
    }

    private static void AppendDefinitionListInsidePanel(PdfCore.PdfParagraphBuilder builder, DefinitionListBlock definitionList, InlineStyle style, ref bool wroteContent, bool lineBreakBeforeFirst) {
        IReadOnlyList<DefinitionListInlineItem> items = definitionList.InlineItems;
        for (int i = 0; i < items.Count; i++) {
            StartPanelLine(builder, ref wroteContent, lineBreakBeforeFirst);
            AppendInlines(builder, items[i].Term, style with { Bold = true });
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

    private static void RenderImageBlock(PdfCore.PdfDoc pdf, ImageBlock image, MarkdownPdfSaveOptions options) {
        if (!options.IncludeLocalImages) {
            RenderImagePlaceholder(pdf, image);
            return;
        }

        if (!TryReadImageBytes(image.Path, options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage)) {
            AddWarning(options, warningCode, image.Path, warningMessage);
            RenderImagePlaceholder(pdf, image);
            return;
        }

        OfficeImageInfo? info = OfficeImageReader.TryIdentify(bytes, sourceName, out OfficeImageInfo? detected)
            ? detected
            : null;
        double width = image.Width ?? GetImageWidthPoints(info, options);
        double height = image.Height ?? GetImageHeightPoints(info, width, options);
        string? linkUri = NormalizeAbsoluteLink(image.LinkUrl);
        pdf.Image(bytes, width, height, PdfCore.PdfAlign.Left, spacingBefore: 4, spacingAfter: 6, linkUri: linkUri, linkContents: linkUri == null ? null : image.PlainAlt ?? image.Alt);

        if (!string.IsNullOrWhiteSpace(image.Caption)) {
            pdf.Paragraph(builder => builder.Italic(image.Caption!), style: new PdfCore.PdfParagraphStyle { SpacingAfter = 8 });
        }
    }

    private static void RenderImagePlaceholder(PdfCore.PdfDoc pdf, ImageBlock image) {
        string label = image.PlainAlt ?? image.Alt ?? image.Path;
        if (string.IsNullOrWhiteSpace(label)) {
            label = "Image";
        }

        pdf.Paragraph(builder => builder.Italic("[Image: " + label + "]"));
    }

    private static void RenderFrontMatter(PdfCore.PdfDoc pdf, FrontMatterBlock frontMatter, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        if (frontMatter.Entries.Count == 0) {
            return;
        }

        switch (options.FrontMatterRenderMode) {
            case MarkdownPdfFrontMatterRenderMode.Hidden:
                return;
            case MarkdownPdfFrontMatterRenderMode.DocumentHeader:
                if (RenderFrontMatterDocumentHeader(pdf, frontMatter, document, visualTheme)) {
                    return;
                }

                break;
            case MarkdownPdfFrontMatterRenderMode.Table:
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(options.FrontMatterRenderMode), options.FrontMatterRenderMode, "Unsupported Markdown PDF front matter render mode.");
        }

        RenderFrontMatterTable(pdf, frontMatter, visualTheme);
    }

    private static bool RenderFrontMatterDocumentHeader(PdfCore.PdfDoc pdf, FrontMatterBlock frontMatter, MarkdownDoc document, MarkdownPdfVisualTheme visualTheme) {
        string? title = GetFrontMatterMetadata(frontMatter, "title");
        if (title == null) {
            return false;
        }

        string? anchor = FindMatchingFirstHeadingAnchor(document, title);
        if (!string.IsNullOrWhiteSpace(anchor)) {
            pdf.Bookmark(anchor!);
        }

        pdf.H1(title, PdfCore.PdfAlign.Left, visualTheme.DocumentHeaderTitleColorSnapshot, style: new PdfCore.PdfHeadingStyle {
            FontSize = visualTheme.DocumentHeaderTitleFontSizeSnapshot,
            LineHeight = 1.12,
            SpacingBefore = 0,
            SpacingAfter = 3,
            Color = visualTheme.DocumentHeaderTitleColorSnapshot,
            KeepWithNext = true
        });

        string? subtitle = GetFrontMatterMetadata(frontMatter, "subtitle")
            ?? GetFrontMatterMetadata(frontMatter, "description")
            ?? GetFrontMatterMetadata(frontMatter, "summary")
            ?? GetFrontMatterMetadata(frontMatter, "subject");
        if (subtitle != null) {
            pdf.Paragraph(builder => builder
                    .FontSize(visualTheme.DocumentHeaderSubtitleFontSizeSnapshot)
                    .Color(visualTheme.DocumentHeaderSubtitleColorSnapshot)
                    .Text(subtitle),
                defaultColor: visualTheme.DocumentHeaderSubtitleColorSnapshot,
                style: new PdfCore.PdfParagraphStyle {
                    LineHeight = 1.25,
                    SpacingAfter = 4,
                    KeepWithNext = true
                });
        }

        string? metadataLine = BuildFrontMatterMetadataLine(frontMatter);
        if (metadataLine != null) {
            pdf.Paragraph(builder => builder
                    .FontSize(visualTheme.DocumentHeaderMetadataFontSizeSnapshot)
                    .Color(visualTheme.DocumentHeaderMetadataColorSnapshot)
                    .Text(metadataLine),
                defaultColor: visualTheme.DocumentHeaderMetadataColorSnapshot,
                style: new PdfCore.PdfParagraphStyle {
                    LineHeight = 1.2,
                    SpacingAfter = 6,
                    KeepWithNext = true
                });
        }

        pdf.HR(style: new PdfCore.PdfHorizontalRuleStyle {
            Color = visualTheme.DocumentHeaderRuleColorSnapshot,
            Thickness = 0.8,
            SpacingBefore = 2,
            SpacingAfter = 12,
            KeepWithNext = false
        });
        return true;
    }

    private static void RenderFrontMatterTable(PdfCore.PdfDoc pdf, FrontMatterBlock frontMatter, MarkdownPdfVisualTheme visualTheme) {
        var rows = new List<string[]>();
        rows.Add(new[] { "Key", "Value" });
        for (int i = 0; i < frontMatter.Entries.Count; i++) {
            rows.Add(new[] { frontMatter.Entries[i].Key, ConvertMetadataValue(frontMatter.Entries[i].Value) ?? string.Empty });
        }

        PdfCore.PdfTableStyle style = visualTheme.FrontMatterTableStyleSnapshot;
        style.HeaderRowCount = 1;
        pdf.Table(rows, style: style);
    }

    private static string? BuildFrontMatterMetadataLine(FrontMatterBlock frontMatter) {
        var parts = new List<string>();
        AddMetadataPart(parts, GetFrontMatterMetadata(frontMatter, "author"));
        AddMetadataPart(parts, GetFrontMatterMetadata(frontMatter, "date") ?? GetFrontMatterMetadata(frontMatter, "published") ?? GetFrontMatterMetadata(frontMatter, "updated"));
        string? tags = GetFrontMatterMetadata(frontMatter, "tags") ?? GetFrontMatterMetadata(frontMatter, "keywords");
        if (tags != null) {
            AddMetadataPart(parts, "Tags: " + tags);
        }

        return parts.Count == 0 ? null : string.Join(" | ", parts);
    }

    private static void AddMetadataPart(List<string> parts, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) {
            parts.Add(value!.Trim());
        }
    }

    private static InlineStyle CreateInlineStyle(MarkdownPdfVisualTheme visualTheme) =>
        InlineStyle.Default with {
            LinkColor = visualTheme.LinkColorSnapshot,
            UnderlineLinks = visualTheme.UnderlineLinksSnapshot
        };

    private static void AppendInlines(PdfCore.PdfParagraphBuilder builder, InlineSequence sequence, InlineStyle style) {
        foreach (IMarkdownInline inline in sequence.Nodes) {
            AppendInline(builder, inline, style);
        }
    }

    private static void AppendInline(PdfCore.PdfParagraphBuilder builder, IMarkdownInline inline, InlineStyle style) {
        switch (inline) {
            case OfficeIMO.Markdown.TextRun text:
                ApplyStyle(builder, style).Text(text.Text);
                break;
            case BoldInline bold:
                ApplyStyle(builder, style with { Bold = true }).Text(bold.Text);
                break;
            case ItalicInline italic:
                ApplyStyle(builder, style with { Italic = true }).Text(italic.Text);
                break;
            case BoldItalicInline boldItalic:
                ApplyStyle(builder, style with { Bold = true, Italic = true }).Text(boldItalic.Text);
                break;
            case UnderlineInline underline:
                ApplyStyle(builder, style with { Underline = true }).Text(underline.Text);
                break;
            case StrikethroughInline strike:
                ApplyStyle(builder, style with { Strike = true }).Text(strike.Text);
                break;
            case HighlightInline highlight:
                ApplyStyle(builder, style with { Background = PdfCore.PdfColor.FromRgb(254, 243, 199) }).Text(highlight.Text);
                break;
            case CodeSpanInline code:
                ApplyStyle(builder, style with { Background = PdfCore.PdfColor.FromRgb(241, 245, 249), Color = PdfCore.PdfColor.FromRgb(30, 41, 59) }).Text(code.Text);
                break;
            case LinkInline link:
                AppendLinkInline(builder, link, style);
                break;
            case ImageInline image:
                ApplyStyle(builder, style with { Italic = true }).Text("[Image: " + (image.PlainAlt.Length == 0 ? image.Src : image.PlainAlt) + "]");
                break;
            case ImageLinkInline imageLink:
                ApplyStyle(builder, style with { Italic = true }).Text("[Image: " + (imageLink.PlainAlt.Length == 0 ? imageLink.ImageUrl : imageLink.PlainAlt) + "]");
                break;
            case HardBreakInline:
                builder.LineBreak();
                break;
            case HtmlTagSequenceInline htmlTag:
                AppendHtmlTagInline(builder, htmlTag, style);
                break;
            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                AppendInlines(builder, container.NestedInlines!, style);
                break;
            case IPlainTextMarkdownInline plain:
                var textBuilder = new StringBuilder();
                plain.AppendPlainText(textBuilder);
                ApplyStyle(builder, style).Text(textBuilder.ToString());
                break;
        }
    }

    private static void AppendLinkInline(PdfCore.PdfParagraphBuilder builder, LinkInline link, InlineStyle style) {
        string label = string.IsNullOrEmpty(link.Text) ? link.Url : link.Text;
        bool underline = style.UnderlineLinks ?? true;
        InlineStyle linkStyle = style with { Underline = underline, Color = style.LinkColor ?? PdfCore.PdfColor.FromRgb(37, 99, 235) };
        if (TryGetBookmarkTarget(link.Url, out string? bookmark)) {
            ApplyStyle(builder, linkStyle).LinkToBookmark(label, bookmark!, color: linkStyle.Color, underline: underline, contents: link.Title ?? label);
            return;
        }

        string? absolute = NormalizeAbsoluteLink(link.Url);
        if (absolute != null) {
            ApplyStyle(builder, linkStyle).Link(label, absolute, color: linkStyle.Color, underline: underline, contents: link.Title ?? label);
            return;
        }

        ApplyStyle(builder, style).Text(label);
    }

    private static void AppendHtmlTagInline(PdfCore.PdfParagraphBuilder builder, HtmlTagSequenceInline htmlTag, InlineStyle style) {
        InlineStyle tagStyle = htmlTag.TagName switch {
            "strong" or "b" => style with { Bold = true },
            "em" or "i" => style with { Italic = true },
            "u" or "ins" => style with { Underline = true },
            "del" or "s" => style with { Strike = true },
            "sup" => style with { Baseline = PdfCore.PdfTextBaseline.Superscript },
            "sub" => style with { Baseline = PdfCore.PdfTextBaseline.Subscript },
            "mark" => style with { Background = PdfCore.PdfColor.FromRgb(254, 243, 199) },
            _ => style
        };
        AppendInlines(builder, htmlTag.Inlines, tagStyle);
    }

    private static IReadOnlyList<PdfTextRun> ToTextRuns(InlineSequence sequence, InlineStyle style) {
        var runs = new List<PdfTextRun>();
        foreach (IMarkdownInline inline in sequence.Nodes) {
            AddTextRuns(runs, inline, style);
        }

        return runs;
    }

    private static void AddTextRuns(List<PdfTextRun> runs, IMarkdownInline inline, InlineStyle style) {
        switch (inline) {
            case OfficeIMO.Markdown.TextRun text:
                runs.Add(CreateRun(text.Text, style));
                break;
            case BoldInline bold:
                runs.Add(CreateRun(bold.Text, style with { Bold = true }));
                break;
            case ItalicInline italic:
                runs.Add(CreateRun(italic.Text, style with { Italic = true }));
                break;
            case BoldItalicInline boldItalic:
                runs.Add(CreateRun(boldItalic.Text, style with { Bold = true, Italic = true }));
                break;
            case UnderlineInline underline:
                runs.Add(CreateRun(underline.Text, style with { Underline = true }));
                break;
            case StrikethroughInline strike:
                runs.Add(CreateRun(strike.Text, style with { Strike = true }));
                break;
            case HighlightInline highlight:
                runs.Add(CreateRun(highlight.Text, style with { Background = PdfCore.PdfColor.FromRgb(254, 243, 199) }));
                break;
            case CodeSpanInline code:
                runs.Add(CreateRun(code.Text, style with { Background = PdfCore.PdfColor.FromRgb(241, 245, 249), Color = PdfCore.PdfColor.FromRgb(30, 41, 59) }));
                break;
            case LinkInline link:
                AddLinkRun(runs, link, style);
                break;
            case HardBreakInline:
                runs.Add(PdfTextRun.LineBreak());
                break;
            case HtmlTagSequenceInline htmlTag:
                InlineStyle tagStyle = htmlTag.TagName switch {
                    "strong" or "b" => style with { Bold = true },
                    "em" or "i" => style with { Italic = true },
                    "u" or "ins" => style with { Underline = true },
                    "del" or "s" => style with { Strike = true },
                    "sup" => style with { Baseline = PdfCore.PdfTextBaseline.Superscript },
                    "sub" => style with { Baseline = PdfCore.PdfTextBaseline.Subscript },
                    "mark" => style with { Background = PdfCore.PdfColor.FromRgb(254, 243, 199) },
                    _ => style
                };
                foreach (IMarkdownInline nested in htmlTag.Inlines.Nodes) {
                    AddTextRuns(runs, nested, tagStyle);
                }
                break;
            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                foreach (IMarkdownInline nested in container.NestedInlines!.Nodes) {
                    AddTextRuns(runs, nested, style);
                }
                break;
            case IPlainTextMarkdownInline plain:
                var textBuilder = new StringBuilder();
                plain.AppendPlainText(textBuilder);
                runs.Add(CreateRun(textBuilder.ToString(), style));
                break;
        }
    }

    private static void AddLinkRun(List<PdfTextRun> runs, LinkInline link, InlineStyle style) {
        string label = string.IsNullOrEmpty(link.Text) ? link.Url : link.Text;
        PdfCore.PdfColor linkColor = style.LinkColor ?? PdfCore.PdfColor.FromRgb(37, 99, 235);
        bool underline = style.UnderlineLinks ?? true;
        if (TryGetBookmarkTarget(link.Url, out string? bookmark)) {
            runs.Add(CreateLinkRun(label, style, linkColor, underline, link.Title ?? label, uri: null, bookmark: bookmark));
            return;
        }

        string? absolute = NormalizeAbsoluteLink(link.Url);
        if (absolute != null) {
            runs.Add(CreateLinkRun(label, style, linkColor, underline, link.Title ?? label, uri: absolute, bookmark: null));
            return;
        }

        runs.Add(CreateRun(label, style));
    }

    private static PdfTextRun CreateLinkRun(string text, InlineStyle style, PdfCore.PdfColor color, bool underline, string contents, string? uri, string? bookmark) =>
        new PdfTextRun(
            text,
            bold: style.Bold,
            underline: underline,
            color: color,
            italic: style.Italic,
            strike: style.Strike,
            fontSize: style.FontSize,
            font: style.Font,
            linkUri: uri,
            linkContents: contents,
            baseline: style.Baseline,
            linkDestinationName: bookmark,
            backgroundColor: style.Background);

    private static PdfTextRun CreateRun(string text, InlineStyle style) {
        return new PdfTextRun(
            text,
            bold: style.Bold,
            underline: style.Underline,
            color: style.Color,
            italic: style.Italic,
            strike: style.Strike,
            fontSize: style.FontSize,
            font: style.Font,
            baseline: style.Baseline,
            backgroundColor: style.Background);
    }

    private static PdfCore.PdfParagraphBuilder ApplyStyle(PdfCore.PdfParagraphBuilder builder, InlineStyle style) {
        builder.Bold(style.Bold)
            .Italic(style.Italic)
            .Underline(style.Underline)
            .Strike(style.Strike)
            .Baseline(style.Baseline);

        if (style.FontSize.HasValue) {
            builder.FontSize(style.FontSize.Value);
        } else {
            builder.ResetFontSize();
        }

        if (style.Font.HasValue) {
            builder.Font(style.Font.Value);
        } else {
            builder.ResetFont();
        }

        if (style.Color.HasValue) {
            builder.Color(style.Color.Value);
        } else {
            builder.ResetColor();
        }

        if (style.Background.HasValue) {
            builder.BackgroundColor(style.Background.Value);
        } else {
            builder.ResetBackgroundColor();
        }

        return builder;
    }

    private static void AppendTextWithLineBreaks(PdfCore.PdfParagraphBuilder builder, string text) {
        string[] lines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int i = 0; i < lines.Length; i++) {
            if (i > 0) {
                builder.LineBreak();
            }

            builder.Text(lines[i]);
        }
    }

    private static bool IsEmpty(InlineSequence sequence) {
        if (sequence.Nodes.Count == 0) {
            return true;
        }

        var builder = new StringBuilder();
        foreach (IMarkdownInline inline in sequence.Nodes) {
            if (inline is IPlainTextMarkdownInline plain) {
                plain.AppendPlainText(builder);
            }
        }

        return string.IsNullOrWhiteSpace(builder.ToString());
    }

    private static string? NormalizeMetadata(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string normalized = value!.Trim();
        return normalized.Length == 0 ? null : normalized;
    }

    private static string? FindMatchingFirstHeadingAnchor(MarkdownDoc document, string title) {
        for (int i = 0; i < document.TopLevelBlocks.Count; i++) {
            if (document.TopLevelBlocks[i] is HeadingBlock heading && heading.Level == 1 && IsSameNormalizedText(heading.Text, title)) {
                return document.GetHeadingAnchor(heading);
            }
        }

        return null;
    }

    private static bool IsSameNormalizedText(string? left, string? right) {
        string? normalizedLeft = NormalizeMetadata(left);
        string? normalizedRight = NormalizeMetadata(right);
        return normalizedLeft != null
            && normalizedRight != null
            && string.Equals(normalizedLeft, normalizedRight, StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatTitleFromKind(string? kind) {
        if (string.IsNullOrWhiteSpace(kind)) {
            return "Note";
        }

        string trimmed = kind!.Trim();
        return trimmed.Length == 1
            ? trimmed.ToUpperInvariant()
            : char.ToUpperInvariant(trimmed[0]) + trimmed.Substring(1).ToLowerInvariant();
    }

    private static bool TryReadImageBytes(string path, MarkdownPdfSaveOptions options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage) {
        bytes = Array.Empty<byte>();
        sourceName = string.Empty;
        warningCode = "UnsupportedImage";
        warningMessage = "Only resolvable local Markdown images or supported base64 data URI images are embedded in the Markdown PDF adapter.";

        if (IsDataUri(path)) {
            return TryReadDataUriImageBytes(path, options, out bytes, out sourceName, out warningCode, out warningMessage);
        }

        if (TryCreateRemoteImageUri(path, out Uri? remoteUri)) {
            return TryReadRemoteImageBytes(remoteUri!, options, out bytes, out sourceName, out warningCode, out warningMessage);
        }

        string? resolvedPath = ResolveImagePath(path, options.BaseDirectory);
        if (resolvedPath == null) {
            return false;
        }

        bytes = File.ReadAllBytes(resolvedPath);
        sourceName = resolvedPath;
        return true;
    }

    private static string? ResolveImagePath(string path, string? baseDirectory) {
        if (string.IsNullOrWhiteSpace(path) || Uri.TryCreate(path, UriKind.Absolute, out Uri? uri) && !uri.IsFile) {
            return null;
        }

        string candidate = path;
        if (Uri.TryCreate(path, UriKind.Absolute, out Uri? fileUri) && fileUri.IsFile) {
            candidate = fileUri.LocalPath;
        } else if (!Path.IsPathRooted(candidate) && !string.IsNullOrWhiteSpace(baseDirectory)) {
            candidate = Path.Combine(baseDirectory!, candidate);
        }

        return File.Exists(candidate) ? Path.GetFullPath(candidate) : null;
    }

    private static bool TryReadRemoteImageBytes(Uri uri, MarkdownPdfSaveOptions options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage) {
        bytes = Array.Empty<byte>();
        sourceName = uri.ToString();
        warningCode = "UnsupportedImage";
        warningMessage = "Remote Markdown images require MarkdownPdfSaveOptions.RemoteImageResolver so callers can choose their own download, cache, and trust policy.";

        if (options.RemoteImageResolver == null) {
            return false;
        }

        byte[]? resolvedBytes;
        try {
            resolvedBytes = options.RemoteImageResolver(uri);
        } catch (Exception ex) when (ex is not OutOfMemoryException) {
            warningCode = "RemoteImageResolverFailed";
            warningMessage = "The configured Markdown remote image resolver failed: " + ex.Message;
            return false;
        }

        if (resolvedBytes == null || resolvedBytes.Length == 0) {
            warningMessage = "The configured Markdown remote image resolver did not return image bytes.";
            return false;
        }

        if (resolvedBytes.Length > options.MaximumRemoteImageBytes) {
            warningCode = "ImageTooLarge";
            warningMessage = "The resolved Markdown remote image exceeds the configured maximum byte length.";
            return false;
        }

        bytes = resolvedBytes;
        return true;
    }

    private static bool TryReadDataUriImageBytes(string path, MarkdownPdfSaveOptions options, out byte[] bytes, out string sourceName, out string warningCode, out string warningMessage) {
        bytes = Array.Empty<byte>();
        sourceName = string.Empty;
        warningCode = "UnsupportedImage";
        warningMessage = "The Markdown data URI image is not a supported base64 PNG or JPEG image.";

        if (!options.IncludeDataUriImages) {
            warningMessage = "Data URI images are disabled for this Markdown PDF export.";
            return false;
        }

        int commaIndex = path.IndexOf(',');
        if (commaIndex < 0) {
            return false;
        }

        string metadata = path.Substring("data:".Length, commaIndex - "data:".Length);
        string payload = path.Substring(commaIndex + 1);
        string[] parts = metadata.Split(';');
        string mediaType = parts.Length == 0 ? string.Empty : parts[0].Trim().ToLowerInvariant();
        if (mediaType != "image/png" && mediaType != "image/jpeg" && mediaType != "image/jpg") {
            return false;
        }

        bool isBase64 = false;
        for (int i = 1; i < parts.Length; i++) {
            if (string.Equals(parts[i].Trim(), "base64", StringComparison.OrdinalIgnoreCase)) {
                isBase64 = true;
                break;
            }
        }

        if (!isBase64) {
            return false;
        }

        string compactPayload = RemoveAsciiWhitespace(payload);
        long estimatedBytes = compactPayload.Length * 3L / 4L;
        if (estimatedBytes > options.MaximumDataUriImageBytes + 2L) {
            warningCode = "ImageTooLarge";
            warningMessage = "The decoded Markdown data URI image exceeds the configured maximum byte length.";
            return false;
        }

        try {
            bytes = Convert.FromBase64String(compactPayload);
        } catch (FormatException) {
            return false;
        }

        if (bytes.Length > options.MaximumDataUriImageBytes) {
            bytes = Array.Empty<byte>();
            warningCode = "ImageTooLarge";
            warningMessage = "The decoded Markdown data URI image exceeds the configured maximum byte length.";
            return false;
        }

        sourceName = mediaType == "image/png" ? "data-uri.png" : "data-uri.jpg";
        return true;
    }

    private static bool IsDataUri(string path) {
        return path.StartsWith("data:", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryCreateRemoteImageUri(string path, out Uri? uri) {
        uri = null;
        if (!Uri.TryCreate(path, UriKind.Absolute, out Uri? parsed)) {
            return false;
        }

        if (parsed.Scheme != Uri.UriSchemeHttp && parsed.Scheme != Uri.UriSchemeHttps) {
            return false;
        }

        uri = parsed;
        return true;
    }

    private static string RemoveAsciiWhitespace(string value) {
        StringBuilder? builder = null;
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n') {
                if (builder == null) {
                    builder = new StringBuilder(value.Length);
                    builder.Append(value, 0, i);
                }

                continue;
            }

            builder?.Append(ch);
        }

        return builder == null ? value : builder.ToString();
    }

    private static double GetImageWidthPoints(OfficeImageInfo? info, MarkdownPdfSaveOptions options) {
        if (info == null || info.Width <= 0) {
            return options.DefaultImageWidth;
        }

        return info.Width * 72D / info.DpiX;
    }

    private static double GetImageHeightPoints(OfficeImageInfo? info, double width, MarkdownPdfSaveOptions options) {
        if (info == null || info.Height <= 0 || info.Width <= 0) {
            return options.DefaultImageHeight;
        }

        return width * info.Height / info.Width;
    }

    private static bool TryGetBookmarkTarget(string? value, out string? bookmarkName) {
        bookmarkName = null;
        if (string.IsNullOrWhiteSpace(value) || !value!.StartsWith("#", StringComparison.Ordinal) || value.Length == 1) {
            return false;
        }

        bookmarkName = value.Substring(1);
        return !string.IsNullOrWhiteSpace(bookmarkName);
    }

    private static string? NormalizeAbsoluteLink(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        return Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeMailto)
            ? uri.ToString()
            : null;
    }

    private static void AddWarning(MarkdownPdfSaveOptions options, string code, string source, string message) {
        options.Warnings.Add(new MarkdownPdfExportWarning(code, source, message));
    }

    private readonly record struct InlineStyle(
        bool Bold,
        bool Italic,
        bool Underline,
        bool Strike,
        PdfCore.PdfTextBaseline Baseline,
        PdfCore.PdfColor? Color,
        PdfCore.PdfColor? Background,
        double? FontSize,
        PdfCore.PdfStandardFont? Font,
        PdfCore.PdfColor? LinkColor,
        bool? UnderlineLinks) {
        public static InlineStyle Default { get; } = new InlineStyle(false, false, false, false, PdfCore.PdfTextBaseline.Normal, null, null, null, null, null, null);
    }
}
