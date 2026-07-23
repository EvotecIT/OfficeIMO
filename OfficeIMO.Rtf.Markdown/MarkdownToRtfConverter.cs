using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Markdown;

internal static class MarkdownToRtfConverter {
    private const int MarkdownListIdBase = 7000;
    private const int ListIndentTwips = 720;

    internal static RtfDocument Convert(MarkdownDoc markdown, MarkdownToRtfConversionContext context) {
        var document = RtfDocument.Create();
        EnsureDocumentDefaults(document);
        var footnoteDefinitions = BuildFootnoteDefinitions(markdown);

        for (int i = 0; i < markdown.Blocks.Count; i++) {
            if (markdown.Blocks[i] is FootnoteDefinitionBlock) {
                continue;
            }

            ConvertBlock(document, markdown.Blocks[i], context, footnoteDefinitions);
        }

        return document;
    }

    private static void EnsureDocumentDefaults(RtfDocument document) {
        EnsureHighlightColor(document);
        document.AddFont("Consolas");
    }

    private static Dictionary<string, FootnoteDefinitionBlock> BuildFootnoteDefinitions(MarkdownDoc markdown) {
        var definitions = new Dictionary<string, FootnoteDefinitionBlock>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < markdown.Blocks.Count; i++) {
            if (markdown.Blocks[i] is FootnoteDefinitionBlock footnote && !string.IsNullOrEmpty(footnote.Label)) {
                definitions[footnote.Label] = footnote;
            }
        }

        return definitions;
    }

    private static void ConvertBlock(RtfDocument document, IMarkdownBlock block, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        switch (block) {
            case ParagraphBlock paragraph:
                ConvertParagraph(document, paragraph, context, footnoteDefinitions);
                break;
            case HeadingBlock heading:
                ConvertHeading(document, heading, context, footnoteDefinitions);
                break;
            case UnorderedListBlock unorderedList:
                ConvertList(document, unorderedList.Items, RtfListKind.Bullet, 1, context, footnoteDefinitions);
                break;
            case OrderedListBlock orderedList:
                ConvertList(document, orderedList.Items, RtfListKind.Decimal, Math.Max(1, orderedList.Start), context, footnoteDefinitions);
                break;
            case TableBlock table:
                ConvertTable(document, table, context, footnoteDefinitions);
                break;
            case ImageBlock image:
                ConvertImageBlock(document, image, context);
                break;
            case CodeBlock code:
                ConvertCodeBlock(document, code);
                break;
            case HtmlCommentBlock comment:
                ConvertRawHtml(document, comment.Comment, context, "Markdown HTML comment block", footnoteDefinitions);
                break;
            case HtmlRawBlock html:
                ConvertRawHtml(document, html.Html, context, "Markdown raw HTML block", footnoteDefinitions);
                break;
            case QuoteBlock quote:
                ConvertChildBlocks(document, quote.ChildBlocks, context, "Markdown quote flattened to paragraphs.", footnoteDefinitions);
                break;
            case DefinitionListBlock definitionList:
                ConvertDefinitionList(document, definitionList, context, footnoteDefinitions);
                break;
            case IChildMarkdownBlockContainer container:
                ConvertChildBlocks(document, container.ChildBlocks, context, block.GetType().Name + " child blocks flattened.", footnoteDefinitions);
                break;
            default:
                document.AddParagraph(block.RenderMarkdown());
                context.Report("MDRTF001", RtfMarkdownDiagnosticSeverity.Warning, "Markdown block converted using rendered Markdown fallback.", block.GetType().Name, RtfConversionAction.Flattened);
                break;
        }
    }

    private static void ConvertHeading(RtfDocument document, HeadingBlock heading, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        int level = heading.Level < 1 ? 1 : heading.Level > 6 ? 6 : heading.Level;
        int styleId = 100 + level;
        RtfStyle style = document.AddStyle(styleId, "Heading " + level);
        style.OutlineLevel = level - 1;

        RtfParagraph paragraph = document.AddParagraph();
        paragraph.SetStyle(styleId);
        paragraph.OutlineLevel = level - 1;
        AppendInlineSequence(paragraph, heading.Inlines, document, context, InlineStyle.Normal, footnoteDefinitions);
    }

    private static void ConvertDefinitionList(RtfDocument document, DefinitionListBlock definitionList, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        if (definitionList.InlineItems.Count == 0) {
            document.AddParagraph(((IMarkdownBlock)definitionList).RenderMarkdown());
            context.Report("MDRTF017", RtfMarkdownDiagnosticSeverity.Info, "Markdown definition list converted using rendered Markdown fallback.", action: RtfConversionAction.Flattened);
            return;
        }

        for (int i = 0; i < definitionList.InlineItems.Count; i++) {
            DefinitionListInlineItem item = definitionList.InlineItems[i];
            RtfParagraph paragraph = document.AddParagraph();
            AppendInlineSequence(paragraph, item.Term, document, context, InlineStyle.Normal, footnoteDefinitions);
            paragraph.AddText(": ");
            AppendInlineSequence(paragraph, item.Definition, document, context, InlineStyle.Normal, footnoteDefinitions);
        }
    }

    private static void ConvertList(RtfDocument document, IReadOnlyList<ListItem> items, RtfListKind kind, int start, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        int listId = CreateListDefinition(document, kind, start);
        ConvertListItems(document, items, listId, kind, start, 0, context, footnoteDefinitions);
    }

    private static void ConvertListItems(RtfDocument document, IReadOnlyList<ListItem> items, int listId, RtfListKind kind, int start, int levelOffset, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        for (int i = 0; i < items.Count; i++) {
            ListItem item = items[i];
            RtfParagraph paragraph = document.AddParagraph();
            long requestedLevel = (long)levelOffset + item.Level;
            int level = (int)Math.Min(context.MaxListNestingDepth - 1L, Math.Max(0L, requestedLevel));
            if (requestedLevel >= context.MaxListNestingDepth) {
                context.Report(
                    "MDRTF018",
                    RtfMarkdownDiagnosticSeverity.Warning,
                    $"Markdown list nesting was limited to {context.MaxListNestingDepth} levels.",
                    action: RtfConversionAction.Flattened);
            }
            EnsureListLevel(document, listId, level, kind, start);
            paragraph.SetList(listId, level, kind);
            paragraph.ListDefinitionId = listId;
            if (item.IsTask) {
                paragraph.AddText(item.Checked ? "[x] " : "[ ] ");
            }

            AppendInlineSequence(paragraph, item.Content, document, context, InlineStyle.Normal, footnoteDefinitions);

            for (int paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
                RtfParagraph continuation = AddListContinuationParagraph(document, level);
                AppendInlineSequence(continuation, item.AdditionalParagraphs[paragraphIndex], document, context, InlineStyle.Normal, footnoteDefinitions);
            }

            for (int childIndex = 0; childIndex < item.ChildBlocks.Count; childIndex++) {
                if (IsRenderedListItemParagraphBlock(item, item.ChildBlocks[childIndex])) {
                    continue;
                }

                ConvertNestedListOrBlock(document, item.ChildBlocks[childIndex], listId, level, context, footnoteDefinitions);
            }
        }
    }

    private static bool IsRenderedListItemParagraphBlock(ListItem item, IMarkdownBlock block) {
        var paragraphBlocks = item.ParagraphBlocks;
        for (int i = 0; i < paragraphBlocks.Count; i++) {
            if (ReferenceEquals(paragraphBlocks[i], block)) {
                return true;
            }
        }

        return false;
    }

    private static int CreateListDefinition(RtfDocument document, RtfListKind kind, int start) {
        int listId = MarkdownListIdBase + document.ListOverrides.Count + 1;
        document.AddListDefinition(listId, "Markdown list " + listId.ToString(System.Globalization.CultureInfo.InvariantCulture));
        document.AddListOverride(listId, listId);
        EnsureListLevel(document, listId, 0, kind, start);
        return listId;
    }

    private static void EnsureListLevel(RtfDocument document, int listId, int levelIndex, RtfListKind kind, int start) {
        RtfListDefinition? definition = document.ListDefinitions.FirstOrDefault(item => item.Id == listId);
        if (definition == null) {
            definition = document.AddListDefinition(listId, "Markdown list " + listId.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        while (definition.Levels.Count <= levelIndex) {
            definition.AddLevel(kind);
        }

        RtfListLevel level = definition.Levels[levelIndex];
        if (level.Kind != kind) {
            level.Kind = kind;
        }

        if (kind == RtfListKind.Decimal) {
            level.StartAt = Math.Max(1, start);
        }

        RtfListOverride? listOverride = document.ListOverrides.FirstOrDefault(item => item.Id == listId);
        if (listOverride == null) {
            listOverride = document.AddListOverride(listId, listId);
        }

        if (kind == RtfListKind.Decimal && start != 1) {
            EnsureListLevelOverrideCount(listOverride, levelIndex);
            RtfListLevelOverride levelOverride = listOverride.LevelOverrides[levelIndex];
            levelOverride.OverrideStartAt = true;
            levelOverride.StartAt = Math.Max(1, start);
        } else if (kind == RtfListKind.Decimal && listOverride.LevelOverrides.Count > levelIndex) {
            RtfListLevelOverride levelOverride = listOverride.LevelOverrides[levelIndex];
            levelOverride.OverrideStartAt = false;
            levelOverride.StartAt = null;
        }
    }

    private static void EnsureListLevelOverrideCount(RtfListOverride listOverride, int levelIndex) {
        while (listOverride.LevelOverrides.Count <= levelIndex) {
            RtfListLevelOverride paddingOverride = listOverride.AddLevelOverride();
            paddingOverride.OverrideStartAt = false;
        }
    }

    private static void ConvertNestedListOrBlock(RtfDocument document, IMarkdownBlock block, int listId, int parentLevel, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        if (parentLevel >= context.MaxListNestingDepth - 1 &&
            (block is UnorderedListBlock || block is OrderedListBlock)) {
            ConvertFlattenedListAsContinuations(document, block, parentLevel, context, footnoteDefinitions);
            context.Report(
                "MDRTF018",
                RtfMarkdownDiagnosticSeverity.Warning,
                $"Markdown list nesting was limited to {context.MaxListNestingDepth} levels.",
                action: RtfConversionAction.Flattened);
            return;
        }

        switch (block) {
            case UnorderedListBlock unorderedList:
                ConvertListItems(document, unorderedList.Items, listId, RtfListKind.Bullet, 1, parentLevel + 1, context, footnoteDefinitions);
                break;
            case OrderedListBlock orderedList:
                int childListId = CreateListDefinition(document, RtfListKind.Decimal, Math.Max(1, orderedList.Start));
                ConvertListItems(document, orderedList.Items, childListId, RtfListKind.Decimal, Math.Max(1, orderedList.Start), parentLevel + 1, context, footnoteDefinitions);
                break;
            default:
                ConvertNestedContinuationBlock(document, block, parentLevel, context, footnoteDefinitions);
                break;
        }
    }

    private static void ConvertFlattenedListAsContinuations(
        RtfDocument document,
        IMarkdownBlock listBlock,
        int parentLevel,
        MarkdownToRtfConversionContext context,
        IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        var pending = new Stack<object>();
        PushListItems(listBlock, pending);
        while (pending.Count > 0) {
            object next = pending.Pop();
            if (next is IMarkdownBlock continuationBlock) {
                ConvertNestedContinuationBlock(document, continuationBlock, parentLevel, context, footnoteDefinitions);
                continue;
            }

            FlattenedListItem flattenedItem = (FlattenedListItem)next;
            ListItem item = flattenedItem.Item;
            RtfParagraph paragraph = AddListContinuationParagraph(document, parentLevel);
            if (flattenedItem.MarkerText != null) {
                paragraph.AddText(flattenedItem.MarkerText + " ");
            }
            if (item.IsTask) {
                paragraph.AddText(item.Checked ? "[x] " : "[ ] ");
            }
            AppendInlineSequence(paragraph, item.Content, document, context, InlineStyle.Normal, footnoteDefinitions);
            for (int paragraphIndex = 0; paragraphIndex < item.AdditionalParagraphs.Count; paragraphIndex++) {
                AppendInlineSequence(AddListContinuationParagraph(document, parentLevel), item.AdditionalParagraphs[paragraphIndex], document, context, InlineStyle.Normal, footnoteDefinitions);
            }

            for (int childIndex = item.ChildBlocks.Count - 1; childIndex >= 0; childIndex--) {
                IMarkdownBlock child = item.ChildBlocks[childIndex];
                if (IsRenderedListItemParagraphBlock(item, child)) {
                    continue;
                }

                if (child is UnorderedListBlock || child is OrderedListBlock) {
                    PushListItems(child, pending);
                } else {
                    pending.Push(child);
                }
            }
        }
    }

    private static void PushListItems(IMarkdownBlock listBlock, Stack<object> pending) {
        IReadOnlyList<ListItem> items = listBlock is UnorderedListBlock unorderedList
            ? unorderedList.Items
            : ((OrderedListBlock)listBlock).Items;
        for (int index = items.Count - 1; index >= 0; index--) {
            string? markerText = null;
            if (listBlock is OrderedListBlock orderedList) {
                int markerValue = orderedList.Reversed ? orderedList.Start - index : orderedList.Start + index;
                markerText = items[index].MarkerText ?? OrderedListBlock.FormatMarker(markerValue, orderedList.MarkerStyle, orderedList.MarkerDelimiter);
            } else if (listBlock is UnorderedListBlock) {
                markerText = items[index].MarkerText ?? "-";
            }

            pending.Push(new FlattenedListItem(items[index], markerText));
        }
    }

    private sealed class FlattenedListItem {
        internal FlattenedListItem(ListItem item, string? markerText) {
            Item = item;
            MarkerText = markerText;
        }

        internal ListItem Item { get; }
        internal string? MarkerText { get; }
    }

    private static void ConvertNestedContinuationBlock(RtfDocument document, IMarkdownBlock block, int parentLevel, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        switch (block) {
            case ParagraphBlock paragraph:
                AppendInlineSequence(AddListContinuationParagraph(document, parentLevel), paragraph.Inlines, document, context, InlineStyle.Normal, footnoteDefinitions);
                break;
            case HeadingBlock heading:
                AppendInlineSequence(AddListContinuationParagraph(document, parentLevel), heading.Inlines, document, context, InlineStyle.Normal, footnoteDefinitions);
                break;
            case QuoteBlock quote:
                context.Report("MDRTF012", RtfMarkdownDiagnosticSeverity.Info, "Markdown quote inside list item flattened to continuation paragraphs.", action: RtfConversionAction.Flattened);
                ConvertNestedContinuationBlocks(document, quote.ChildBlocks, parentLevel, context, footnoteDefinitions);
                break;
            case CodeBlock code:
                ConvertNestedCodeBlock(document, code, parentLevel);
                break;
            case TableBlock table:
                EnsureTableWithinLimit(table, context);
                context.Report("MDRTF013", RtfMarkdownDiagnosticSeverity.Info, "Markdown table inside list item preserved as continuation text.");
                ConvertRenderedBlockAsContinuation(document, table, parentLevel);
                break;
            case HtmlCommentBlock comment:
                ConvertRawHtmlAsContinuation(document, comment.Comment, parentLevel, context, "Markdown HTML comment block");
                break;
            case HtmlRawBlock html:
                ConvertRawHtmlAsContinuation(document, html.Html, parentLevel, context, "Markdown raw HTML block");
                break;
            case IChildMarkdownBlockContainer container:
                context.Report("MDRTF015", RtfMarkdownDiagnosticSeverity.Info, block.GetType().Name + " inside list item flattened to continuation paragraphs.", action: RtfConversionAction.Flattened);
                ConvertNestedContinuationBlocks(document, container.ChildBlocks, parentLevel, context, footnoteDefinitions);
                break;
            default:
                ConvertRenderedBlockAsContinuation(document, block, parentLevel);
                context.Report("MDRTF016", RtfMarkdownDiagnosticSeverity.Info, "Markdown block inside list item converted using rendered Markdown continuation text.", block.GetType().Name, RtfConversionAction.Flattened);
                break;
        }
    }

    private static void ConvertNestedContinuationBlocks(RtfDocument document, IReadOnlyList<IMarkdownBlock> blocks, int parentLevel, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        for (int i = 0; i < blocks.Count; i++) {
            ConvertNestedContinuationBlock(document, blocks[i], parentLevel, context, footnoteDefinitions);
        }
    }

    private static void ConvertNestedCodeBlock(RtfDocument document, CodeBlock code, int parentLevel) {
        int fontId = document.AddFont("Consolas");
        string[] lines = code.Content.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
        for (int i = 0; i < lines.Length; i++) {
            RtfParagraph paragraph = AddListContinuationParagraph(document, parentLevel);
            paragraph.AddText(lines[i]).FontId = fontId;
        }
    }

    private static void ConvertRenderedBlockAsContinuation(RtfDocument document, IMarkdownBlock block, int parentLevel) {
        string rendered = (block.RenderMarkdown() ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = rendered.Split('\n');
        for (int i = 0; i < lines.Length; i++) {
            if (lines[i].Length == 0) {
                continue;
            }

            AddListContinuationParagraph(document, parentLevel).AddText(lines[i]);
        }
    }

    private static RtfParagraph AddListContinuationParagraph(RtfDocument document, int parentLevel) {
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.MarkdownListContinuation = true;
        paragraph.LeftIndentTwips = (Math.Max(0, parentLevel) + 1) * ListIndentTwips;
        string bookmarkName = RtfMarkdownBridgeMarkers.CreateListContinuationBookmarkName(document.Paragraphs.Count);
        paragraph.AddBookmarkStart(bookmarkName);
        paragraph.AddBookmarkEnd(bookmarkName);
        return paragraph;
    }

    private static void ConvertTable(RtfDocument document, TableBlock table, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        int rowCount = table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0);
        int columnCount = Math.Max(table.Headers.Count, table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Count));
        if (rowCount == 0 || columnCount == 0) {
            context.Report("MDRTF002", RtfMarkdownDiagnosticSeverity.Info, "Empty Markdown table omitted from RTF output.", action: RtfConversionAction.Omitted);
            return;
        }

        EnsureTableWithinLimit(rowCount, columnCount, context);

        RtfTable rtfTable = document.AddTable(rowCount, columnCount);
        int rtfRowIndex = 0;
        if (table.Headers.Count > 0) {
            RtfTableRow headerRow = rtfTable.Rows[rtfRowIndex++];
            headerRow.RepeatHeader = true;
            FillTableRow(headerRow, table.HeaderInlines, table.Alignments, document, context, footnoteDefinitions);
        }

        IReadOnlyList<IReadOnlyList<InlineSequence>> rowInlines = table.RowInlines;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            IReadOnlyList<InlineSequence> cells = rowIndex < rowInlines.Count
                ? rowInlines[rowIndex]
                : Array.Empty<InlineSequence>();
            FillTableRow(rtfTable.Rows[rtfRowIndex++], cells, table.Alignments, document, context, footnoteDefinitions);
        }
    }

    private static void EnsureTableWithinLimit(TableBlock table, MarkdownToRtfConversionContext context) {
        int rowCount = table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0);
        int columnCount = Math.Max(table.Headers.Count, table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Count));
        EnsureTableWithinLimit(rowCount, columnCount, context);
    }

    private static void EnsureTableWithinLimit(int rowCount, int columnCount, MarkdownToRtfConversionContext context) {
        long tableCells = (long)rowCount * columnCount;
        if (tableCells > context.MaxTableCells) {
            throw new InvalidOperationException($"The Markdown table contains {tableCells} cells, exceeding the configured limit of {context.MaxTableCells}.");
        }
    }

    private static void FillTableRow(RtfTableRow row, IReadOnlyList<InlineSequence> cells, IReadOnlyList<ColumnAlignment> alignments, RtfDocument document, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        for (int column = 0; column < row.Cells.Count; column++) {
            RtfParagraph paragraph = row.Cells[column].AddParagraph();
            if (TryMapAlignment(alignments, column, out RtfTextAlignment alignment)) {
                paragraph.SetAlignment(alignment);
            }

            if (column < cells.Count) {
                AppendInlineSequence(paragraph, cells[column], document, context, InlineStyle.Normal, footnoteDefinitions);
            }
        }
    }

    private static bool TryMapAlignment(IReadOnlyList<ColumnAlignment> alignments, int column, out RtfTextAlignment alignment) {
        alignment = RtfTextAlignment.Left;
        if (alignments == null || column < 0 || column >= alignments.Count) {
            return false;
        }

        switch (alignments[column]) {
            case ColumnAlignment.Left:
                alignment = RtfTextAlignment.Left;
                return true;
            case ColumnAlignment.Center:
                alignment = RtfTextAlignment.Center;
                return true;
            case ColumnAlignment.Right:
                alignment = RtfTextAlignment.Right;
                return true;
            default:
                return false;
        }
    }

    private static void ConvertParagraph(RtfDocument document, ParagraphBlock paragraph, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        if (ShouldOmitUnsafeHtmlOnlyParagraph(paragraph.Inlines, context)) {
            return;
        }

        AppendInlineSequence(document.AddParagraph(), paragraph.Inlines, document, context, InlineStyle.Normal, footnoteDefinitions);
    }

    private static bool ShouldOmitUnsafeHtmlOnlyParagraph(InlineSequence sequence, MarkdownToRtfConversionContext context) {
        if (context.PreserveRawHtmlAsText || sequence.Nodes.Count != 1) {
            return false;
        }

        if (sequence.Nodes[0] is HtmlTagSequenceInline htmlTagSequence &&
            IsSupportedHtmlFormattingTag(htmlTagSequence.TagName) &&
            ContainsUnsupportedHtml(htmlTagSequence.Inlines)) {
            context.Report("MDRTF004", RtfMarkdownDiagnosticSeverity.Warning, "Markdown raw HTML block omitted. Set PreserveRawHtmlAsText to keep it as visible text.", htmlTagSequence.RenderMarkdown(), RtfConversionAction.Omitted);
            return true;
        }

        return false;
    }

    private static void ConvertImageBlock(RtfDocument document, ImageBlock image, MarkdownToRtfConversionContext context) {
        string label = string.IsNullOrWhiteSpace(image.PlainAlt) ? image.Path : image.PlainAlt!;
        document.AddParagraph("[Image: " + label + "]");
        context.Report("MDRTF003", RtfMarkdownDiagnosticSeverity.Warning, "Markdown image source represented as text placeholder; binary embedding requires caller-provided media bytes.", image.Path, RtfConversionAction.Flattened);
    }

    private static void ConvertCodeBlock(RtfDocument document, CodeBlock code) {
        int fontId = document.AddFont("Consolas");
        string[] lines = code.Content.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
        string bookmarkName = RtfMarkdownBridgeMarkers.CreateCodeBlockBookmarkName(document.Paragraphs.Count, code.InfoString);
        for (int i = 0; i < lines.Length; i++) {
            RtfParagraph paragraph = document.AddParagraph();
            paragraph.AddBookmarkStart(bookmarkName);
            paragraph.AddBookmarkEnd(bookmarkName);
            paragraph.AddText(lines[i]).FontId = fontId;
        }
    }

    private static void ConvertRawHtml(RtfDocument document, string html, MarkdownToRtfConversionContext context, string source, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        if (context.PreserveRawHtmlAsText) {
            document.AddParagraph(html);
            return;
        }

        if (TryConvertRawHtmlAsInlineFormatting(document, html, context, footnoteDefinitions)) {
            return;
        }

        context.Report("MDRTF004", RtfMarkdownDiagnosticSeverity.Warning, source + " omitted. Set PreserveRawHtmlAsText to keep it as visible text.", html, RtfConversionAction.Omitted);
    }

    private static void ConvertRawHtmlAsContinuation(RtfDocument document, string html, int parentLevel, MarkdownToRtfConversionContext context, string source) {
        if (context.PreserveRawHtmlAsText) {
            AddListContinuationParagraph(document, parentLevel).AddText(html);
        } else {
            context.Report("MDRTF014", RtfMarkdownDiagnosticSeverity.Warning, source + " inside list item omitted. Set PreserveRawHtmlAsText to keep it as visible text.", html, RtfConversionAction.Omitted);
        }
    }

    private static bool TryConvertRawHtmlAsInlineFormatting(RtfDocument document, string html, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        string trimmed = html.Trim();
        if (trimmed.Length == 0 ||
            trimmed.IndexOf('\r') >= 0 ||
            trimmed.IndexOf('\n') >= 0) {
            return false;
        }

        InlineSequence sequence = MarkdownReader.ParseInlineText(trimmed, context.ReaderOptions);
        if (sequence.Nodes.Count == 0 ||
            !ContainsSupportedHtmlTag(sequence) ||
            ContainsUnsupportedHtml(sequence)) {
            return false;
        }

        AppendInlineSequence(document.AddParagraph(), sequence, document, context, InlineStyle.Normal, footnoteDefinitions);
        return true;
    }

    private static bool ContainsSupportedHtmlTag(InlineSequence sequence) {
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            if (ContainsSupportedHtmlTag(sequence.Nodes[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsSupportedHtmlTag(IMarkdownInline inline) {
        if (inline is HtmlTagSequenceInline htmlTagSequence && IsSupportedHtmlFormattingTag(htmlTagSequence.TagName)) {
            return true;
        }

        InlineSequence? nested = (inline as IInlineContainerMarkdownInline)?.NestedInlines;
        return nested != null && ContainsSupportedHtmlTag(nested);
    }

    private static bool ContainsUnsupportedHtml(InlineSequence sequence) {
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            if (ContainsUnsupportedHtml(sequence.Nodes[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsUnsupportedHtml(IMarkdownInline inline) {
        if (inline is HtmlRawInline) {
            return true;
        }

        if (inline is HtmlTagSequenceInline htmlTagSequence && !IsSupportedHtmlFormattingTag(htmlTagSequence.TagName)) {
            return true;
        }

        if (inline is TextRun textRun && ContainsRawHtmlTagLikeText(textRun.Text)) {
            return true;
        }

        if (inline is DecodedHtmlEntityTextRun decodedTextRun && ContainsRawHtmlTagLikeText(decodedTextRun.Text)) {
            return true;
        }

        InlineSequence? nested = (inline as IInlineContainerMarkdownInline)?.NestedInlines;
        return nested != null && ContainsUnsupportedHtml(nested);
    }

    private static bool IsSupportedHtmlFormattingTag(string tagName) {
        return tagName == "u" || tagName == "sup" || tagName == "sub";
    }

    private static bool ContainsRawHtmlTagLikeText(string? text) {
        if (string.IsNullOrEmpty(text)) {
            return false;
        }

        string value = text!;
        for (int i = 0; i < value.Length - 2; i++) {
            if (value[i] != '<') {
                continue;
            }

            char next = value[i + 1] == '/' && i + 2 < value.Length
                ? value[i + 2]
                : value[i + 1];
            if (next == '!' || next == '?' || char.IsLetter(next)) {
                return true;
            }
        }

        return false;
    }

    private static void ConvertChildBlocks(RtfDocument document, IReadOnlyList<IMarkdownBlock> blocks, MarkdownToRtfConversionContext context, string message, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions) {
        context.Report("MDRTF005", RtfMarkdownDiagnosticSeverity.Info, message, action: RtfConversionAction.Flattened);
        for (int i = 0; i < blocks.Count; i++) {
            ConvertBlock(document, blocks[i], context, footnoteDefinitions);
        }
    }

    private static void AppendInlineSequence(RtfParagraph paragraph, InlineSequence sequence, RtfDocument document, MarkdownToRtfConversionContext context, InlineStyle style, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions, HashSet<string>? activeFootnotes = null, bool allowTextRunMerging = true) {
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            AppendInline(paragraph, sequence.Nodes[i], document, context, style, footnoteDefinitions, activeFootnotes, allowTextRunMerging);
        }
    }

    private static void AppendInline(RtfParagraph paragraph, IMarkdownInline inline, RtfDocument document, MarkdownToRtfConversionContext context, InlineStyle style, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions, HashSet<string>? activeFootnotes = null, bool allowTextRunMerging = true) {
        switch (inline) {
            case TextRun text:
                AddStyledText(paragraph, text.Text, style, allowTextRunMerging);
                break;
            case DecodedHtmlEntityTextRun decodedText:
                AddStyledTextRaw(paragraph, decodedText.Text, style, allowTextRunMerging);
                break;
            case BoldInline bold:
                AddStyledText(paragraph, bold.Text, style.WithBold(), allowTextRunMerging);
                break;
            case ItalicInline italic:
                AddStyledText(paragraph, italic.Text, style.WithItalic(), allowTextRunMerging);
                break;
            case BoldItalicInline boldItalic:
                AddStyledText(paragraph, boldItalic.Text, style.WithBold().WithItalic(), allowTextRunMerging);
                break;
            case StrikethroughInline strike:
                AddStyledText(paragraph, strike.Text, style.WithStrike(), allowTextRunMerging);
                break;
            case UnderlineInline underline:
                AddStyledText(paragraph, underline.Text, style.WithUnderline(), allowTextRunMerging);
                break;
            case HighlightInline highlight:
                AddStyledText(paragraph, highlight.Text, style.WithHighlight(EnsureHighlightColor(document)), allowTextRunMerging);
                break;
            case CodeSpanInline code:
                AddStyledTextRaw(paragraph, code.Text, style.WithFont(document.AddFont("Consolas")), allowTextRunMerging);
                break;
            case LinkInline link:
                AppendLink(paragraph, link, document, context, style, footnoteDefinitions, activeFootnotes);
                break;
            case FootnoteRefInline footnote:
                AppendFootnoteReference(paragraph, footnote, document, context, style, footnoteDefinitions, activeFootnotes);
                break;
            case ImageInline image:
                AddStyledText(paragraph, "[Image: " + image.PlainAlt + "]", style, allowTextRunMerging);
                context.Report("MDRTF006", RtfMarkdownDiagnosticSeverity.Warning, "Markdown inline image represented as text placeholder; binary embedding requires caller-provided media bytes.", image.Src, RtfConversionAction.Flattened);
                break;
            case HardBreakInline:
                paragraph.AddLineBreak();
                break;
            case SoftBreakInline:
                AddStyledText(paragraph, " ", style, allowTextRunMerging);
                break;
            case HtmlRawInline html:
                AppendInlineRawHtml(paragraph, html.Html, context, style);
                break;
            case BoldSequenceInline boldSequence:
                AppendInlineSequence(paragraph, boldSequence.Inlines, document, context, style.WithBold(), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case ItalicSequenceInline italicSequence:
                AppendInlineSequence(paragraph, italicSequence.Inlines, document, context, style.WithItalic(), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case BoldItalicSequenceInline boldItalicSequence:
                AppendInlineSequence(paragraph, boldItalicSequence.Inlines, document, context, style.WithBold().WithItalic(), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case StrikethroughSequenceInline strikeSequence:
                AppendInlineSequence(paragraph, strikeSequence.Inlines, document, context, style.WithStrike(), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case HighlightSequenceInline highlightSequence:
                AppendInlineSequence(paragraph, highlightSequence.Inlines, document, context, style.WithHighlight(EnsureHighlightColor(document)), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case InsertedSequenceInline insertedSequence:
                AppendInlineSequence(paragraph, insertedSequence.Inlines, document, context, style.WithUnderline(), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case SuperscriptSequenceInline superscriptSequence:
                AppendInlineSequence(paragraph, superscriptSequence.Inlines, document, context, style.WithVerticalPosition(RtfVerticalPosition.Superscript), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case SubscriptSequenceInline subscriptSequence:
                AppendInlineSequence(paragraph, subscriptSequence.Inlines, document, context, style.WithVerticalPosition(RtfVerticalPosition.Subscript), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case InsertedInline inserted:
                AddStyledText(paragraph, inserted.Text, style.WithUnderline(), allowTextRunMerging);
                break;
            case SuperscriptInline superscript:
                AddStyledText(paragraph, superscript.Text, style.WithVerticalPosition(RtfVerticalPosition.Superscript), allowTextRunMerging);
                break;
            case SubscriptInline subscript:
                AddStyledText(paragraph, subscript.Text, style.WithVerticalPosition(RtfVerticalPosition.Subscript), allowTextRunMerging);
                break;
            case HtmlTagSequenceInline htmlTagSequence:
                AppendHtmlTagSequence(paragraph, htmlTagSequence, document, context, style, footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                AppendInlineSequence(paragraph, container.NestedInlines!, document, context, style, footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            default:
                AddStyledText(paragraph, RtfMarkdownText.PlainText(inline), style, allowTextRunMerging);
                context.Report("MDRTF007", RtfMarkdownDiagnosticSeverity.Info, "Markdown inline converted using plain text fallback.", inline.GetType().Name, RtfConversionAction.Flattened);
                break;
        }
    }

    private static void AppendHtmlTagSequence(RtfParagraph paragraph, HtmlTagSequenceInline htmlTagSequence, RtfDocument document, MarkdownToRtfConversionContext context, InlineStyle style, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions, HashSet<string>? activeFootnotes = null, bool allowTextRunMerging = true) {
        if (!context.PreserveRawHtmlAsText &&
            IsSupportedHtmlFormattingTag(htmlTagSequence.TagName) &&
            ContainsUnsupportedHtml(htmlTagSequence.Inlines)) {
            context.Report("MDRTF004", RtfMarkdownDiagnosticSeverity.Warning, "Markdown raw HTML block omitted. Set PreserveRawHtmlAsText to keep it as visible text.", htmlTagSequence.RenderMarkdown(), RtfConversionAction.Omitted);
            return;
        }

        switch (htmlTagSequence.TagName) {
            case "u":
                AppendInlineSequence(paragraph, htmlTagSequence.Inlines, document, context, style.WithUnderline(), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case "sup":
                AppendInlineSequence(paragraph, htmlTagSequence.Inlines, document, context, style.WithVerticalPosition(RtfVerticalPosition.Superscript), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            case "sub":
                AppendInlineSequence(paragraph, htmlTagSequence.Inlines, document, context, style.WithVerticalPosition(RtfVerticalPosition.Subscript), footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                break;
            default:
                AppendInlineSequence(paragraph, htmlTagSequence.Inlines, document, context, style, footnoteDefinitions, activeFootnotes, allowTextRunMerging);
                context.Report("MDRTF011", RtfMarkdownDiagnosticSeverity.Info, "Markdown HTML inline tag converted using nested text fallback.", htmlTagSequence.TagName, RtfConversionAction.Flattened);
                break;
        }
    }

    private static void AppendLink(RtfParagraph paragraph, LinkInline link, RtfDocument document, MarkdownToRtfConversionContext context, InlineStyle style, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions, HashSet<string>? activeFootnotes = null) {
        Uri? uri = null;
        if (!Uri.TryCreate(link.Url, UriKind.RelativeOrAbsolute, out uri)) {
            context.Report("MDRTF009", RtfMarkdownDiagnosticSeverity.Warning, "Markdown link URL was not valid for RTF hyperlink metadata.", link.Url, RtfConversionAction.Flattened);
        }

        if (link.LabelInlines != null) {
            int before = paragraph.Inlines.Count;
            AppendInlineSequence(paragraph, link.LabelInlines, document, context, style, footnoteDefinitions, activeFootnotes, allowTextRunMerging: false);
            if (uri != null) {
                for (int i = before; i < paragraph.Inlines.Count; i++) {
                    if (paragraph.Inlines[i] is RtfRun hyperlinkRun) {
                        hyperlinkRun.SetHyperlink(uri);
                    }
                }
            }

            if (paragraph.Inlines.Count > before || uri == null) {
                return;
            }
        }

        RtfRun simpleRun = AddStyledText(paragraph, link.Text, style, allowMerge: false);
        if (uri != null) {
            simpleRun.SetHyperlink(uri);
        }
    }

    private static void AppendFootnoteReference(RtfParagraph paragraph, FootnoteRefInline footnote, RtfDocument document, MarkdownToRtfConversionContext context, InlineStyle style, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions, HashSet<string>? activeFootnotes = null) {
        if (!footnoteDefinitions.TryGetValue(footnote.Label, out FootnoteDefinitionBlock? definition)) {
            AddStyledText(paragraph, footnote.RenderMarkdown(), style);
            context.Report("MDRTF018", RtfMarkdownDiagnosticSeverity.Warning, "Markdown footnote reference has no matching definition.", footnote.Label, RtfConversionAction.Omitted);
            return;
        }

        activeFootnotes ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!activeFootnotes.Add(footnote.Label)) {
            AddStyledText(paragraph, footnote.RenderMarkdown(), style);
            context.Report("MDRTF020", RtfMarkdownDiagnosticSeverity.Warning, "Markdown footnote reference cycle preserved as literal text.", footnote.Label, RtfConversionAction.Flattened);
            return;
        }

        RtfNote note = document.AddNote(RtfNoteKind.Footnote);
        try {
            AddFootnoteDefinitionContent(note, definition, document, context, footnoteDefinitions, activeFootnotes);
            paragraph.AddNoteReference(note, footnote.Label);
        } finally {
            activeFootnotes.Remove(footnote.Label);
        }
    }

    private static void AddFootnoteDefinitionContent(RtfNote note, FootnoteDefinitionBlock definition, RtfDocument document, MarkdownToRtfConversionContext context, IReadOnlyDictionary<string, FootnoteDefinitionBlock> footnoteDefinitions, HashSet<string> activeFootnotes) {
        IReadOnlyList<IMarkdownBlock> blocks = definition.ChildBlocks.Count > 0
            ? definition.ChildBlocks
            : new IMarkdownBlock[] { new ParagraphBlock(MarkdownReader.ParseInlineText(definition.Text, context.ReaderOptions)) };

        for (int i = 0; i < blocks.Count; i++) {
            switch (blocks[i]) {
                case ParagraphBlock paragraphBlock:
                    AppendInlineSequence(note.AddParagraph(), paragraphBlock.Inlines, document, context, InlineStyle.Normal, footnoteDefinitions, activeFootnotes);
                    break;
                default:
                    note.AddParagraph(blocks[i].RenderMarkdown());
                    context.Report("MDRTF019", RtfMarkdownDiagnosticSeverity.Info, "Markdown footnote child block converted to rendered Markdown text.", blocks[i].GetType().Name, RtfConversionAction.Flattened);
                    break;
            }
        }
    }

    private static void AppendInlineRawHtml(RtfParagraph paragraph, string html, MarkdownToRtfConversionContext context, InlineStyle style) {
        if (context.PreserveRawHtmlAsText) {
            AddStyledText(paragraph, html, style);
        } else {
            context.Report("MDRTF010", RtfMarkdownDiagnosticSeverity.Warning, "Markdown raw inline HTML omitted. Set PreserveRawHtmlAsText to keep it as visible text.", html, RtfConversionAction.Omitted);
        }
    }

    private static RtfRun AddStyledText(RtfParagraph paragraph, string text, InlineStyle style, bool allowMerge = true) {
        return AddStyledTextRaw(paragraph, DecodeMarkdownVisibleText(text), style, allowMerge);
    }

    private static RtfRun AddStyledTextRaw(RtfParagraph paragraph, string text, InlineStyle style, bool allowMerge = true) {
        string value = text ?? string.Empty;
        if (allowMerge &&
            paragraph.Inlines.Count > 0 &&
            paragraph.Inlines[paragraph.Inlines.Count - 1] is RtfRun previous &&
            CanMergeRun(previous, style)) {
            previous.Text += value;
            return previous;
        }

        RtfRun run = paragraph.AddText(value);
        ApplyStyle(run, style);
        return run;
    }

    private static bool CanMergeRun(RtfRun run, InlineStyle style) {
        return run.Hyperlink == null &&
            run.Note == null &&
            run.Bold == style.Bold &&
            run.Italic == style.Italic &&
            run.Strike == style.Strike &&
            run.UnderlineStyle == (style.Underline ? RtfUnderlineStyle.Single : RtfUnderlineStyle.None) &&
            run.HighlightColorIndex == style.HighlightColorIndex &&
            run.FontId == style.FontId &&
            run.VerticalPosition == (style.VerticalPosition ?? RtfVerticalPosition.Baseline);
    }

    private static void ApplyStyle(RtfRun run, InlineStyle style) {
        if (style.Bold) run.SetBold();
        if (style.Italic) run.SetItalic();
        if (style.Strike) run.SetStrike();
        if (style.Underline) run.SetUnderline(RtfUnderlineStyle.Single);
        if (style.HighlightColorIndex.HasValue) run.SetHighlightColor(style.HighlightColorIndex.Value);
        if (style.FontId.HasValue) run.FontId = style.FontId.Value;
        if (style.VerticalPosition.HasValue) run.VerticalPosition = style.VerticalPosition.Value;
    }

    private static string DecodeMarkdownVisibleText(string? text) {
        return System.Net.WebUtility.HtmlDecode(text ?? string.Empty);
    }

    private static int EnsureHighlightColor(RtfDocument document) {
        return document.AddColor(255, 255, 0);
    }

    private readonly struct InlineStyle {
        internal static readonly InlineStyle Normal = new InlineStyle(false, false, false, false, null, null, null);

        private InlineStyle(bool bold, bool italic, bool strike, bool underline, int? highlightColorIndex, int? fontId, RtfVerticalPosition? verticalPosition) {
            Bold = bold;
            Italic = italic;
            Strike = strike;
            Underline = underline;
            HighlightColorIndex = highlightColorIndex;
            FontId = fontId;
            VerticalPosition = verticalPosition;
        }

        internal bool Bold { get; }

        internal bool Italic { get; }

        internal bool Strike { get; }

        internal bool Underline { get; }

        internal int? HighlightColorIndex { get; }

        internal int? FontId { get; }

        internal RtfVerticalPosition? VerticalPosition { get; }

        internal InlineStyle WithBold() => new InlineStyle(true, Italic, Strike, Underline, HighlightColorIndex, FontId, VerticalPosition);

        internal InlineStyle WithItalic() => new InlineStyle(Bold, true, Strike, Underline, HighlightColorIndex, FontId, VerticalPosition);

        internal InlineStyle WithStrike() => new InlineStyle(Bold, Italic, true, Underline, HighlightColorIndex, FontId, VerticalPosition);

        internal InlineStyle WithUnderline() => new InlineStyle(Bold, Italic, Strike, true, HighlightColorIndex, FontId, VerticalPosition);

        internal InlineStyle WithHighlight(int colorIndex) => new InlineStyle(Bold, Italic, Strike, Underline, colorIndex, FontId, VerticalPosition);

        internal InlineStyle WithFont(int fontId) => new InlineStyle(Bold, Italic, Strike, Underline, HighlightColorIndex, fontId, VerticalPosition);

        internal InlineStyle WithVerticalPosition(RtfVerticalPosition verticalPosition) => new InlineStyle(Bold, Italic, Strike, Underline, HighlightColorIndex, FontId, verticalPosition);
    }
}
