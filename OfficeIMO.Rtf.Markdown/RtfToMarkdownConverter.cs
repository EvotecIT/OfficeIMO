using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Markdown;

internal static class RtfToMarkdownConverter {
    private const int ListIndentTwips = 720;

    internal static MarkdownDoc Convert(RtfDocument document, RtfToMarkdownConversionContext context) {
        RtfTableTraversalGuard.ValidateDocument(document);
        var blocks = new List<IMarkdownBlock>(document.Blocks.Count + document.Notes.Count + 1);
        int imageIndex = 0;
        var listStartLookup = new ListStartLookup(document);
        context.NoteRegistry = new RtfMarkdownNoteRegistry();

        try {
            for (int i = 0; i < document.Blocks.Count; i++) {
                IRtfBlock block = document.Blocks[i];
                switch (block) {
                    case RtfParagraph paragraph:
                        if (paragraph.ListKind != RtfListKind.None) {
                            i = ConvertListRun(document, listStartLookup, context, blocks, i, ref imageIndex);
                        } else if (TryConvertCodeBlockRun(document, blocks, i, out int codeBlockEndIndex)) {
                            i = codeBlockEndIndex;
                        } else {
                            blocks.Add(ConvertParagraph(document, paragraph, context, ref imageIndex));
                        }
                        break;
                    case RtfTable table:
                        blocks.Add(ConvertTable(document, table, context, ref imageIndex));
                        break;
                    case RtfImage image:
                        blocks.Add(ConvertImageBlock(image, context, ref imageIndex));
                        break;
                    case RtfObject:
                        AddUnsupportedBlock(blocks, context, "RTF object block omitted.", "rtf-object");
                        break;
                    case RtfShape:
                        AddUnsupportedBlock(blocks, context, "RTF drawing shape block omitted.", "rtf-shape");
                        break;
                    default:
                        context.Report("RTFMD001", RtfMarkdownDiagnosticSeverity.Warning, "Unsupported RTF block omitted.", block.GetType().Name, RtfConversionAction.Omitted);
                        break;
                }
            }

            ReportOmittedHeaderFooters(document, blocks, context);
            AppendNoteDefinitions(document, blocks, context, ref imageIndex);
            return MarkdownDoc.Create().AddRange(blocks);
        } finally {
            context.NoteRegistry = null;
        }
    }

    private static int ConvertListRun(RtfDocument document, ListStartLookup listStartLookup, RtfToMarkdownConversionContext context, ICollection<IMarkdownBlock> blocks, int startIndex, ref int imageIndex) {
        var first = (RtfParagraph)document.Blocks[startIndex];
        int? firstListId = first.ListId;
        int? firstListDefinitionId = first.ListDefinitionId;
        var paragraphs = new List<RtfParagraph>();
        int i = startIndex;

        for (; i < document.Blocks.Count; i++) {
            if (!(document.Blocks[i] is RtfParagraph paragraph)) {
                break;
            }

            if (paragraph.ListKind == RtfListKind.None) {
                if (paragraphs.Count == 0 || !IsListContinuationParagraph(paragraph)) {
                    break;
                }

                paragraphs.Add(paragraph);
                continue;
            }

            int level = Math.Max(0, paragraph.ListLevel ?? 0);
            if (level == 0 && (paragraph.ListId != firstListId ||
                paragraph.ListDefinitionId != firstListDefinitionId)) {
                break;
            }

            if (level == 0 && paragraph.ListKind != first.ListKind) {
                break;
            }

            paragraphs.Add(paragraph);
        }

        blocks.Add(ConvertListParagraphs(listStartLookup, context, paragraphs, ref imageIndex));
        return i - 1;
    }

    private static IMarkdownBlock ConvertListParagraphs(ListStartLookup listStartLookup, RtfToMarkdownConversionContext context, IReadOnlyList<RtfParagraph> paragraphs, ref int imageIndex) {
        RtfParagraph first = paragraphs[0];
        IMarkdownListBlock root = CreateMarkdownListBlock(listStartLookup, first);
        var frames = new List<ListFrame> {
            CreateListFrame(listStartLookup, first, 0, NormalizeListKind(first.ListKind), root)
        };

        for (int i = 0; i < paragraphs.Count; i++) {
            RtfParagraph paragraph = paragraphs[i];
            if (paragraph.ListKind == RtfListKind.None) {
                AddListContinuationParagraph(frames, paragraph, context, ref imageIndex);
                continue;
            }

            int level = Math.Max(0, paragraph.ListLevel ?? 0);
            RtfListKind kind = NormalizeListKind(paragraph.ListKind);
            InlineSequence inlines = ConvertParagraphInlines(paragraph, context, ref imageIndex);
            ListItem item = CreateListItem(inlines);

            ListFrame frame = GetOrCreateListFrame(listStartLookup, frames, paragraph, level, kind);
            AddListItem(frame.List, item);
            frame.LastItem = item;
        }

        return (IMarkdownBlock)root;
    }

    private static bool IsListContinuationParagraph(RtfParagraph paragraph) {
        return paragraph.ListKind == RtfListKind.None &&
               (paragraph.MarkdownListContinuation || HasListContinuationBookmark(paragraph));
    }

    private static bool HasListContinuationBookmark(RtfParagraph paragraph) {
        for (int i = 0; i < paragraph.Inlines.Count; i++) {
            if (paragraph.Inlines[i] is RtfBookmarkMarker marker &&
                RtfMarkdownBridgeMarkers.IsListContinuationBookmark(marker)) {
                return true;
            }
        }

        return false;
    }

    private static void AddListContinuationParagraph(List<ListFrame> frames, RtfParagraph paragraph, RtfToMarkdownConversionContext context, ref int imageIndex) {
        ListFrame? frame = FindContinuationFrame(frames, ResolveContinuationLevel(paragraph));
        if (frame?.LastItem == null) {
            return;
        }

        frame.LastItem.AdditionalParagraphs.Add(ConvertParagraphInlines(paragraph, context, ref imageIndex));
    }

    private static ListFrame? FindContinuationFrame(List<ListFrame> frames, int level) {
        for (int i = frames.Count - 1; i >= 0; i--) {
            if (frames[i].Level <= level && frames[i].LastItem != null) {
                return frames[i];
            }
        }

        return frames.Count > 0 ? frames[frames.Count - 1] : null;
    }

    private static int ResolveContinuationLevel(RtfParagraph paragraph) {
        int leftIndent = Math.Max(0, paragraph.LeftIndentTwips ?? 0);
        return Math.Max(0, (leftIndent / ListIndentTwips) - 1);
    }

    private static ListItem CreateListItem(InlineSequence inlines) {
        return TryStripTaskMarker(inlines, out InlineSequence content, out bool isChecked)
            ? ListItem.TaskInlines(content, isChecked)
            : new ListItem(inlines);
    }

    private static bool TryStripTaskMarker(InlineSequence inlines, out InlineSequence content, out bool isChecked) {
        content = inlines;
        isChecked = false;

        if (inlines.Nodes.Count == 0) {
            return false;
        }

        IMarkdownInline first = inlines.Nodes[0];
        string? text = first switch {
            TextRun run => run.Text,
            DecodedHtmlEntityTextRun run => run.Text,
            _ => null
        };

        if (string.IsNullOrEmpty(text) || text!.Length < 3 || text[0] != '[' || text[2] != ']') {
            return false;
        }

        if (text[1] == ' ') {
            isChecked = false;
        } else if (text[1] == 'x' || text[1] == 'X') {
            isChecked = true;
        } else {
            return false;
        }

        if (text.Length <= 3 || !char.IsWhiteSpace(text[3])) {
            return false;
        }

        int consume = 4;
        content = CreateInlineSequence();
        string remaining = text.Substring(consume);
        if (remaining.Length > 0) {
            content.AddRaw(first is TextRun ? new TextRun(remaining) : new DecodedHtmlEntityTextRun(remaining));
        }

        for (int i = 1; i < inlines.Nodes.Count; i++) {
            content.AddRaw(inlines.Nodes[i]);
        }

        return true;
    }

    private static ListFrame GetOrCreateListFrame(ListStartLookup listStartLookup, List<ListFrame> frames, RtfParagraph paragraph, int level, RtfListKind kind) {
        if (level <= 0) {
            while (frames.Count > 1) {
                frames.RemoveAt(frames.Count - 1);
            }

            return frames[0];
        }

        while (frames.Count > 0) {
            ListFrame current = frames[frames.Count - 1];
            if (current.Level < level || MatchesListFrame(listStartLookup, current, paragraph, level, kind)) {
                break;
            }

            frames.RemoveAt(frames.Count - 1);
        }

        ListFrame last = frames[frames.Count - 1];
        if (MatchesListFrame(listStartLookup, last, paragraph, level, kind)) {
            return last;
        }

        if (last.LastItem == null) {
            return frames[0];
        }

        IMarkdownListBlock childList = CreateMarkdownListBlock(listStartLookup, paragraph);
        last.LastItem.NestedBlocks.Add((IMarkdownBlock)childList);
        var childFrame = CreateListFrame(listStartLookup, paragraph, level, kind, childList);
        frames.Add(childFrame);
        return childFrame;
    }

    private static ListFrame CreateListFrame(ListStartLookup listStartLookup, RtfParagraph paragraph, int level, RtfListKind kind, IMarkdownListBlock list) =>
        new ListFrame(
            level,
            kind,
            list,
            paragraph.ListId,
            paragraph.ListDefinitionId,
            kind == RtfListKind.Decimal ? ResolveListStart(listStartLookup, paragraph) : 1);

    private static bool MatchesListFrame(ListStartLookup listStartLookup, ListFrame frame, RtfParagraph paragraph, int level, RtfListKind kind) =>
        frame.Level == level &&
        frame.Kind == kind &&
        frame.ListId == paragraph.ListId &&
        frame.ListDefinitionId == paragraph.ListDefinitionId &&
        (kind != RtfListKind.Decimal || frame.Start == ResolveListStart(listStartLookup, paragraph));

    private static IMarkdownListBlock CreateMarkdownListBlock(ListStartLookup listStartLookup, RtfParagraph paragraph) {
        RtfListKind kind = NormalizeListKind(paragraph.ListKind);
        return kind == RtfListKind.Decimal
            ? new OrderedListBlock { Start = ResolveListStart(listStartLookup, paragraph) }
            : new UnorderedListBlock();
    }

    private static void AddListItem(IMarkdownListBlock list, ListItem item) {
        if (list is OrderedListBlock orderedList) {
            orderedList.Items.Add(item);
        } else {
            ((UnorderedListBlock)list).Items.Add(item);
        }
    }

    private static RtfListKind NormalizeListKind(RtfListKind kind) {
        return kind == RtfListKind.Decimal ? RtfListKind.Decimal : RtfListKind.Bullet;
    }

    private static int ResolveListStart(ListStartLookup listStartLookup, RtfParagraph paragraph) {
        int levelIndex = Math.Max(0, paragraph.ListLevel ?? 0);
        if (paragraph.ListId.HasValue) {
            listStartLookup.Overrides.TryGetValue(paragraph.ListId.Value, out RtfListOverride? listOverride);
            RtfListLevelOverride? levelOverride = listOverride?.LevelOverrides.ElementAtOrDefault(levelIndex);
            if (levelOverride?.OverrideStartAt == true && levelOverride.StartAt.HasValue) {
                return Math.Max(1, levelOverride.StartAt.Value);
            }
        }

        if (paragraph.ListDefinitionId.HasValue) {
            listStartLookup.Definitions.TryGetValue(paragraph.ListDefinitionId.Value, out RtfListDefinition? definition);
            RtfListLevel? level = definition?.Levels.FirstOrDefault(item => item.LevelIndex == levelIndex);
            if (level?.StartAt.HasValue == true) {
                return Math.Max(1, level.StartAt.Value);
            }
        }

        return Math.Max(1, paragraph.LegacyNumbering.StartAt ?? 1);
    }

    private sealed class ListStartLookup {
        internal Dictionary<int, RtfListOverride> Overrides { get; } = new();
        internal Dictionary<int, RtfListDefinition> Definitions { get; } = new();

        internal ListStartLookup(RtfDocument document) {
            foreach (RtfListOverride item in document.ListOverrides) {
                if (!Overrides.ContainsKey(item.Id)) {
                    Overrides.Add(item.Id, item);
                }
            }

            foreach (RtfListDefinition item in document.ListDefinitions) {
                if (!Definitions.ContainsKey(item.Id)) {
                    Definitions.Add(item.Id, item);
                }
            }
        }
    }

    private static IMarkdownBlock ConvertParagraph(RtfDocument document, RtfParagraph paragraph, RtfToMarkdownConversionContext context, ref int imageIndex) {
        InlineSequence inlines = ConvertParagraphInlines(paragraph, context, ref imageIndex);
        int? headingLevel = DetectHeadingLevel(document, paragraph);

        return headingLevel.HasValue
            ? new HeadingBlock(headingLevel.Value, inlines)
            : new ParagraphBlock(inlines);
    }

    private static int? DetectHeadingLevel(RtfDocument document, RtfParagraph paragraph) {
        if (paragraph.OutlineLevel.HasValue) {
            int level = paragraph.OutlineLevel.Value + 1;
            if (level < 1) return 1;
            if (level > 6) return 6;
            return level;
        }

        if (!paragraph.StyleId.HasValue) {
            return null;
        }

        RtfStyle? style = document.Styles.FirstOrDefault(item => item.Id == paragraph.StyleId.Value);
        if (style == null) {
            return null;
        }

        if (style.OutlineLevel.HasValue) {
            int level = style.OutlineLevel.Value + 1;
            if (level < 1) return 1;
            if (level > 6) return 6;
            return level;
        }

        string normalized = (style.Name ?? string.Empty).Replace(" ", string.Empty).Replace("-", string.Empty);
        if (normalized.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            && normalized.Length > "Heading".Length
            && int.TryParse(normalized.Substring("Heading".Length, 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out int heading)
            && heading >= 1
            && heading <= 6) {
            return heading;
        }

        return null;
    }

    private static TableBlock ConvertTable(RtfDocument document, RtfTable table, RtfToMarkdownConversionContext context, ref int imageIndex) {
        var markdown = new TableBlock {
            CellsContainRenderedMarkdown = true
        };
        if (table.Rows.Count == 0) {
            context.Report("RTFMD002", RtfMarkdownDiagnosticSeverity.Info, "Empty RTF table converted to an empty Markdown table.");
            return markdown;
        }

        bool hasHeader = table.Rows[0].RepeatHeader;
        int firstBodyRow = hasHeader ? 1 : 0;
        List<InlineSequence>? headerInlines = null;
        if (hasHeader) {
            RtfTableRow firstRow = table.Rows[0];
            headerInlines = new List<InlineSequence>(firstRow.Cells.Count);
            for (int column = 0; column < firstRow.Cells.Count; column++) {
                CellContent content = ConvertCellContent(firstRow.Cells[column], context, ref imageIndex);
                markdown.Headers.Add(content.Markdown);
                headerInlines.Add(content.Inlines);
            }
        }

        var rowInlines = new List<IReadOnlyList<InlineSequence>>();
        for (int rowIndex = firstBodyRow; rowIndex < table.Rows.Count; rowIndex++) {
            var row = table.Rows[rowIndex];
            var cells = new List<string>(row.Cells.Count);
            var inlines = new List<InlineSequence>(row.Cells.Count);
            for (int column = 0; column < row.Cells.Count; column++) {
                CellContent content = ConvertCellContent(row.Cells[column], context, ref imageIndex);
                cells.Add(content.Markdown);
                inlines.Add(content.Inlines);
            }

            markdown.Rows.Add(cells);
            rowInlines.Add(inlines);
        }

        ApplyTableColumnAlignments(markdown, table);
        markdown.SetParsedCells(headerInlines, rowInlines, markdown.ComputeContentSignature());
        if (!hasHeader && table.Rows.Count == 1) {
            markdown.PreserveHeaderlessSingleRowTable = true;
        }

        return markdown;
    }

    private static bool TryConvertCodeBlockRun(RtfDocument document, ICollection<IMarkdownBlock> blocks, int startIndex, out int endIndex) {
        endIndex = startIndex;
        if (!(document.Blocks[startIndex] is RtfParagraph first) ||
            !TryGetCodeBlockMarker(first, out string key, out string infoString)) {
            return false;
        }

        var lines = new List<string>();
        int index = startIndex;
        for (; index < document.Blocks.Count; index++) {
            if (!(document.Blocks[index] is RtfParagraph paragraph) ||
                !TryGetCodeBlockMarker(paragraph, out string currentKey, out string currentLanguage) ||
                !string.Equals(currentKey, key, StringComparison.Ordinal)) {
                break;
            }

            if (lines.Count == 0 && string.IsNullOrEmpty(infoString)) {
                infoString = currentLanguage;
            }

            lines.Add(paragraph.ToPlainText());
        }

        blocks.Add(new CodeBlock(infoString, string.Join("\n", lines)));
        endIndex = index - 1;
        return true;
    }

    private static bool TryGetCodeBlockMarker(RtfParagraph paragraph, out string key, out string language) {
        key = string.Empty;
        language = string.Empty;
        for (int i = 0; i < paragraph.Inlines.Count; i++) {
            if (paragraph.Inlines[i] is RtfBookmarkMarker marker &&
                RtfMarkdownBridgeMarkers.TryGetCodeBlockBookmark(marker, out key, out language)) {
                return true;
            }
        }

        return false;
    }

    private static void ApplyTableColumnAlignments(TableBlock markdown, RtfTable table) {
        int columnCount = table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Cells.Count);
        if (columnCount == 0) {
            return;
        }

        var alignments = new ColumnAlignment[columnCount];
        bool hasRepresentableAlignment = false;
        for (int column = 0; column < columnCount; column++) {
            ColumnAlignment alignment = ResolveColumnAlignment(table, column);
            alignments[column] = alignment;
            hasRepresentableAlignment |= alignment == ColumnAlignment.Center || alignment == ColumnAlignment.Right;
        }

        if (hasRepresentableAlignment) {
            markdown.Alignments.Clear();
            markdown.Alignments.AddRange(alignments);
        }
    }

    private static ColumnAlignment ResolveColumnAlignment(RtfTable table, int column) {
        ColumnAlignment? candidate = null;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            RtfTableRow row = table.Rows[rowIndex];
            if (column >= row.Cells.Count) {
                continue;
            }

            ColumnAlignment cellAlignment = ResolveCellAlignment(row.Cells[column]);
            if (cellAlignment == ColumnAlignment.None) {
                continue;
            }

            if (!candidate.HasValue) {
                candidate = cellAlignment;
                continue;
            }

            if (candidate.Value != cellAlignment) {
                return ColumnAlignment.None;
            }
        }

        return candidate == ColumnAlignment.Center || candidate == ColumnAlignment.Right
            ? candidate.Value
            : ColumnAlignment.None;
    }

    private static ColumnAlignment ResolveCellAlignment(RtfTableCell cell) {
        ColumnAlignment? candidate = null;
        for (int i = 0; i < cell.Paragraphs.Count; i++) {
            RtfParagraph paragraph = cell.Paragraphs[i];
            if (string.IsNullOrWhiteSpace(paragraph.ToPlainText())) {
                continue;
            }

            ColumnAlignment alignment = paragraph.Alignment switch {
                RtfTextAlignment.Center => ColumnAlignment.Center,
                RtfTextAlignment.Right => ColumnAlignment.Right,
                RtfTextAlignment.Left => ColumnAlignment.Left,
                _ => ColumnAlignment.None
            };

            if (alignment == ColumnAlignment.None) {
                return ColumnAlignment.None;
            }

            if (!candidate.HasValue) {
                candidate = alignment;
                continue;
            }

            if (candidate.Value != alignment) {
                return ColumnAlignment.None;
            }
        }

        return candidate ?? ColumnAlignment.None;
    }

    private static CellContent ConvertCellContent(RtfTableCell cell, RtfToMarkdownConversionContext context, ref int imageIndex) {
        var parts = new List<string>();
        InlineSequence combined = CreateInlineSequence();
        bool hasCombinedContent = false;
        foreach (IRtfBlock block in cell.Blocks) {
            if (block is RtfTable nestedTable) {
                string nestedText = FlattenNestedTable(nestedTable);
                if (!string.IsNullOrWhiteSpace(nestedText)) {
                    parts.Add(RtfMarkdownText.EscapeMarkdownText(nestedText));
                    if (hasCombinedContent) combined.AddRaw(new HardBreakInline());
                    combined.AddRaw(new DecodedHtmlEntityTextRun(nestedText));
                    hasCombinedContent = true;
                }

                context.Report("RTFMD016", RtfMarkdownDiagnosticSeverity.Warning, "Nested RTF table flattened to text inside a Markdown table cell.", "nested-table", RtfConversionAction.Flattened);
                continue;
            }

            if (!(block is RtfParagraph paragraph)) continue;
            InlineSequence paragraphInlines = ConvertParagraphInlines(paragraph, context, ref imageIndex);
            string text = RenderInlineSequenceMarkdownForTableCell(paragraphInlines);
            if (!string.IsNullOrEmpty(text)) {
                parts.Add(text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "<br>"));
            }

            if (paragraphInlines.Nodes.Count == 0) {
                continue;
            }

            if (hasCombinedContent) {
                combined.AddRaw(new HardBreakInline());
            }

            for (int nodeIndex = 0; nodeIndex < paragraphInlines.Nodes.Count; nodeIndex++) {
                combined.AddRaw(paragraphInlines.Nodes[nodeIndex]);
            }

            hasCombinedContent = true;
        }

        return new CellContent(string.Join("<br>", parts), combined);
    }

    private static string FlattenNestedTable(RtfTable table) {
        return string.Join(" / ", table.Rows.Select(row =>
            string.Join(" | ", row.Cells.Select(cell =>
                string.Join(" ", cell.Blocks.Select(block => block is RtfParagraph paragraph
                    ? paragraph.ToPlainText()
                    : block is RtfTable nested ? FlattenNestedTable(nested) : string.Empty)
                    .Where(text => !string.IsNullOrWhiteSpace(text)))))));
    }

    private static ImageBlock ConvertImageBlock(RtfImage image, RtfToMarkdownConversionContext context, ref int imageIndex) {
        int currentIndex = imageIndex++;
        string path = ResolveImagePath(image, currentIndex, context);
        string alt = string.IsNullOrWhiteSpace(image.Description) ? "RTF image" : image.Description!;
        context.Report("RTFMD003", RtfMarkdownDiagnosticSeverity.Info, "RTF image payload represented by Markdown image reference.", path, RtfConversionAction.Flattened);
        return new ImageBlock(path, alt, null);
    }

    private static InlineSequence ConvertParagraphInlines(RtfParagraph paragraph, RtfToMarkdownConversionContext context, ref int imageIndex) {
        InlineSequence sequence = CreateInlineSequence();
        for (int i = 0; i < paragraph.Inlines.Count; i++) {
            AppendInline(sequence, paragraph.Inlines[i], context, ref imageIndex);
        }

        return sequence;
    }

    private static InlineSequence CreateInlineSequence() {
        return new InlineSequence { AutoSpacing = false };
    }

    private static string RenderInlineSequenceMarkdown(InlineSequence sequence) {
        return ((IRenderableMarkdownInline)sequence).RenderMarkdown();
    }

    private static string RenderInlineSequenceMarkdownForTableCell(InlineSequence sequence) {
        if (sequence.Nodes.Count == 0) {
            return string.Empty;
        }

        var builder = new StringBuilder();
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            builder.Append(RenderInlineMarkdownForTableCell(sequence.Nodes[i]));
        }

        return builder.ToString();
    }

    private static string RenderInlineMarkdownForTableCell(IMarkdownInline inline) {
        switch (inline) {
            case DecodedHtmlEntityTextRun text:
                return MarkdownEscaper.EscapeLiteralTableCellText(text.Text);
            case BoldSequenceInline bold:
                return "**" + RenderInlineSequenceMarkdownForTableCell(bold.Inlines) + "**";
            case ItalicSequenceInline italic:
                return "*" + RenderInlineSequenceMarkdownForTableCell(italic.Inlines) + "*";
            case BoldItalicSequenceInline boldItalic:
                return "***" + RenderInlineSequenceMarkdownForTableCell(boldItalic.Inlines) + "***";
            case StrikethroughSequenceInline strike:
                return "~~" + RenderInlineSequenceMarkdownForTableCell(strike.Inlines) + "~~";
            case HighlightSequenceInline highlight:
                return "==" + RenderInlineSequenceMarkdownForTableCell(highlight.Inlines) + "==";
            case HtmlTagSequenceInline htmlTag:
                return "<" + htmlTag.TagName + ">" + RenderInlineSequenceMarkdownForTableCell(htmlTag.Inlines) + "</" + htmlTag.TagName + ">";
            case LinkInline link when link.LabelInlines != null:
                return "[" + RenderInlineSequenceMarkdownForTableCell(link.LabelInlines) + "](" + MarkdownEscaper.EscapeLinkUrl(link.Url) + MarkdownEscaper.FormatOptionalTitle(link.Title) + ")";
            case IRenderableMarkdownInline renderable:
                return renderable.RenderMarkdown();
            default:
                return RtfMarkdownText.EscapeMarkdownText(RtfMarkdownText.PlainText(inline));
        }
    }

    private static InlineSequence InlineSequenceOf(IMarkdownInline inline) {
        InlineSequence sequence = CreateInlineSequence();
        sequence.AddRaw(inline);
        return sequence;
    }

    private static void AppendInline(InlineSequence sequence, IRtfInline inline, RtfToMarkdownConversionContext context, ref int imageIndex) {
        switch (inline) {
            case RtfRun run:
                AppendRun(sequence, run, context);
                break;
            case RtfBreak rtfBreak:
                AppendBreak(sequence, rtfBreak, context);
                break;
            case RtfField field:
                AppendField(sequence, field, context, ref imageIndex);
                break;
            case RtfGeneratedText generatedText:
                AppendGeneratedText(sequence, generatedText, context);
                break;
            case RtfImage image:
                AppendImageInline(sequence, image, context, ref imageIndex);
                break;
            case RtfObject:
                AppendUnsupportedInline(sequence, context, "RTF object inline omitted.", "rtf-object");
                break;
            case RtfShape:
                AppendUnsupportedInline(sequence, context, "RTF drawing shape inline omitted.", "rtf-shape");
                break;
            case RtfBookmarkMarker:
                break;
            default:
                context.Report("RTFMD004", RtfMarkdownDiagnosticSeverity.Warning, "Unsupported RTF inline omitted.", inline.GetType().Name, RtfConversionAction.Omitted);
                break;
        }
    }

    private static void AppendRun(InlineSequence sequence, RtfRun run, RtfToMarkdownConversionContext context) {
        if (run.Hidden && !context.IncludeHiddenText) {
            context.Report("RTFMD005", RtfMarkdownDiagnosticSeverity.Info, "Hidden RTF text omitted from Markdown output.", action: RtfConversionAction.Omitted);
            return;
        }

        IMarkdownInline? inline = BuildRunInline(run);
        if (inline != null && run.Hyperlink != null) {
            sequence.AddRaw(new LinkInline(InlineSequenceOf(inline), FormatMarkdownLinkDestination(run.Hyperlink.ToString()), null));
        } else if (inline != null) {
            sequence.AddRaw(inline);
        }

        AppendNoteReference(sequence, run.Note, context);
    }

    private static IMarkdownInline? BuildRunInline(RtfRun run) {
        if (string.IsNullOrEmpty(run.Text)) {
            return null;
        }

        IMarkdownInline inline = new DecodedHtmlEntityTextRun(run.Text);
        if (run.VerticalPosition == RtfVerticalPosition.Superscript) {
            inline = new HtmlTagSequenceInline("sup", InlineSequenceOf(inline));
        } else if (run.VerticalPosition == RtfVerticalPosition.Subscript) {
            inline = new HtmlTagSequenceInline("sub", InlineSequenceOf(inline));
        }

        if (run.UnderlineStyle != RtfUnderlineStyle.None) {
            inline = new HtmlTagSequenceInline("u", InlineSequenceOf(inline));
        }

        if (run.Bold && run.Italic) {
            inline = new BoldItalicSequenceInline(InlineSequenceOf(inline));
        } else if (run.Bold) {
            inline = new BoldSequenceInline(InlineSequenceOf(inline));
        } else if (run.Italic) {
            inline = new ItalicSequenceInline(InlineSequenceOf(inline));
        }

        if (run.Strike || run.DoubleStrike) {
            inline = new StrikethroughSequenceInline(InlineSequenceOf(inline));
        }

        if (run.HighlightColorIndex.HasValue) {
            inline = new HighlightSequenceInline(InlineSequenceOf(inline));
        }

        return inline;
    }

    private static void AppendBreak(InlineSequence sequence, RtfBreak rtfBreak, RtfToMarkdownConversionContext context) {
        switch (rtfBreak.Kind) {
            case RtfBreakKind.Line:
            case RtfBreakKind.SoftLine:
                sequence.AddRaw(new HardBreakInline());
                break;
            case RtfBreakKind.Page:
            case RtfBreakKind.SoftPage:
                sequence.AddRaw(new HardBreakInline());
                sequence.AddRaw(new DecodedHtmlEntityTextRun("---"));
                sequence.AddRaw(new HardBreakInline());
                break;
            case RtfBreakKind.Column:
                sequence.AddRaw(new HardBreakInline());
                context.Report("RTFMD006", RtfMarkdownDiagnosticSeverity.Warning, "RTF column break represented as a Markdown hard break.", action: RtfConversionAction.Flattened);
                break;
        }
    }

    private static void AppendField(InlineSequence sequence, RtfField field, RtfToMarkdownConversionContext context, ref int imageIndex) {
        InlineSequence result = ConvertParagraphInlines(field.Result, context, ref imageIndex);
        if (field.Hyperlink != null) {
            InlineSequence label = result.Nodes.Count == 0
                ? InlineSequenceOf(new DecodedHtmlEntityTextRun(field.Hyperlink.ToString()))
                : result;
            sequence.AddRaw(new LinkInline(label, FormatMarkdownLinkDestination(field.Hyperlink.ToString()), null));
            return;
        }

        for (int i = 0; i < result.Nodes.Count; i++) {
            sequence.AddRaw(result.Nodes[i]);
        }

        context.Report("RTFMD007", RtfMarkdownDiagnosticSeverity.Info, "RTF field converted using visible field result.", field.Instruction, RtfConversionAction.Flattened);
    }

    private static void AppendGeneratedText(InlineSequence sequence, RtfGeneratedText generatedText, RtfToMarkdownConversionContext context) {
        string text = generatedText.ToPlainText();
        if (generatedText.Note != null) {
            if (AppendNoteReference(sequence, generatedText.Note, context)) {
                context.Report("RTFMD008", RtfMarkdownDiagnosticSeverity.Info, "RTF note reference converted to a Markdown footnote reference.");
                return;
            }
        }

        if (!string.IsNullOrEmpty(text)) {
            sequence.AddRaw(new DecodedHtmlEntityTextRun(text));
        } else {
            context.Report("RTFMD013", RtfMarkdownDiagnosticSeverity.Warning, "RTF generated text omitted because no fallback text is available.", generatedText.Kind.ToString(), RtfConversionAction.Omitted);
            if (context.EmitUnsupportedHtmlComments) {
                sequence.AddRaw(new HtmlRawInline("<!-- RTF generated text omitted because no fallback text is available. -->"));
            }
        }
    }

    private static void AppendImageInline(InlineSequence sequence, RtfImage image, RtfToMarkdownConversionContext context, ref int imageIndex) {
        int currentIndex = imageIndex++;
        string path = ResolveImagePath(image, currentIndex, context);
        string alt = string.IsNullOrWhiteSpace(image.Description) ? "RTF image" : image.Description!;
        sequence.AddRaw(new ImageInline(alt, path));
        context.Report("RTFMD009", RtfMarkdownDiagnosticSeverity.Info, "Inline RTF image represented by Markdown image reference.", path, RtfConversionAction.Flattened);
    }

    private static void AddUnsupportedBlock(ICollection<IMarkdownBlock> blocks, RtfToMarkdownConversionContext context, string message, string source) {
        context.Report("RTFMD010", RtfMarkdownDiagnosticSeverity.Warning, message, source, RtfConversionAction.Omitted);
        if (context.EmitUnsupportedHtmlComments) {
            blocks.Add(new HtmlRawBlock("<!-- " + message + " -->"));
        }
    }

    private static void AppendUnsupportedInline(InlineSequence sequence, RtfToMarkdownConversionContext context, string message, string source) {
        context.Report("RTFMD011", RtfMarkdownDiagnosticSeverity.Warning, message, source, RtfConversionAction.Omitted);
        if (context.EmitUnsupportedHtmlComments) {
            sequence.AddRaw(new DecodedHtmlEntityTextRun("[" + message.TrimEnd('.') + "]"));
        }
    }

    private static void ReportOmittedHeaderFooters(RtfDocument document, ICollection<IMarkdownBlock> blocks, RtfToMarkdownConversionContext context) {
        if (document.HeaderFooters.Count == 0) {
            return;
        }

        const string message = "RTF header/footer content omitted from Markdown output.";
        context.Report("RTFMD014", RtfMarkdownDiagnosticSeverity.Warning, message, document.HeaderFooters.Count.ToString(CultureInfo.InvariantCulture), RtfConversionAction.Omitted);
        if (context.EmitUnsupportedHtmlComments) {
            blocks.Add(new HtmlRawBlock("<!-- " + message + " -->"));
        }
    }

    private static string BuildDefaultImagePath(RtfImage image, int imageIndex) {
        string extension = image.Format.ToString().ToLowerInvariant();
        return "rtf-image-" + imageIndex.ToString(CultureInfo.InvariantCulture) + "." + extension;
    }

    private static string ResolveImagePath(RtfImage image, int imageIndex, RtfToMarkdownConversionContext context) {
        string logicalPath = context.ImagePathFactory?.Invoke(image, imageIndex) ?? BuildDefaultImagePath(image, imageIndex);
        context.ImageExporter?.Invoke(image, imageIndex, logicalPath);
        return FormatMarkdownLinkDestination(logicalPath);
    }

    private static bool AppendNoteReference(InlineSequence sequence, RtfNote? note, RtfToMarkdownConversionContext context) {
        if (note == null || context.NoteRegistry == null) return false;
        string? label = context.NoteRegistry.Register(note, context);
        if (label == null) return false;
        sequence.AddRaw(new FootnoteRefInline(label));
        return true;
    }

    private static void AppendNoteDefinitions(RtfDocument document, ICollection<IMarkdownBlock> blocks, RtfToMarkdownConversionContext context, ref int imageIndex) {
        RtfMarkdownNoteRegistry? registry = context.NoteRegistry;
        if (registry == null) return;
        foreach (RtfNote note in document.Notes) {
            registry.Register(note, context);
        }

        foreach (KeyValuePair<string, RtfNote> entry in registry.Ordered) {
            var noteBlocks = new List<IMarkdownBlock>(entry.Value.Paragraphs.Count);
            foreach (RtfParagraph paragraph in entry.Value.Paragraphs) {
                noteBlocks.Add(ConvertParagraph(document, paragraph, context, ref imageIndex));
            }

            blocks.Add(new FootnoteDefinitionBlock(entry.Key, entry.Value.ToPlainText(), noteBlocks));
        }

        if (registry.Ordered.Count > 0) {
            context.Report("RTFMD015", RtfMarkdownDiagnosticSeverity.Info, "RTF footnotes and endnotes were converted to Markdown footnote definitions.", registry.Ordered.Count.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static string FormatMarkdownLinkDestination(string destination) {
        if (string.IsNullOrEmpty(destination)) {
            return string.Empty;
        }

        var builder = new StringBuilder(destination.Length);
        for (int i = 0; i < destination.Length; i++) {
            char ch = destination[i];
            if (char.IsWhiteSpace(ch)) {
                byte[] bytes = Encoding.UTF8.GetBytes(new[] { ch });
                for (int b = 0; b < bytes.Length; b++) {
                    builder.Append('%');
                    builder.Append(bytes[b].ToString("X2", CultureInfo.InvariantCulture));
                }
            } else {
                builder.Append(ch);
            }
        }

        return builder.ToString();
    }

    private sealed class ListFrame {
        internal ListFrame(int level, RtfListKind kind, IMarkdownListBlock list, int? listId, int? listDefinitionId, int start) {
            Level = level;
            Kind = kind;
            List = list;
            ListId = listId;
            ListDefinitionId = listDefinitionId;
            Start = start;
        }

        internal int Level { get; }

        internal RtfListKind Kind { get; }

        internal IMarkdownListBlock List { get; }

        internal int? ListId { get; }

        internal int? ListDefinitionId { get; }

        internal int Start { get; }

        internal ListItem? LastItem { get; set; }
    }

    private readonly struct CellContent {
        internal CellContent(string markdown, InlineSequence inlines) {
            Markdown = markdown ?? string.Empty;
            Inlines = inlines ?? CreateInlineSequence();
        }

        internal string Markdown { get; }

        internal InlineSequence Inlines { get; }
    }
}
