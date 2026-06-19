using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Markdown;

internal static class RtfToMarkdownConverter {
    internal static MarkdownDoc Convert(RtfDocument document, RtfToMarkdownOptions options) {
        var markdown = MarkdownDoc.Create();
        int imageIndex = 0;

        for (int i = 0; i < document.Blocks.Count; i++) {
            IRtfBlock block = document.Blocks[i];
            switch (block) {
                case RtfParagraph paragraph:
                    if (paragraph.ListKind != RtfListKind.None) {
                        i = ConvertListRun(document, options, markdown, i, ref imageIndex);
                    } else {
                        markdown.Add(ConvertParagraph(document, paragraph, options, ref imageIndex));
                    }
                    break;
                case RtfTable table:
                    markdown.Add(ConvertTable(document, table, options, ref imageIndex));
                    break;
                case RtfImage image:
                    markdown.Add(ConvertImageBlock(image, options, ref imageIndex));
                    break;
                case RtfObject:
                    AddUnsupportedBlock(markdown, options, "RTF object block omitted.", "rtf-object");
                    break;
                case RtfShape:
                    AddUnsupportedBlock(markdown, options, "RTF drawing shape block omitted.", "rtf-shape");
                    break;
                default:
                    options.Report("RTFMD001", RtfMarkdownDiagnosticSeverity.Warning, "Unsupported RTF block omitted.", block.GetType().Name);
                    break;
            }
        }

        return markdown;
    }

    private static int ConvertListRun(RtfDocument document, RtfToMarkdownOptions options, MarkdownDoc markdown, int startIndex, ref int imageIndex) {
        var first = (RtfParagraph)document.Blocks[startIndex];
        int? firstListId = first.ListId;
        int? firstListDefinitionId = first.ListDefinitionId;
        var paragraphs = new List<RtfParagraph>();
        int i = startIndex;

        for (; i < document.Blocks.Count; i++) {
            if (!(document.Blocks[i] is RtfParagraph paragraph) || paragraph.ListKind == RtfListKind.None) {
                break;
            }

            int level = Math.Max(0, paragraph.ListLevel ?? 0);
            if (paragraph.ListId != firstListId ||
                paragraph.ListDefinitionId != firstListDefinitionId) {
                break;
            }

            if (level == 0 && paragraph.ListKind != first.ListKind) {
                break;
            }

            paragraphs.Add(paragraph);
        }

        markdown.Add(ConvertListParagraphs(document, options, paragraphs, ref imageIndex));
        return i - 1;
    }

    private static IMarkdownBlock ConvertListParagraphs(RtfDocument document, RtfToMarkdownOptions options, IReadOnlyList<RtfParagraph> paragraphs, ref int imageIndex) {
        RtfParagraph first = paragraphs[0];
        IMarkdownListBlock root = CreateMarkdownListBlock(document, first);
        var frames = new List<ListFrame> {
            new ListFrame(0, first.ListKind, root)
        };

        for (int i = 0; i < paragraphs.Count; i++) {
            RtfParagraph paragraph = paragraphs[i];
            int level = Math.Max(0, paragraph.ListLevel ?? 0);
            RtfListKind kind = NormalizeListKind(paragraph.ListKind);
            var item = new ListItem(ConvertParagraphInlines(paragraph, options, ref imageIndex));

            ListFrame frame = GetOrCreateListFrame(document, frames, paragraph, level, kind);
            AddListItem(frame.List, item);
            frame.LastItem = item;
        }

        return (IMarkdownBlock)root;
    }

    private static ListFrame GetOrCreateListFrame(RtfDocument document, List<ListFrame> frames, RtfParagraph paragraph, int level, RtfListKind kind) {
        if (level <= 0) {
            while (frames.Count > 1) {
                frames.RemoveAt(frames.Count - 1);
            }

            return frames[0];
        }

        while (frames.Count > 0) {
            ListFrame current = frames[frames.Count - 1];
            if (current.Level < level || (current.Level == level && current.Kind == kind)) {
                break;
            }

            frames.RemoveAt(frames.Count - 1);
        }

        ListFrame last = frames[frames.Count - 1];
        if (last.Level == level && last.Kind == kind) {
            return last;
        }

        if (last.LastItem == null) {
            return frames[0];
        }

        IMarkdownListBlock childList = CreateMarkdownListBlock(document, paragraph);
        last.LastItem.Children.Add((IMarkdownBlock)childList);
        var childFrame = new ListFrame(level, kind, childList);
        frames.Add(childFrame);
        return childFrame;
    }

    private static IMarkdownListBlock CreateMarkdownListBlock(RtfDocument document, RtfParagraph paragraph) {
        RtfListKind kind = NormalizeListKind(paragraph.ListKind);
        return kind == RtfListKind.Decimal
            ? new OrderedListBlock { Start = ResolveListStart(document, paragraph) }
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

    private static int ResolveListStart(RtfDocument document, RtfParagraph paragraph) {
        int levelIndex = Math.Max(0, paragraph.ListLevel ?? 0);
        if (paragraph.ListId.HasValue) {
            RtfListOverride? listOverride = document.ListOverrides.FirstOrDefault(item => item.Id == paragraph.ListId.Value);
            RtfListLevelOverride? levelOverride = listOverride?.LevelOverrides.ElementAtOrDefault(levelIndex);
            if (levelOverride?.StartAt.HasValue == true) {
                return Math.Max(1, levelOverride.StartAt.Value);
            }
        }

        if (paragraph.ListDefinitionId.HasValue) {
            RtfListDefinition? definition = document.ListDefinitions.FirstOrDefault(item => item.Id == paragraph.ListDefinitionId.Value);
            RtfListLevel? level = definition?.Levels.FirstOrDefault(item => item.LevelIndex == levelIndex);
            if (level?.StartAt.HasValue == true) {
                return Math.Max(1, level.StartAt.Value);
            }
        }

        return Math.Max(1, paragraph.LegacyNumbering.StartAt ?? 1);
    }

    private static IMarkdownBlock ConvertParagraph(RtfDocument document, RtfParagraph paragraph, RtfToMarkdownOptions options, ref int imageIndex) {
        InlineSequence inlines = ConvertParagraphInlines(paragraph, options, ref imageIndex);
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

    private static TableBlock ConvertTable(RtfDocument document, RtfTable table, RtfToMarkdownOptions options, ref int imageIndex) {
        var markdown = new TableBlock();
        if (table.Rows.Count == 0) {
            options.Report("RTFMD002", RtfMarkdownDiagnosticSeverity.Info, "Empty RTF table converted to an empty Markdown table.");
            return markdown;
        }

        bool hasHeader = table.Rows[0].RepeatHeader;
        int firstBodyRow = hasHeader ? 1 : 0;
        if (hasHeader) {
            RtfTableRow firstRow = table.Rows[0];
            for (int column = 0; column < firstRow.Cells.Count; column++) {
                markdown.Headers.Add(ConvertCellMarkdown(firstRow.Cells[column], options, ref imageIndex));
            }
        }

        for (int rowIndex = firstBodyRow; rowIndex < table.Rows.Count; rowIndex++) {
            var row = table.Rows[rowIndex];
            var cells = new List<string>(row.Cells.Count);
            for (int column = 0; column < row.Cells.Count; column++) {
                cells.Add(ConvertCellMarkdown(row.Cells[column], options, ref imageIndex));
            }

            markdown.Rows.Add(cells);
        }

        return markdown;
    }

    private static string ConvertCellMarkdown(RtfTableCell cell, RtfToMarkdownOptions options, ref int imageIndex) {
        var parts = new List<string>();
        for (int i = 0; i < cell.Paragraphs.Count; i++) {
            string text = RenderInlineSequenceMarkdown(ConvertParagraphInlines(cell.Paragraphs[i], options, ref imageIndex));
            if (!string.IsNullOrEmpty(text)) {
                parts.Add(text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "<br>"));
            }
        }

        return string.Join("<br>", parts);
    }

    private static ImageBlock ConvertImageBlock(RtfImage image, RtfToMarkdownOptions options, ref int imageIndex) {
        int currentIndex = imageIndex++;
        string path = options.ImagePathFactory?.Invoke(image, currentIndex) ?? BuildDefaultImagePath(image, currentIndex);
        string alt = string.IsNullOrWhiteSpace(image.Description) ? "RTF image" : image.Description!;
        options.Report("RTFMD003", RtfMarkdownDiagnosticSeverity.Info, "RTF image payload represented by Markdown image reference.", path);
        return new ImageBlock(path, alt, null);
    }

    private static InlineSequence ConvertParagraphInlines(RtfParagraph paragraph, RtfToMarkdownOptions options, ref int imageIndex) {
        InlineSequence sequence = CreateInlineSequence();
        for (int i = 0; i < paragraph.Inlines.Count; i++) {
            AppendInline(sequence, paragraph.Inlines[i], options, ref imageIndex);
        }

        return sequence;
    }

    private static InlineSequence CreateInlineSequence() {
        return new InlineSequence { AutoSpacing = false };
    }

    private static string RenderInlineSequenceMarkdown(InlineSequence sequence) {
        return ((IRenderableMarkdownInline)sequence).RenderMarkdown();
    }

    private static InlineSequence InlineSequenceOf(IMarkdownInline inline) {
        InlineSequence sequence = CreateInlineSequence();
        sequence.AddRaw(inline);
        return sequence;
    }

    private static void AppendInline(InlineSequence sequence, IRtfInline inline, RtfToMarkdownOptions options, ref int imageIndex) {
        switch (inline) {
            case RtfRun run:
                AppendRun(sequence, run, options);
                break;
            case RtfBreak rtfBreak:
                AppendBreak(sequence, rtfBreak, options);
                break;
            case RtfField field:
                AppendField(sequence, field, options, ref imageIndex);
                break;
            case RtfGeneratedText generatedText:
                AppendGeneratedText(sequence, generatedText, options);
                break;
            case RtfImage image:
                AppendImageInline(sequence, image, options, ref imageIndex);
                break;
            case RtfObject:
                AppendUnsupportedInline(sequence, options, "RTF object inline omitted.", "rtf-object");
                break;
            case RtfShape:
                AppendUnsupportedInline(sequence, options, "RTF drawing shape inline omitted.", "rtf-shape");
                break;
            case RtfBookmarkMarker:
                break;
            default:
                options.Report("RTFMD004", RtfMarkdownDiagnosticSeverity.Warning, "Unsupported RTF inline omitted.", inline.GetType().Name);
                break;
        }
    }

    private static void AppendRun(InlineSequence sequence, RtfRun run, RtfToMarkdownOptions options) {
        if (run.Hidden && !options.IncludeHiddenText) {
            options.Report("RTFMD005", RtfMarkdownDiagnosticSeverity.Info, "Hidden RTF text omitted from Markdown output.");
            return;
        }

        IMarkdownInline? inline = BuildRunInline(run);
        if (inline == null) {
            return;
        }

        if (run.Hyperlink != null) {
            sequence.AddRaw(new LinkInline(InlineSequenceOf(inline), run.Hyperlink.ToString(), null));
            return;
        }

        sequence.AddRaw(inline);
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

    private static void AppendBreak(InlineSequence sequence, RtfBreak rtfBreak, RtfToMarkdownOptions options) {
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
                options.Report("RTFMD006", RtfMarkdownDiagnosticSeverity.Warning, "RTF column break represented as a Markdown hard break.");
                break;
        }
    }

    private static void AppendField(InlineSequence sequence, RtfField field, RtfToMarkdownOptions options, ref int imageIndex) {
        InlineSequence result = ConvertParagraphInlines(field.Result, options, ref imageIndex);
        if (field.Hyperlink != null) {
            InlineSequence label = result.Nodes.Count == 0
                ? InlineSequenceOf(new DecodedHtmlEntityTextRun(field.Hyperlink.ToString()))
                : result;
            sequence.AddRaw(new LinkInline(label, field.Hyperlink.ToString(), null));
            return;
        }

        for (int i = 0; i < result.Nodes.Count; i++) {
            sequence.AddRaw(result.Nodes[i]);
        }

        options.Report("RTFMD007", RtfMarkdownDiagnosticSeverity.Info, "RTF field converted using visible field result.", field.Instruction);
    }

    private static void AppendGeneratedText(InlineSequence sequence, RtfGeneratedText generatedText, RtfToMarkdownOptions options) {
        string text = generatedText.ToPlainText();
        if (generatedText.Note != null) {
            options.Report("RTFMD008", RtfMarkdownDiagnosticSeverity.Warning, "RTF note reference converted using fallback text.");
        }

        if (!string.IsNullOrEmpty(text)) {
            sequence.AddRaw(new DecodedHtmlEntityTextRun(text));
        }
    }

    private static void AppendImageInline(InlineSequence sequence, RtfImage image, RtfToMarkdownOptions options, ref int imageIndex) {
        int currentIndex = imageIndex++;
        string path = options.ImagePathFactory?.Invoke(image, currentIndex) ?? BuildDefaultImagePath(image, currentIndex);
        string alt = string.IsNullOrWhiteSpace(image.Description) ? "RTF image" : image.Description!;
        sequence.AddRaw(new ImageInline(alt, path));
        options.Report("RTFMD009", RtfMarkdownDiagnosticSeverity.Info, "Inline RTF image represented by Markdown image reference.", path);
    }

    private static void AddUnsupportedBlock(MarkdownDoc markdown, RtfToMarkdownOptions options, string message, string source) {
        options.Report("RTFMD010", RtfMarkdownDiagnosticSeverity.Warning, message, source);
        if (options.EmitUnsupportedHtmlComments) {
            markdown.Add(new HtmlRawBlock("<!-- " + message + " -->"));
        }
    }

    private static void AppendUnsupportedInline(InlineSequence sequence, RtfToMarkdownOptions options, string message, string source) {
        options.Report("RTFMD011", RtfMarkdownDiagnosticSeverity.Warning, message, source);
        if (options.EmitUnsupportedHtmlComments) {
            sequence.AddRaw(new HtmlRawInline("<!-- " + message + " -->"));
        }
    }

    private static string BuildDefaultImagePath(RtfImage image, int imageIndex) {
        string extension = image.Format.ToString().ToLowerInvariant();
        return "rtf-image-" + imageIndex.ToString(CultureInfo.InvariantCulture) + "." + extension;
    }

    private sealed class ListFrame {
        internal ListFrame(int level, RtfListKind kind, IMarkdownListBlock list) {
            Level = level;
            Kind = kind;
            List = list;
        }

        internal int Level { get; }

        internal RtfListKind Kind { get; }

        internal IMarkdownListBlock List { get; }

        internal ListItem? LastItem { get; set; }
    }
}
