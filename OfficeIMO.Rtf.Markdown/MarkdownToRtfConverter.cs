using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;

namespace OfficeIMO.Rtf.Markdown;

internal static class MarkdownToRtfConverter {
    internal static RtfDocument Convert(MarkdownDoc markdown, MarkdownToRtfOptions options) {
        var document = RtfDocument.Create();
        EnsureDocumentDefaults(document);

        for (int i = 0; i < markdown.Blocks.Count; i++) {
            ConvertBlock(document, markdown.Blocks[i], options);
        }

        return document;
    }

    private static void EnsureDocumentDefaults(RtfDocument document) {
        EnsureHighlightColor(document);
        document.AddFont("Consolas");
    }

    private static void ConvertBlock(RtfDocument document, IMarkdownBlock block, MarkdownToRtfOptions options) {
        switch (block) {
            case ParagraphBlock paragraph:
                AppendInlineSequence(document.AddParagraph(), paragraph.Inlines, document, options, InlineStyle.Normal);
                break;
            case HeadingBlock heading:
                ConvertHeading(document, heading, options);
                break;
            case UnorderedListBlock unorderedList:
                ConvertList(document, unorderedList.Items, RtfListKind.Bullet, options);
                break;
            case OrderedListBlock orderedList:
                ConvertList(document, orderedList.Items, RtfListKind.Decimal, options);
                break;
            case TableBlock table:
                ConvertTable(document, table, options);
                break;
            case ImageBlock image:
                ConvertImageBlock(document, image, options);
                break;
            case CodeBlock code:
                ConvertCodeBlock(document, code);
                break;
            case HtmlRawBlock html:
                ConvertRawHtml(document, html.Html, options, "Markdown raw HTML block");
                break;
            case QuoteBlock quote:
                ConvertChildBlocks(document, quote.ChildBlocks, options, "Markdown quote flattened to paragraphs.");
                break;
            case IChildMarkdownBlockContainer container:
                ConvertChildBlocks(document, container.ChildBlocks, options, block.GetType().Name + " child blocks flattened.");
                break;
            default:
                document.AddParagraph(block.RenderMarkdown());
                options.Report("MDRTF001", RtfMarkdownDiagnosticSeverity.Warning, "Markdown block converted using rendered Markdown fallback.", block.GetType().Name);
                break;
        }
    }

    private static void ConvertHeading(RtfDocument document, HeadingBlock heading, MarkdownToRtfOptions options) {
        int level = heading.Level < 1 ? 1 : heading.Level > 6 ? 6 : heading.Level;
        int styleId = 100 + level;
        RtfStyle style = document.AddStyle(styleId, "Heading " + level);
        style.OutlineLevel = level - 1;

        RtfParagraph paragraph = document.AddParagraph();
        paragraph.SetStyle(styleId);
        paragraph.OutlineLevel = level - 1;
        AppendInlineSequence(paragraph, heading.Inlines, document, options, InlineStyle.Normal);
    }

    private static void ConvertList(RtfDocument document, IReadOnlyList<ListItem> items, RtfListKind kind, MarkdownToRtfOptions options) {
        for (int i = 0; i < items.Count; i++) {
            ListItem item = items[i];
            RtfParagraph paragraph = document.AddParagraph();
            paragraph.SetList(kind == RtfListKind.Decimal ? 2 : 1, Math.Max(0, item.Level), kind);
            if (item.IsTask) {
                paragraph.AddText(item.Checked ? "[x] " : "[ ] ");
            }

            AppendInlineSequence(paragraph, item.Content, document, options, InlineStyle.Normal);

            for (int childIndex = 0; childIndex < item.ChildBlocks.Count; childIndex++) {
                ConvertBlock(document, item.ChildBlocks[childIndex], options);
            }
        }
    }

    private static void ConvertTable(RtfDocument document, TableBlock table, MarkdownToRtfOptions options) {
        int rowCount = table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0);
        int columnCount = Math.Max(table.Headers.Count, table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.Count));
        if (rowCount == 0 || columnCount == 0) {
            options.Report("MDRTF002", RtfMarkdownDiagnosticSeverity.Info, "Empty Markdown table omitted from RTF output.");
            return;
        }

        RtfTable rtfTable = document.AddTable(rowCount, columnCount);
        int rtfRowIndex = 0;
        if (table.Headers.Count > 0) {
            FillTableRow(rtfTable.Rows[rtfRowIndex++], table.Headers);
        }

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            FillTableRow(rtfTable.Rows[rtfRowIndex++], table.Rows[rowIndex]);
        }
    }

    private static void FillTableRow(RtfTableRow row, IReadOnlyList<string> cells) {
        for (int column = 0; column < row.Cells.Count; column++) {
            string text = column < cells.Count ? cells[column] : string.Empty;
            row.Cells[column].AddParagraph(text.Replace("<br>", "\n"));
        }
    }

    private static void ConvertImageBlock(RtfDocument document, ImageBlock image, MarkdownToRtfOptions options) {
        string label = string.IsNullOrWhiteSpace(image.PlainAlt) ? image.Path : image.PlainAlt!;
        document.AddParagraph("[Image: " + label + "]");
        options.Report("MDRTF003", RtfMarkdownDiagnosticSeverity.Warning, "Markdown image source represented as text placeholder; binary embedding requires caller-provided media bytes.", image.Path);
    }

    private static void ConvertCodeBlock(RtfDocument document, CodeBlock code) {
        int fontId = document.AddFont("Consolas");
        string[] lines = code.Content.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
        for (int i = 0; i < lines.Length; i++) {
            RtfParagraph paragraph = document.AddParagraph();
            paragraph.AddText(lines[i]).FontId = fontId;
        }
    }

    private static void ConvertRawHtml(RtfDocument document, string html, MarkdownToRtfOptions options, string source) {
        if (options.PreserveRawHtmlAsText) {
            document.AddParagraph(html);
        } else {
            options.Report("MDRTF004", RtfMarkdownDiagnosticSeverity.Warning, source + " omitted. Set PreserveRawHtmlAsText to keep it as visible text.", html);
        }
    }

    private static void ConvertChildBlocks(RtfDocument document, IReadOnlyList<IMarkdownBlock> blocks, MarkdownToRtfOptions options, string message) {
        options.Report("MDRTF005", RtfMarkdownDiagnosticSeverity.Info, message);
        for (int i = 0; i < blocks.Count; i++) {
            ConvertBlock(document, blocks[i], options);
        }
    }

    private static void AppendInlineSequence(RtfParagraph paragraph, InlineSequence sequence, RtfDocument document, MarkdownToRtfOptions options, InlineStyle style) {
        for (int i = 0; i < sequence.Nodes.Count; i++) {
            AppendInline(paragraph, sequence.Nodes[i], document, options, style);
        }
    }

    private static void AppendInline(RtfParagraph paragraph, IMarkdownInline inline, RtfDocument document, MarkdownToRtfOptions options, InlineStyle style) {
        switch (inline) {
            case TextRun text:
                AddStyledText(paragraph, text.Text, style);
                break;
            case BoldInline bold:
                AddStyledText(paragraph, bold.Text, style.WithBold());
                break;
            case ItalicInline italic:
                AddStyledText(paragraph, italic.Text, style.WithItalic());
                break;
            case BoldItalicInline boldItalic:
                AddStyledText(paragraph, boldItalic.Text, style.WithBold().WithItalic());
                break;
            case StrikethroughInline strike:
                AddStyledText(paragraph, strike.Text, style.WithStrike());
                break;
            case UnderlineInline underline:
                AddStyledText(paragraph, underline.Text, style.WithUnderline());
                break;
            case HighlightInline highlight:
                AddStyledText(paragraph, highlight.Text, style.WithHighlight(EnsureHighlightColor(document)));
                break;
            case CodeSpanInline code:
                AddStyledText(paragraph, code.Text, style.WithFont(document.AddFont("Consolas")));
                break;
            case LinkInline link:
                AppendLink(paragraph, link, document, options, style);
                break;
            case ImageInline image:
                AddStyledText(paragraph, "[Image: " + image.PlainAlt + "]", style);
                options.Report("MDRTF006", RtfMarkdownDiagnosticSeverity.Warning, "Markdown inline image represented as text placeholder; binary embedding requires caller-provided media bytes.", image.Src);
                break;
            case HardBreakInline:
                paragraph.AddLineBreak();
                break;
            case HtmlRawInline html:
                AppendInlineRawHtml(paragraph, html.Html, options, style);
                break;
            case BoldSequenceInline boldSequence:
                AppendInlineSequence(paragraph, boldSequence.Inlines, document, options, style.WithBold());
                break;
            case ItalicSequenceInline italicSequence:
                AppendInlineSequence(paragraph, italicSequence.Inlines, document, options, style.WithItalic());
                break;
            case BoldItalicSequenceInline boldItalicSequence:
                AppendInlineSequence(paragraph, boldItalicSequence.Inlines, document, options, style.WithBold().WithItalic());
                break;
            case StrikethroughSequenceInline strikeSequence:
                AppendInlineSequence(paragraph, strikeSequence.Inlines, document, options, style.WithStrike());
                break;
            case HighlightSequenceInline highlightSequence:
                AppendInlineSequence(paragraph, highlightSequence.Inlines, document, options, style.WithHighlight(EnsureHighlightColor(document)));
                break;
            case IInlineContainerMarkdownInline container when container.NestedInlines != null:
                AppendInlineSequence(paragraph, container.NestedInlines!, document, options, style);
                break;
            default:
                AddStyledText(paragraph, RtfMarkdownText.PlainText(inline), style);
                options.Report("MDRTF007", RtfMarkdownDiagnosticSeverity.Info, "Markdown inline converted using plain text fallback.", inline.GetType().Name);
                break;
        }
    }

    private static void AppendLink(RtfParagraph paragraph, LinkInline link, RtfDocument document, MarkdownToRtfOptions options, InlineStyle style) {
        Uri? uri = null;
        if (!Uri.TryCreate(link.Url, UriKind.RelativeOrAbsolute, out uri)) {
            options.Report("MDRTF009", RtfMarkdownDiagnosticSeverity.Warning, "Markdown link URL was not valid for RTF hyperlink metadata.", link.Url);
        }

        if (link.LabelInlines != null) {
            int before = paragraph.Inlines.Count;
            AppendInlineSequence(paragraph, link.LabelInlines, document, options, style);
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

        RtfRun simpleRun = AddStyledText(paragraph, link.Text, style);
        if (uri != null) {
            simpleRun.SetHyperlink(uri);
        }
    }

    private static void AppendInlineRawHtml(RtfParagraph paragraph, string html, MarkdownToRtfOptions options, InlineStyle style) {
        if (options.PreserveRawHtmlAsText) {
            AddStyledText(paragraph, html, style);
        } else {
            options.Report("MDRTF010", RtfMarkdownDiagnosticSeverity.Warning, "Markdown raw inline HTML omitted. Set PreserveRawHtmlAsText to keep it as visible text.", html);
        }
    }

    private static RtfRun AddStyledText(RtfParagraph paragraph, string text, InlineStyle style) {
        RtfRun run = paragraph.AddText(text ?? string.Empty);
        if (style.Bold) run.SetBold();
        if (style.Italic) run.SetItalic();
        if (style.Strike) run.SetStrike();
        if (style.Underline) run.SetUnderline(RtfUnderlineStyle.Single);
        if (style.HighlightColorIndex.HasValue) run.SetHighlightColor(style.HighlightColorIndex.Value);
        if (style.FontId.HasValue) run.FontId = style.FontId.Value;
        return run;
    }

    private static int EnsureHighlightColor(RtfDocument document) {
        return document.AddColor(255, 255, 0);
    }

    private readonly struct InlineStyle {
        internal static readonly InlineStyle Normal = new InlineStyle(false, false, false, false, null, null);

        private InlineStyle(bool bold, bool italic, bool strike, bool underline, int? highlightColorIndex, int? fontId) {
            Bold = bold;
            Italic = italic;
            Strike = strike;
            Underline = underline;
            HighlightColorIndex = highlightColorIndex;
            FontId = fontId;
        }

        internal bool Bold { get; }

        internal bool Italic { get; }

        internal bool Strike { get; }

        internal bool Underline { get; }

        internal int? HighlightColorIndex { get; }

        internal int? FontId { get; }

        internal InlineStyle WithBold() => new InlineStyle(true, Italic, Strike, Underline, HighlightColorIndex, FontId);

        internal InlineStyle WithItalic() => new InlineStyle(Bold, true, Strike, Underline, HighlightColorIndex, FontId);

        internal InlineStyle WithStrike() => new InlineStyle(Bold, Italic, true, Underline, HighlightColorIndex, FontId);

        internal InlineStyle WithUnderline() => new InlineStyle(Bold, Italic, Strike, true, HighlightColorIndex, FontId);

        internal InlineStyle WithHighlight(int colorIndex) => new InlineStyle(Bold, Italic, Strike, Underline, colorIndex, FontId);

        internal InlineStyle WithFont(int fontId) => new InlineStyle(Bold, Italic, Strike, Underline, HighlightColorIndex, fontId);
    }
}
