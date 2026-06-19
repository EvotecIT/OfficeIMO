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
                    markdown.Add(ConvertTable(table, options));
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
        bool ordered = first.ListKind == RtfListKind.Decimal;
        var unordered = ordered ? null : new UnorderedListBlock();
        var orderedList = ordered ? new OrderedListBlock() : null;
        int i = startIndex;

        for (; i < document.Blocks.Count; i++) {
            if (!(document.Blocks[i] is RtfParagraph paragraph) || paragraph.ListKind == RtfListKind.None) {
                break;
            }

            bool sameListFamily = ordered
                ? paragraph.ListKind == RtfListKind.Decimal
                : paragraph.ListKind != RtfListKind.Decimal;
            if (!sameListFamily) {
                break;
            }

            string inlineMarkdown = ConvertParagraphInlineMarkdown(paragraph, options, ref imageIndex);
            var item = new ListItem(MarkdownReader.ParseInlineText(inlineMarkdown, options.InlineReaderOptions)) {
                Level = Math.Max(0, paragraph.ListLevel ?? 0)
            };

            if (orderedList != null) {
                orderedList.Items.Add(item);
            } else {
                unordered!.Items.Add(item);
            }
        }

        markdown.Add(orderedList as IMarkdownBlock ?? unordered!);
        return i - 1;
    }

    private static IMarkdownBlock ConvertParagraph(RtfDocument document, RtfParagraph paragraph, RtfToMarkdownOptions options, ref int imageIndex) {
        string inlineMarkdown = ConvertParagraphInlineMarkdown(paragraph, options, ref imageIndex);
        var inlines = MarkdownReader.ParseInlineText(inlineMarkdown, options.InlineReaderOptions);
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

    private static TableBlock ConvertTable(RtfTable table, RtfToMarkdownOptions options) {
        var markdown = new TableBlock();
        if (table.Rows.Count == 0) {
            options.Report("RTFMD002", RtfMarkdownDiagnosticSeverity.Info, "Empty RTF table converted to an empty Markdown table.");
            return markdown;
        }

        RtfTableRow firstRow = table.Rows[0];
        for (int column = 0; column < firstRow.Cells.Count; column++) {
            markdown.Headers.Add(ConvertCellText(firstRow.Cells[column]));
        }

        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++) {
            var row = table.Rows[rowIndex];
            var cells = new List<string>(row.Cells.Count);
            for (int column = 0; column < row.Cells.Count; column++) {
                cells.Add(ConvertCellText(row.Cells[column]));
            }

            markdown.Rows.Add(cells);
        }

        return markdown;
    }

    private static string ConvertCellText(RtfTableCell cell) {
        var parts = new List<string>();
        for (int i = 0; i < cell.Paragraphs.Count; i++) {
            string text = cell.Paragraphs[i].ToPlainText();
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

    private static string ConvertParagraphInlineMarkdown(RtfParagraph paragraph, RtfToMarkdownOptions options, ref int imageIndex) {
        var sb = new StringBuilder();
        for (int i = 0; i < paragraph.Inlines.Count; i++) {
            AppendInline(sb, paragraph.Inlines[i], options, ref imageIndex);
        }

        return sb.ToString();
    }

    private static void AppendInline(StringBuilder sb, IRtfInline inline, RtfToMarkdownOptions options, ref int imageIndex) {
        switch (inline) {
            case RtfRun run:
                AppendRun(sb, run, options);
                break;
            case RtfBreak rtfBreak:
                AppendBreak(sb, rtfBreak, options);
                break;
            case RtfField field:
                AppendField(sb, field, options, ref imageIndex);
                break;
            case RtfGeneratedText generatedText:
                AppendGeneratedText(sb, generatedText, options);
                break;
            case RtfImage image:
                AppendImageInline(sb, image, options, ref imageIndex);
                break;
            case RtfObject:
                AppendUnsupportedInline(sb, options, "RTF object inline omitted.", "rtf-object");
                break;
            case RtfShape:
                AppendUnsupportedInline(sb, options, "RTF drawing shape inline omitted.", "rtf-shape");
                break;
            case RtfBookmarkMarker:
                break;
            default:
                options.Report("RTFMD004", RtfMarkdownDiagnosticSeverity.Warning, "Unsupported RTF inline omitted.", inline.GetType().Name);
                break;
        }
    }

    private static void AppendRun(StringBuilder sb, RtfRun run, RtfToMarkdownOptions options) {
        if (run.Hidden && !options.IncludeHiddenText) {
            options.Report("RTFMD005", RtfMarkdownDiagnosticSeverity.Info, "Hidden RTF text omitted from Markdown output.");
            return;
        }

        string text = RtfMarkdownText.EscapeMarkdownText(run.Text);
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        if (run.Hyperlink != null) {
            sb.Append('[').Append(text).Append("](").Append(RtfMarkdownText.EscapeLinkUrl(run.Hyperlink.ToString())).Append(')');
            return;
        }

        bool handledByHtml = false;
        if (run.VerticalPosition == RtfVerticalPosition.Superscript) {
            sb.Append("<sup>").Append(RtfMarkdownText.HtmlEncode(run.Text)).Append("</sup>");
            handledByHtml = true;
        } else if (run.VerticalPosition == RtfVerticalPosition.Subscript) {
            sb.Append("<sub>").Append(RtfMarkdownText.HtmlEncode(run.Text)).Append("</sub>");
            handledByHtml = true;
        } else if (run.UnderlineStyle != RtfUnderlineStyle.None) {
            sb.Append("<u>").Append(RtfMarkdownText.HtmlEncode(run.Text)).Append("</u>");
            handledByHtml = true;
        }

        if (handledByHtml) {
            return;
        }

        if (run.Bold && run.Italic) {
            text = "***" + text + "***";
        } else if (run.Bold) {
            text = "**" + text + "**";
        } else if (run.Italic) {
            text = "_" + text + "_";
        }

        if (run.Strike || run.DoubleStrike) {
            text = "~~" + text + "~~";
        }

        if (run.HighlightColorIndex.HasValue) {
            text = "==" + text + "==";
        }

        sb.Append(text);
    }

    private static void AppendBreak(StringBuilder sb, RtfBreak rtfBreak, RtfToMarkdownOptions options) {
        switch (rtfBreak.Kind) {
            case RtfBreakKind.Line:
            case RtfBreakKind.SoftLine:
                sb.Append("  \n");
                break;
            case RtfBreakKind.Page:
            case RtfBreakKind.SoftPage:
                sb.Append("\n\n---\n\n");
                break;
            case RtfBreakKind.Column:
                sb.Append("  \n");
                options.Report("RTFMD006", RtfMarkdownDiagnosticSeverity.Warning, "RTF column break represented as a Markdown hard break.");
                break;
        }
    }

    private static void AppendField(StringBuilder sb, RtfField field, RtfToMarkdownOptions options, ref int imageIndex) {
        string text = ConvertParagraphInlineMarkdown(field.Result, options, ref imageIndex);
        if (field.Hyperlink != null) {
            string label = string.IsNullOrWhiteSpace(text) ? RtfMarkdownText.EscapeMarkdownText(field.Hyperlink.ToString()) : text;
            sb.Append('[').Append(label).Append("](").Append(RtfMarkdownText.EscapeLinkUrl(field.Hyperlink.ToString())).Append(')');
            return;
        }

        sb.Append(text);
        options.Report("RTFMD007", RtfMarkdownDiagnosticSeverity.Info, "RTF field converted using visible field result.", field.Instruction);
    }

    private static void AppendGeneratedText(StringBuilder sb, RtfGeneratedText generatedText, RtfToMarkdownOptions options) {
        string text = RtfMarkdownText.EscapeMarkdownText(generatedText.ToPlainText());
        if (generatedText.Note != null) {
            options.Report("RTFMD008", RtfMarkdownDiagnosticSeverity.Warning, "RTF note reference converted using fallback text.");
        }

        sb.Append(text);
    }

    private static void AppendImageInline(StringBuilder sb, RtfImage image, RtfToMarkdownOptions options, ref int imageIndex) {
        int currentIndex = imageIndex++;
        string path = options.ImagePathFactory?.Invoke(image, currentIndex) ?? BuildDefaultImagePath(image, currentIndex);
        string alt = string.IsNullOrWhiteSpace(image.Description) ? "RTF image" : image.Description!;
        sb.Append("![").Append(RtfMarkdownText.EscapeImageAlt(alt)).Append("](").Append(RtfMarkdownText.EscapeLinkUrl(path)).Append(')');
        options.Report("RTFMD009", RtfMarkdownDiagnosticSeverity.Info, "Inline RTF image represented by Markdown image reference.", path);
    }

    private static void AddUnsupportedBlock(MarkdownDoc markdown, RtfToMarkdownOptions options, string message, string source) {
        options.Report("RTFMD010", RtfMarkdownDiagnosticSeverity.Warning, message, source);
        if (options.EmitUnsupportedHtmlComments) {
            markdown.Add(new HtmlRawBlock("<!-- " + message + " -->"));
        }
    }

    private static void AppendUnsupportedInline(StringBuilder sb, RtfToMarkdownOptions options, string message, string source) {
        options.Report("RTFMD011", RtfMarkdownDiagnosticSeverity.Warning, message, source);
        if (options.EmitUnsupportedHtmlComments) {
            sb.Append("<!-- ").Append(message).Append(" -->");
        }
    }

    private static string BuildDefaultImagePath(RtfImage image, int imageIndex) {
        string extension = image.Format.ToString().ToLowerInvariant();
        return "rtf-image-" + imageIndex.ToString(CultureInfo.InvariantCulture) + "." + extension;
    }
}
