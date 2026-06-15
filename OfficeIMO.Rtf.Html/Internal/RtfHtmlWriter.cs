namespace OfficeIMO.Rtf.Html;

internal static class RtfHtmlWriter {
    internal static string Write(RtfDocument document, RtfHtmlSaveOptions options) {
        string newline = options.GetNewLine();
        var builder = new StringBuilder();

        if (!options.FragmentOnly) {
            AppendDocumentStart(builder, document, options, newline);
        }

        RtfListKind openList = RtfListKind.None;
        for (int i = 0; i < document.Blocks.Count; i++) {
            IRtfBlock block = document.Blocks[i];
            if (block is RtfParagraph paragraph && paragraph.ListKind != RtfListKind.None) {
                if (openList != paragraph.ListKind) {
                    CloseList(builder, openList);
                    OpenList(builder, paragraph.ListKind);
                    openList = paragraph.ListKind;
                }

                AppendParagraph(builder, paragraph, options);
            } else {
                CloseList(builder, openList);
                openList = RtfListKind.None;
                AppendBlock(builder, block, options);
            }

            if (i + 1 < document.Blocks.Count) {
                builder.Append(newline);
            }
        }

        CloseList(builder, openList);

        if (!options.FragmentOnly) {
            builder.Append(newline);
            builder.Append("</body>");
            builder.Append(newline);
            builder.Append("</html>");
        }

        return builder.ToString();
    }

    private static void AppendDocumentStart(StringBuilder builder, RtfDocument document, RtfHtmlSaveOptions options, string newline) {
        builder.Append("<!doctype html>");
        builder.Append(newline);
        builder.Append("<html>");
        builder.Append(newline);
        builder.Append("<head>");
        builder.Append(newline);
        builder.Append("<meta charset=\"utf-8\">");
        if (options.IncludeMetadata) {
            string? title = options.Title ?? document.Info.Title;
            if (!string.IsNullOrWhiteSpace(title)) {
                builder.Append(newline);
                builder.Append("<title>");
                builder.Append(Encode(title!));
                builder.Append("</title>");
            }
        }

        builder.Append(newline);
        builder.Append("</head>");
        builder.Append(newline);
        builder.Append("<body>");
        builder.Append(newline);
    }

    private static void AppendBlock(StringBuilder builder, IRtfBlock block, RtfHtmlSaveOptions options) {
        switch (block) {
            case RtfParagraph paragraph:
                AppendParagraph(builder, paragraph, options);
                break;
            case RtfTable table:
                AppendTable(builder, table, options);
                break;
            case RtfImage image:
                AppendImage(builder, image, options);
                break;
            case RtfObject rtfObject:
                builder.Append("<p>");
                builder.Append(Encode(rtfObject.Kind.ToString()));
                builder.Append("</p>");
                break;
            default:
                break;
        }
    }

    private static void AppendParagraph(StringBuilder builder, RtfParagraph paragraph, RtfHtmlSaveOptions options) {
        string tagName = paragraph.ListKind == RtfListKind.None ? "p" : "li";
        builder.Append('<');
        builder.Append(tagName);
        AppendParagraphStyle(builder, paragraph);
        builder.Append('>');
        AppendInlines(builder, paragraph.Inlines, options);
        builder.Append("</");
        builder.Append(tagName);
        builder.Append('>');
    }

    private static void AppendParagraphStyle(StringBuilder builder, RtfParagraph paragraph) {
        string? align = paragraph.Alignment == RtfTextAlignment.Left ? null : paragraph.Alignment.ToString().ToLowerInvariant();
        if (align == null) {
            return;
        }

        builder.Append(" style=\"text-align:");
        builder.Append(align);
        builder.Append("\"");
    }

    private static void AppendInlines(StringBuilder builder, IReadOnlyList<IRtfInline> inlines, RtfHtmlSaveOptions options) {
        foreach (IRtfInline inline in inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(builder, run);
                    break;
                case RtfBreak rtfBreak when rtfBreak.Kind == RtfBreakKind.Line:
                    builder.Append("<br>");
                    break;
                case RtfField field:
                    AppendInlines(builder, field.Result.Inlines, options);
                    break;
                case RtfImage image:
                    AppendImage(builder, image, options);
                    break;
                case RtfBookmarkMarker marker when marker.Kind == RtfBookmarkMarkerKind.Start:
                    builder.Append("<a id=\"");
                    builder.Append(EncodeAttribute(marker.Name));
                    builder.Append("\"></a>");
                    break;
            }
        }
    }

    private static void AppendRun(StringBuilder builder, RtfRun run) {
        int opened = 0;
        if (run.Hyperlink != null) {
            builder.Append("<a href=\"");
            builder.Append(EncodeAttribute(run.Hyperlink.ToString()));
            builder.Append("\">");
            opened++;
        }

        OpenTag(builder, "strong", run.Bold, ref opened);
        OpenTag(builder, "em", run.Italic, ref opened);
        OpenTag(builder, "u", run.Underline, ref opened);
        OpenTag(builder, "s", run.Strike || run.DoubleStrike, ref opened);
        OpenTag(builder, "sup", run.VerticalPosition == RtfVerticalPosition.Superscript, ref opened);
        OpenTag(builder, "sub", run.VerticalPosition == RtfVerticalPosition.Subscript, ref opened);

        builder.Append(Encode(run.Text));

        CloseTag(builder, "sub", run.VerticalPosition == RtfVerticalPosition.Subscript);
        CloseTag(builder, "sup", run.VerticalPosition == RtfVerticalPosition.Superscript);
        CloseTag(builder, "s", run.Strike || run.DoubleStrike);
        CloseTag(builder, "u", run.Underline);
        CloseTag(builder, "em", run.Italic);
        CloseTag(builder, "strong", run.Bold);
        if (run.Hyperlink != null) {
            builder.Append("</a>");
        }
    }

    private static void OpenTag(StringBuilder builder, string tag, bool condition, ref int opened) {
        if (!condition) {
            return;
        }

        builder.Append('<');
        builder.Append(tag);
        builder.Append('>');
        opened++;
    }

    private static void CloseTag(StringBuilder builder, string tag, bool condition) {
        if (!condition) {
            return;
        }

        builder.Append("</");
        builder.Append(tag);
        builder.Append('>');
    }

    private static void AppendTable(StringBuilder builder, RtfTable table, RtfHtmlSaveOptions options) {
        builder.Append("<table><tbody>");
        foreach (RtfTableRow row in table.Rows) {
            builder.Append("<tr>");
            foreach (RtfTableCell cell in row.Cells) {
                builder.Append("<td>");
                for (int i = 0; i < cell.Paragraphs.Count; i++) {
                    AppendParagraph(builder, cell.Paragraphs[i], options);
                }

                builder.Append("</td>");
            }

            builder.Append("</tr>");
        }

        builder.Append("</tbody></table>");
    }

    private static void AppendImage(StringBuilder builder, RtfImage image, RtfHtmlSaveOptions options) {
        if (!options.EmbedImagesAsDataUri || image.Data.Length == 0 || !TryGetImageMediaType(image.Format, out string? mediaType)) {
            return;
        }

        builder.Append("<img src=\"data:");
        builder.Append(mediaType);
        builder.Append(";base64,");
        builder.Append(Convert.ToBase64String(image.Data));
        builder.Append('"');
        if (!string.IsNullOrWhiteSpace(image.Description)) {
            builder.Append(" alt=\"");
            builder.Append(EncodeAttribute(image.Description!));
            builder.Append('"');
        }

        builder.Append('>');
    }

    private static void OpenList(StringBuilder builder, RtfListKind kind) {
        if (kind == RtfListKind.None) {
            return;
        }

        builder.Append(kind == RtfListKind.Decimal ? "<ol>" : "<ul>");
    }

    private static void CloseList(StringBuilder builder, RtfListKind kind) {
        if (kind == RtfListKind.None) {
            return;
        }

        builder.Append(kind == RtfListKind.Decimal ? "</ol>" : "</ul>");
    }

    private static bool TryGetImageMediaType(RtfImageFormat format, out string? mediaType) {
        switch (format) {
            case RtfImageFormat.Png:
                mediaType = "image/png";
                return true;
            case RtfImageFormat.Jpeg:
                mediaType = "image/jpeg";
                return true;
            default:
                mediaType = null;
                return false;
        }
    }

    private static string Encode(string value) => WebUtility.HtmlEncode(value);

    private static string EncodeAttribute(string value) => WebUtility.HtmlEncode(value);
}
