using System.Globalization;

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

                AppendParagraph(builder, paragraph, options, document);
            } else {
                CloseList(builder, openList);
                openList = RtfListKind.None;
                AppendBlock(builder, block, options, document);
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

    private static void AppendBlock(StringBuilder builder, IRtfBlock block, RtfHtmlSaveOptions options, RtfDocument document) {
        switch (block) {
            case RtfParagraph paragraph:
                AppendParagraph(builder, paragraph, options, document);
                break;
            case RtfTable table:
                AppendTable(builder, table, options, document);
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

    private static void AppendParagraph(StringBuilder builder, RtfParagraph paragraph, RtfHtmlSaveOptions options, RtfDocument document) {
        string tagName = GetParagraphTagName(paragraph);
        builder.Append('<');
        builder.Append(tagName);
        AppendParagraphStyle(builder, paragraph);
        builder.Append('>');
        AppendInlines(builder, paragraph.Inlines, options, document);
        builder.Append("</");
        builder.Append(tagName);
        builder.Append('>');
    }

    private static string GetParagraphTagName(RtfParagraph paragraph) {
        if (paragraph.ListKind != RtfListKind.None) {
            return "li";
        }

        if (paragraph.OutlineLevel.HasValue && paragraph.OutlineLevel.Value >= 0 && paragraph.OutlineLevel.Value <= 5) {
            return "h" + (paragraph.OutlineLevel.Value + 1).ToString(CultureInfo.InvariantCulture);
        }

        return "p";
    }

    private static void AppendParagraphStyle(StringBuilder builder, RtfParagraph paragraph) {
        if (!TryGetParagraphStyle(paragraph, out string? style)) {
            return;
        }

        builder.Append(" style=\"");
        builder.Append(EncodeAttribute(style!));
        builder.Append("\"");
    }

    private static bool TryGetParagraphStyle(RtfParagraph paragraph, out string? style) {
        var builder = new StringBuilder();
        string? align = paragraph.Alignment == RtfTextAlignment.Left ? null : paragraph.Alignment.ToString().ToLowerInvariant();
        if (align != null) {
            builder.Append("text-align:");
            builder.Append(align);
            builder.Append(';');
        }

        if (paragraph.PageBreakBefore) {
            builder.Append("page-break-before:always;break-before:page;");
        }

        AppendTwipStyle(builder, "margin-left", paragraph.LeftIndentTwips);
        AppendTwipStyle(builder, "margin-right", paragraph.RightIndentTwips);
        AppendTwipStyle(builder, "text-indent", paragraph.FirstLineIndentTwips);

        style = builder.Length == 0 ? null : builder.ToString();
        return style != null;
    }

    private static void AppendTwipStyle(StringBuilder builder, string name, int? twips) {
        if (!twips.HasValue || twips.Value == 0) {
            return;
        }

        builder.Append(name);
        builder.Append(':');
        builder.Append(FormatPoints(twips.Value / 20d));
        builder.Append("pt;");
    }

    private static void AppendInlines(StringBuilder builder, IReadOnlyList<IRtfInline> inlines, RtfHtmlSaveOptions options, RtfDocument document) {
        foreach (IRtfInline inline in inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(builder, run, document);
                    break;
                case RtfBreak rtfBreak when rtfBreak.Kind == RtfBreakKind.Line:
                    builder.Append("<br>");
                    break;
                case RtfBreak rtfBreak when rtfBreak.Kind == RtfBreakKind.Page:
                    builder.Append("<br style=\"page-break-before:always;break-before:page;\">");
                    break;
                case RtfField field:
                    AppendInlines(builder, field.Result.Inlines, options, document);
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

    private static void AppendRun(StringBuilder builder, RtfRun run, RtfDocument document) {
        int opened = 0;
        if (run.Hyperlink != null) {
            builder.Append("<a href=\"");
            builder.Append(EncodeAttribute(run.Hyperlink.ToString()));
            builder.Append("\">");
            opened++;
        }

        OpenRunStyle(builder, run, document, ref opened);
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
        CloseRunStyle(builder, run, document);
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

    private static void AppendTable(StringBuilder builder, RtfTable table, RtfHtmlSaveOptions options, RtfDocument document) {
        builder.Append("<table>");
        bool inHead = false;
        bool inBody = false;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            RtfTableRow row = table.Rows[rowIndex];
            if (row.RepeatHeader) {
                if (inBody) {
                    builder.Append("</tbody>");
                    inBody = false;
                }

                if (!inHead) {
                    builder.Append("<thead>");
                    inHead = true;
                }
            } else {
                if (inHead) {
                    builder.Append("</thead>");
                    inHead = false;
                }

                if (!inBody) {
                    builder.Append("<tbody>");
                    inBody = true;
                }
            }

            AppendTableRow(builder, table, rowIndex, options, document, row.RepeatHeader);
        }

        if (inHead) {
            builder.Append("</thead>");
        }

        if (inBody) {
            builder.Append("</tbody>");
        }

        builder.Append("</table>");
    }

    private static void AppendTableRow(StringBuilder builder, RtfTable table, int rowIndex, RtfHtmlSaveOptions options, RtfDocument document, bool isHeader) {
        builder.Append("<tr>");
        RtfTableRow row = table.Rows[rowIndex];
        string cellTag = isHeader ? "th" : "td";
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
            RtfTableCell cell = row.Cells[cellIndex];
            if (IsMergeContinuation(cell)) {
                continue;
            }

            int columnSpan = GetColumnSpan(row, cellIndex);
            int rowSpan = GetRowSpan(table, rowIndex, cellIndex, columnSpan);
            builder.Append('<');
            builder.Append(cellTag);
            AppendCellSpanAttributes(builder, columnSpan, rowSpan);
            AppendCellStyle(builder, cell, document);
            builder.Append('>');
            for (int i = 0; i < cell.Paragraphs.Count; i++) {
                AppendParagraph(builder, cell.Paragraphs[i], options, document);
            }

            builder.Append("</");
            builder.Append(cellTag);
            builder.Append('>');
        }

        builder.Append("</tr>");
    }

    private static bool IsMergeContinuation(RtfTableCell cell) {
        return cell.HorizontalMerge == RtfTableCellMerge.Continue ||
               cell.VerticalMerge == RtfTableCellMerge.Continue;
    }

    private static int GetColumnSpan(RtfTableRow row, int cellIndex) {
        if (row.Cells[cellIndex].HorizontalMerge != RtfTableCellMerge.First) {
            return 1;
        }

        int span = 1;
        for (int i = cellIndex + 1; i < row.Cells.Count; i++) {
            if (row.Cells[i].HorizontalMerge != RtfTableCellMerge.Continue) {
                break;
            }

            span++;
        }

        return span;
    }

    private static int GetRowSpan(RtfTable table, int rowIndex, int cellIndex, int columnSpan) {
        if (table.Rows[rowIndex].Cells[cellIndex].VerticalMerge != RtfTableCellMerge.First) {
            return 1;
        }

        int span = 1;
        for (int nextRowIndex = rowIndex + 1; nextRowIndex < table.Rows.Count; nextRowIndex++) {
            RtfTableRow nextRow = table.Rows[nextRowIndex];
            if (cellIndex + columnSpan > nextRow.Cells.Count || !IsVerticalContinuation(nextRow, cellIndex, columnSpan)) {
                break;
            }

            span++;
        }

        return span;
    }

    private static bool IsVerticalContinuation(RtfTableRow row, int cellIndex, int columnSpan) {
        for (int offset = 0; offset < columnSpan; offset++) {
            if (row.Cells[cellIndex + offset].VerticalMerge != RtfTableCellMerge.Continue) {
                return false;
            }
        }

        return true;
    }

    private static void AppendCellSpanAttributes(StringBuilder builder, int columnSpan, int rowSpan) {
        if (columnSpan > 1) {
            builder.Append(" colspan=\"");
            builder.Append(columnSpan.ToString(CultureInfo.InvariantCulture));
            builder.Append('"');
        }

        if (rowSpan > 1) {
            builder.Append(" rowspan=\"");
            builder.Append(rowSpan.ToString(CultureInfo.InvariantCulture));
            builder.Append('"');
        }
    }

    private static void AppendCellStyle(StringBuilder builder, RtfTableCell cell, RtfDocument document) {
        if (!TryGetCellStyle(cell, document, out string? style)) {
            return;
        }

        builder.Append(" style=\"");
        builder.Append(EncodeAttribute(style!));
        builder.Append('"');
    }

    private static bool TryGetCellStyle(RtfTableCell cell, RtfDocument document, out string? style) {
        var builder = new StringBuilder();
        if (TryGetColor(document, cell.BackgroundColorIndex, out RtfColor? background)) {
            builder.Append("background-color:");
            builder.Append(FormatColor(background!));
            builder.Append(';');
        }

        if (cell.PreferredWidth.HasValue && cell.PreferredWidthUnit.HasValue) {
            builder.Append("width:");
            builder.Append(FormatTableWidth(cell.PreferredWidth.Value, cell.PreferredWidthUnit.Value));
            builder.Append(';');
        }

        if (cell.VerticalAlignment.HasValue) {
            builder.Append("vertical-align:");
            builder.Append(FormatCellVerticalAlignment(cell.VerticalAlignment.Value));
            builder.Append(';');
        }

        AppendCellPaddingStyle(builder, "padding-top", cell.PaddingTopTwips);
        AppendCellPaddingStyle(builder, "padding-left", cell.PaddingLeftTwips);
        AppendCellPaddingStyle(builder, "padding-bottom", cell.PaddingBottomTwips);
        AppendCellPaddingStyle(builder, "padding-right", cell.PaddingRightTwips);
        AppendCellBorderStyle(builder, "border-top", cell.TopBorder, document);
        AppendCellBorderStyle(builder, "border-left", cell.LeftBorder, document);
        AppendCellBorderStyle(builder, "border-bottom", cell.BottomBorder, document);
        AppendCellBorderStyle(builder, "border-right", cell.RightBorder, document);

        style = builder.Length == 0 ? null : builder.ToString();
        return style != null;
    }

    private static void AppendCellPaddingStyle(StringBuilder builder, string name, int? twips) {
        if (!twips.HasValue) {
            return;
        }

        builder.Append(name);
        builder.Append(':');
        builder.Append(FormatPoints(twips.Value / 20d));
        builder.Append("pt;");
    }

    private static void AppendCellBorderStyle(StringBuilder builder, string name, RtfTableCellBorder border, RtfDocument document) {
        if (!border.HasAnyValue) {
            return;
        }

        builder.Append(name);
        builder.Append(':');
        if (border.Width.HasValue) {
            builder.Append(FormatPoints(border.Width.Value / 20d));
            builder.Append("pt ");
        }

        builder.Append(FormatCellBorderStyle(border.Style));
        if (TryGetColor(document, border.ColorIndex, out RtfColor? color)) {
            builder.Append(' ');
            builder.Append(FormatColor(color!));
        }

        builder.Append(';');
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

    private static void OpenRunStyle(StringBuilder builder, RtfRun run, RtfDocument document, ref int opened) {
        if (!TryGetRunStyle(run, document, out string? style)) {
            return;
        }

        builder.Append("<span style=\"");
        builder.Append(EncodeAttribute(style!));
        builder.Append("\">");
        opened++;
    }

    private static void CloseRunStyle(StringBuilder builder, RtfRun run, RtfDocument document) {
        if (TryGetRunStyle(run, document, out _)) {
            builder.Append("</span>");
        }
    }

    private static bool TryGetRunStyle(RtfRun run, RtfDocument document, out string? style) {
        var builder = new StringBuilder();
        if (TryGetFont(document, run.FontId, out RtfFont? font)) {
            builder.Append("font-family:");
            builder.Append(FormatFontFamily(font!.Name));
            builder.Append(';');
        }

        if (run.FontSize.HasValue) {
            builder.Append("font-size:");
            builder.Append(FormatPoints(run.FontSize.Value));
            builder.Append("pt;");
        }

        if (TryGetColor(document, run.ForegroundColorIndex, out RtfColor? foreground)) {
            builder.Append("color:");
            builder.Append(FormatColor(foreground!));
            builder.Append(';');
        }

        int? backgroundIndex = run.CharacterBackgroundColorIndex ?? run.HighlightColorIndex;
        if (TryGetColor(document, backgroundIndex, out RtfColor? background)) {
            builder.Append("background-color:");
            builder.Append(FormatColor(background!));
            builder.Append(';');
        }

        style = builder.Length == 0 ? null : builder.ToString();
        return style != null;
    }

    private static bool TryGetColor(RtfDocument document, int? index, out RtfColor? color) {
        if (!index.HasValue || index.Value <= 0 || index.Value > document.Colors.Count) {
            color = null;
            return false;
        }

        color = document.Colors[index.Value - 1];
        return true;
    }

    private static bool TryGetFont(RtfDocument document, int? id, out RtfFont? font) {
        if (!id.HasValue) {
            font = null;
            return false;
        }

        font = document.Fonts.FirstOrDefault(item => item.Id == id.Value);
        return font != null;
    }

    private static string FormatFontFamily(string fontFamily) {
        return "\"" + fontFamily.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }

    private static string FormatPoints(double points) {
        return points.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static string FormatTableWidth(int width, RtfTableWidthUnit unit) {
        switch (unit) {
            case RtfTableWidthUnit.Percent:
                return FormatPoints(width / 50d) + "%";
            case RtfTableWidthUnit.Auto:
                return "auto";
            default:
                return FormatPoints(width / 20d) + "pt";
        }
    }

    private static string FormatCellVerticalAlignment(RtfTableCellVerticalAlignment alignment) {
        switch (alignment) {
            case RtfTableCellVerticalAlignment.Center:
                return "middle";
            case RtfTableCellVerticalAlignment.Bottom:
                return "bottom";
            default:
                return "top";
        }
    }

    private static string FormatCellBorderStyle(RtfTableCellBorderStyle style) {
        switch (style) {
            case RtfTableCellBorderStyle.Double:
                return "double";
            case RtfTableCellBorderStyle.Dotted:
                return "dotted";
            case RtfTableCellBorderStyle.Dashed:
                return "dashed";
            case RtfTableCellBorderStyle.None:
                return "none";
            default:
                return "solid";
        }
    }

    private static string FormatColor(RtfColor color) {
        return "#" + color.Red.ToString("X2", CultureInfo.InvariantCulture) +
               color.Green.ToString("X2", CultureInfo.InvariantCulture) +
               color.Blue.ToString("X2", CultureInfo.InvariantCulture);
    }

    private static string Encode(string value) => WebUtility.HtmlEncode(value);

    private static string EncodeAttribute(string value) => WebUtility.HtmlEncode(value);
}
