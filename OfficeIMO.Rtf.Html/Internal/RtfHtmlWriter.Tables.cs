using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
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
        builder.Append("<tr");
        RtfTableRow row = table.Rows[rowIndex];
        AppendRowDirectionAttributes(builder, row);
        AppendTableRowMetadataAttributes(builder, row);
        AppendRowStyle(builder, row, document);
        builder.Append('>');
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
            AppendTableCellMetadataAttributes(builder, cell, (cellIndex + 1) * 2400);
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

    private static void AppendRowStyle(StringBuilder builder, RtfTableRow row, RtfDocument document) {
        if (!TryGetRowStyle(row, document, out string? style)) {
            return;
        }

        builder.Append(" style=\"");
        builder.Append(EncodeAttribute(style!));
        builder.Append('"');
    }

    private static bool TryGetRowStyle(RtfTableRow row, RtfDocument document, out string? style) {
        var builder = new StringBuilder();
        if (TryGetColor(document, row.BackgroundColorIndex, out RtfColor? background)) {
            builder.Append("background-color:");
            builder.Append(FormatColor(background!));
            builder.Append(';');
        }

        AppendTableRowShadingStyle(builder, row, document);

        if (row.Alignment.HasValue) {
            builder.Append("text-align:");
            builder.Append(FormatTableAlignment(row.Alignment.Value));
            builder.Append(';');
        }

        if (row.PreferredWidth.HasValue && row.PreferredWidthUnit.HasValue) {
            builder.Append("width:");
            builder.Append(FormatTableWidth(row.PreferredWidth.Value, row.PreferredWidthUnit.Value));
            builder.Append(';');
        }

        if (row.HeightTwips.HasValue) {
            builder.Append("height:");
            builder.Append(FormatPoints(row.HeightTwips.Value / 20d));
            builder.Append("pt;");
        }

        AppendCellPaddingStyle(builder, "padding-top", row.PaddingTopTwips);
        AppendCellPaddingStyle(builder, "padding-left", row.PaddingLeftTwips);
        AppendCellPaddingStyle(builder, "padding-bottom", row.PaddingBottomTwips);
        AppendCellPaddingStyle(builder, "padding-right", row.PaddingRightTwips);
        AppendRowDirectionStyle(builder, row);

        style = builder.Length == 0 ? null : builder.ToString();
        return style != null;
    }

    private static void AppendRowDirectionAttributes(StringBuilder builder, RtfTableRow row) {
        string? direction = FormatTableRowDirection(row.Direction);
        if (direction == null) {
            return;
        }

        builder.Append(" dir=\"");
        builder.Append(direction);
        builder.Append('"');
    }

    private static void AppendRowDirectionStyle(StringBuilder builder, RtfTableRow row) {
        string? direction = FormatTableRowDirection(row.Direction);
        if (direction == null) {
            return;
        }

        builder.Append("direction:");
        builder.Append(direction);
        builder.Append(";unicode-bidi:isolate;--officeimo-rtf-direction:");
        builder.Append(direction);
        builder.Append(';');
    }

    private static string? FormatTableRowDirection(RtfTableRowDirection? direction) {
        if (!direction.HasValue) {
            return null;
        }

        return direction.Value == RtfTableRowDirection.RightToLeft ? "rtl" : "ltr";
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

        AppendTableCellShadingStyle(builder, cell, document);

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

        AppendCellTextFlowStyle(builder, cell.TextFlow);

        if (cell.NoWrap) {
            builder.Append("white-space:nowrap;");
        }

        AppendCellFlagStyles(builder, cell);
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

    private static void AppendCellFlagStyles(StringBuilder builder, RtfTableCell cell) {
        if (cell.HideCellMark) {
            builder.Append("--officeimo-rtf-hide-cell-mark:true;");
        }

        if (cell.NoWrap) {
            builder.Append("--officeimo-rtf-cell-nowrap:true;");
        }

        if (cell.FitText) {
            builder.Append("--officeimo-rtf-fit-text:true;");
        }
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

    private static void AppendCellTextFlowStyle(StringBuilder builder, RtfTableCellTextFlow? textFlow) {
        if (!textFlow.HasValue) {
            return;
        }

        builder.Append("writing-mode:");
        builder.Append(FormatCellWritingMode(textFlow.Value));
        builder.Append(';');
        string? orientation = FormatCellTextOrientation(textFlow.Value);
        if (orientation != null) {
            builder.Append("text-orientation:");
            builder.Append(orientation);
            builder.Append(';');
        }

        builder.Append("--officeimo-rtf-text-flow:");
        builder.Append(FormatRtfTableCellTextFlow(textFlow.Value));
        builder.Append(';');
    }

    private static string FormatCellWritingMode(RtfTableCellTextFlow textFlow) {
        switch (textFlow) {
            case RtfTableCellTextFlow.TopToBottomRightToLeft:
            case RtfTableCellTextFlow.TopToBottomRightToLeftVertical:
                return "vertical-rl";
            case RtfTableCellTextFlow.BottomToTopLeftToRight:
                return "sideways-lr";
            case RtfTableCellTextFlow.LeftToRightTopToBottomVertical:
                return "vertical-lr";
            default:
                return "horizontal-tb";
        }
    }

    private static string? FormatCellTextOrientation(RtfTableCellTextFlow textFlow) {
        switch (textFlow) {
            case RtfTableCellTextFlow.LeftToRightTopToBottomVertical:
            case RtfTableCellTextFlow.TopToBottomRightToLeftVertical:
                return "upright";
            default:
                return null;
        }
    }

    private static string FormatRtfTableCellTextFlow(RtfTableCellTextFlow textFlow) {
        switch (textFlow) {
            case RtfTableCellTextFlow.TopToBottomRightToLeft:
                return "tb-rl";
            case RtfTableCellTextFlow.BottomToTopLeftToRight:
                return "bt-lr";
            case RtfTableCellTextFlow.LeftToRightTopToBottomVertical:
                return "ltr-tb-v";
            case RtfTableCellTextFlow.TopToBottomRightToLeftVertical:
                return "tb-rl-v";
            default:
                return "ltr-tb";
        }
    }
}
