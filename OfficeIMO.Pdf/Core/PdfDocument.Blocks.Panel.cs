using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Adds a paragraph inside a simple panel (background + optional border).</summary>
    public PdfDocument PanelParagraph(System.Action<PdfParagraphBuilder> compose, PanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        Guard.NotNull(compose, nameof(compose));
        Guard.ParagraphAlign(align, nameof(align), "Panel paragraph");
        var builder = new PdfParagraphBuilder(align, defaultColor);
        compose(builder);
        AddBlock(new PanelParagraphBlock(builder.Build().Runs, align, defaultColor, style));
        return this;
    }

    /// <summary>
    /// Adds a styled panel from common flow blocks such as paragraphs, headings, lists, simple tables, rules, and nested panel paragraphs.
    /// </summary>
    public PdfDocument Panel(System.Action<PdfItemCompose> compose, PanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        Guard.NotNull(compose, nameof(compose));
        Guard.ParagraphAlign(align, nameof(align), "Panel");
        var blocks = new System.Collections.Generic.List<IPdfBlock>();
        using (PushBlockScope(blocks.Add)) {
            compose(new PdfItemCompose(this));
        }

        AddBlock(CreatePanelParagraphBlock(blocks, style, align, defaultColor));
        return this;
    }

    internal static PanelParagraphBlock CreatePanelParagraphBlock(System.Collections.Generic.IEnumerable<IPdfBlock> blocks, PanelStyle? style, PdfAlign align, PdfColor? defaultColor) {
        Guard.NotNull(blocks, nameof(blocks));
        var runs = new System.Collections.Generic.List<TextRun>();
        bool wroteContent = false;
        foreach (IPdfBlock block in blocks) {
            AppendPanelBlockRuns(runs, block, ref wroteContent);
        }

        if (runs.Count == 0) {
            runs.Add(TextRun.Normal(string.Empty));
        }

        return new PanelParagraphBlock(runs, align, defaultColor, style);
    }

    private static void AppendPanelBlockRuns(System.Collections.Generic.List<TextRun> runs, IPdfBlock block, ref bool wroteContent) {
        switch (block) {
            case BookmarkBlock:
                return;
            case SpacerBlock spacer:
                if (spacer.Height > 0) {
                    AppendPanelLineBreak(runs, ref wroteContent);
                }

                return;
            case HorizontalRuleBlock:
                AppendPanelBlockBreak(runs, ref wroteContent);
                runs.Add(TextRun.Normal(new string('-', 32)));
                return;
            case RichParagraphBlock paragraph:
                if (paragraph.Runs.Count == 0) {
                    return;
                }

                AppendPanelBlockBreak(runs, ref wroteContent);
                runs.AddRange(paragraph.Runs);
                return;
            case PanelParagraphBlock panel:
                if (panel.Runs.Count == 0) {
                    return;
                }

                AppendPanelBlockBreak(runs, ref wroteContent);
                runs.AddRange(panel.Runs);
                return;
            case HeadingBlock heading:
                AppendPanelBlockBreak(runs, ref wroteContent);
                AppendPanelHeadingRun(runs, heading);
                return;
            case BulletListBlock bullets:
                AppendPanelListRuns(runs, bullets.RichItems, startNumber: null, ref wroteContent);
                return;
            case NumberedListBlock numbered:
                AppendPanelListRuns(runs, numbered.RichItems, numbered.StartNumber, ref wroteContent);
                return;
            case TableBlock table:
                AppendPanelTableRuns(runs, table, ref wroteContent);
                return;
            default:
                throw new System.NotSupportedException("Panel currently supports paragraphs, headings, lists, simple tables, horizontal rules, spacers, bookmarks, and nested panel paragraphs.");
        }
    }

    private static void AppendPanelHeadingRun(System.Collections.Generic.List<TextRun> runs, HeadingBlock heading) {
        double fontSize = heading.Style?.GetFontSize(heading.Level) ?? PdfHeadingStyle.GetDefaultFontSize(heading.Level);
        bool bold = heading.Style?.Bold ?? true;
        PdfColor? color = heading.Color ?? heading.Style?.Color;
        runs.Add(new TextRun(
            heading.Text,
            bold: bold,
            underline: heading.LinkUri != null || heading.LinkDestinationName != null,
            color: color,
            fontSize: fontSize,
            font: heading.Style?.Font,
            linkUri: heading.LinkUri,
            linkContents: heading.LinkContents,
            linkDestinationName: heading.LinkDestinationName,
            fontFamily: heading.Style?.FontFamily));
    }

    private static void AppendPanelListRuns(System.Collections.Generic.List<TextRun> runs, System.Collections.Generic.IReadOnlyList<PdfListItem> items, int? startNumber, ref bool wroteContent) {
        for (int i = 0; i < items.Count; i++) {
            AppendPanelLineBreak(runs, ref wroteContent);
            string marker = items[i].Marker ?? (startNumber.HasValue
                ? (startNumber.Value + i).ToString(System.Globalization.CultureInfo.InvariantCulture) + "."
                : "•");
            runs.Add(TextRun.Normal(marker + " "));
            runs.AddRange(items[i].Runs);
        }
    }

    private static void AppendPanelTableRuns(System.Collections.Generic.List<TextRun> runs, TableBlock table, ref bool wroteContent) {
        if (table.Cells.Count == 0) {
            return;
        }

        PdfTableStyle? tableStyle = table.Style;
        if (TryAppendChecklistTableRuns(runs, table, tableStyle, ref wroteContent)) {
            return;
        }

        int headerRowCount = tableStyle?.HeaderRowCount ?? 0;
        System.Collections.Generic.IReadOnlyList<PdfTableCell>? headers = headerRowCount > 0 ? table.Cells[0] : null;
        int startRow = headers == null ? 0 : 1;
        for (int rowIndex = startRow; rowIndex < table.Cells.Count; rowIndex++) {
            AppendPanelLineBreak(runs, ref wroteContent);
            System.Collections.Generic.IReadOnlyList<PdfTableCell> row = table.Cells[rowIndex];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                if (columnIndex > 0) {
                    runs.Add(TextRun.Normal(" | "));
                }

                if (headers != null && columnIndex < headers.Count) {
                    runs.Add(TextRun.Bolded(headers[columnIndex].Text + ": "));
                }

                runs.AddRange(row[columnIndex].Runs);
            }
        }
    }

    private static bool TryAppendChecklistTableRuns(System.Collections.Generic.List<TextRun> runs, TableBlock table, PdfTableStyle? tableStyle, ref bool wroteContent) {
        if (tableStyle?.CellIcons == null || table.ColumnCount < 2) {
            return false;
        }

        bool foundChecklistRow = false;
        for (int rowIndex = 0; rowIndex < table.Cells.Count; rowIndex++) {
            if (!tableStyle.CellIcons.TryGetValue((rowIndex, 0), out PdfCellIcon? icon)
                || (icon.Kind != PdfCellIconKind.CheckBoxChecked && icon.Kind != PdfCellIconKind.CheckBoxUnchecked)) {
                continue;
            }

            foundChecklistRow = true;
            AppendPanelLineBreak(runs, ref wroteContent);
            bool isChecked = icon.Kind == PdfCellIconKind.CheckBoxChecked;
            runs.Add(TextRun.Bolded(isChecked ? "Done: " : "Open: "));
            System.Collections.Generic.IReadOnlyList<PdfTableCell> row = table.Cells[rowIndex];
            for (int columnIndex = 1; columnIndex < row.Count; columnIndex++) {
                if (columnIndex > 1) {
                    runs.Add(TextRun.Normal(" "));
                }

                runs.AddRange(row[columnIndex].Runs);
            }
        }

        return foundChecklistRow;
    }

    private static void AppendPanelBlockBreak(System.Collections.Generic.List<TextRun> runs, ref bool wroteContent) {
        if (wroteContent) {
            runs.Add(TextRun.LineBreak());
            runs.Add(TextRun.LineBreak());
        }

        wroteContent = true;
    }

    private static void AppendPanelLineBreak(System.Collections.Generic.List<TextRun> runs, ref bool wroteContent) {
        if (wroteContent) {
            runs.Add(TextRun.LineBreak());
        }

        wroteContent = true;
    }
}
