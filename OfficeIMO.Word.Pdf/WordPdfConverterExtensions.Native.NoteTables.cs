using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Keeps note blocks in their authored order. Table content is projected by the
    /// ordinary table owner, including its effective cell styles and inline content.
    /// Callers first qualify the note's supported block profile.
    /// </summary>
    private static IReadOnlyList<WordElement> GetNativeNoteElements(WordDocument document, OfficeXmlNode source) {
        var elements = new List<WordElement>();
        foreach (OfficeXmlNode child in source.Elements()) {
            if (child.IsW("p")) {
                // Canonical paragraph views retain bookmark metadata as well as
                // run content; a bare wrapper loses loaded note destinations.
                elements.AddRange(CollapseNativeParagraphElements(
                    document.ConvertParagraphNodeToWordParagraphs(child)));
            } else if (child.IsW("tbl")) elements.Add(document.ReadOnlyTableView(child));
        }
        return elements;
    }

    /// <summary>
    /// Checks the native grid and its effective mapped style before note capture.
    /// Unsupported placement or merged-object/row constraints retain the reported
    /// fallback without rendering a second copy merely to probe admission.
    /// </summary>
    private static bool CanProjectNativeNoteTable(OfficeXmlNode table, WordSection section, WordPdfExportOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts) {
        if (table.HasDescendant(WordXml.W, "tblpPr")) return false;
        if (table.HasDescendant(WordXml.W, "tbl") &&
            !CanProjectNativeContainedTable(section._document.ReadOnlyTableView(table),
                options, defaults, fonts, 0)) return false;
        if (table.DescendantElements(WordXml.W, "textDirection").Any(direction =>
            direction.WAttr("val") is { } value && value != "lrTb")) return false;
        // Block controls and custom row/cell carriers have their own projection
        // rules. A note table currently admits the directly represented grid.
        if (table.Elements().Any(child => !child.IsW("tblPr") && !child.IsW("tblGrid") && !child.IsW("tr")) ||
            table.WElements("tr").Any(row => row.Elements().Any(child => !child.IsW("trPr") && !child.IsW("tc"))) ||
            table.Descendants(WordXml.W, "tc").Any(cell => cell.Elements().Any(child => !child.IsW("tcPr") && !child.IsW("p") && !child.IsW("tbl"))))
            return false;
        var view = section._document.ReadOnlyTableView(table);
        TableLayout layout = TableLayoutCache.GetLayout(view);
        NativeTableStyleDefaults tableDefaults = GetNativeTableStyleDefaults(view, defaults,
            ignoreFallbackTableStyle: options?.PdfOptions?.HasExplicitDefaultTableStyle == true);
        PdfCore.PdfTableStyle style = CreateNativeTableStyle(view, layout.Rows.Count, options,
            GetNativeFootnoteSectionWidth(section, options), defaults, tableDefaults, layout, fonts);
        if (style.Position != null || !style.ConsumesVerticalFlow || style.Caption != null) return false;
        for (int row = 0; row < layout.Rows.Count; row++) {
            options?.CancellationToken.ThrowIfCancellationRequested();
            foreach (WordTableCell cell in layout.Rows[row]) {
                if (IsNativeHorizontalMergeContinuation(cell) || IsNativeVerticalMergeContinuation(cell)) continue;
                int span = GetNativeCellRowSpan(cell);
                if (span <= 1) continue;
                // Resolve header suppression and row defaults through the same
                // table mapper used for painting, rather than raw tblHeader flags.
                if (row < style.HeaderRowCount || row + span > layout.Rows.Count - style.FooterRowCount ||
                    style.CellDataBars?.Keys.Any(key => key.Row == row) == true ||
                    style.CellIcons?.Keys.Any(key => key.Row == row) == true ||
                    Enumerable.Range(row, span).Any(index => HasNativeNoteTableRowConstraint(style, index)) ||
                    EnumerateNativeTableCellParagraphs(cell).Any(paragraph =>
                        EnumerateNativeParagraphImages(paragraph, options?.CancellationToken ?? default).Any())) return false;
            }
        }
        return true;
    }

    private static bool HasNativeNoteTableRowConstraint(PdfCore.PdfTableStyle style, int row) =>
        style.FixedRowHeights != null && row < style.FixedRowHeights.Count && style.FixedRowHeights[row].HasValue ||
        !PdfCore.PdfWriter.GetTableRowAllowBreakAcrossPages(style, row);

    /// <summary>
    /// Native Word does not repeat an authored table header in a continued note story.
    /// Preserve its visual header role and styles while changing only repetition on
    /// the effective projection style; the loaded table and ordinary body policy stay intact.
    /// </summary>
    private static PdfCore.PdfTableStyle CreateNativeNoteTableStyle(WordTable table, PdfCore.PdfTableStyle style,
        NativeDocumentDefaults defaults) {
        if (defaults.NoteContentSection == null ||
            !table.Node().Ancestors().Any(parent => parent.IsW("footnote") || parent.IsW("endnote"))) return style;
        PdfCore.PdfTableStyle noteStyle = style.Clone();
        noteStyle.RepeatHeaderRowCount = 0;
        return noteStyle;
    }
}
