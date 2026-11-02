using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Applies theme alignment to plain numeric values and their otherwise unaligned header labels.</summary>
    private void ApplyDefaultTableNumberAlignment() {
        TableCell? cell = _paragraph.Parent as TableCell;
        if (cell == null || cell.Elements<Paragraph>().Count() != 1 || cell.Elements<Table>().Any()) return;
        Table? table = cell.Ancestors<Table>().FirstOrDefault();
        if (table?.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value != "OfficeIMOModern") return;
        WordParagraphAlignment alignment = _document.EffectiveCreationTheme.TableStyle.NumberAlignment switch {
            OfficeAlignment.Right => WordParagraphAlignment.Right,
            OfficeAlignment.Center => WordParagraphAlignment.Center,
            OfficeAlignment.Start => WordParagraphAlignment.Start,
            OfficeAlignment.End => WordParagraphAlignment.End,
            _ => WordParagraphAlignment.Left
        };
        bool numeric = WordNumericText.IsNumber(Text);
        if (numeric && _paragraph.ParagraphProperties?.Justification == null &&
            _paragraph.ParagraphProperties?.ParagraphStyleId == null && _paragraph.Elements<Run>().Count() == 1)
            ParagraphAlignment = alignment;
        ApplyDefaultNumericHeaderAlignment(table, cell, numeric, alignment);
    }

    /// <summary>Infers plain numeric columns while retaining authored header alignment, styles and rich content.</summary>
    private void ApplyDefaultNumericHeaderAlignment(Table table, TableCell sourceCell, bool numeric,
        WordParagraphAlignment alignment) {
        TableRow? firstRow = WordTableContent.Rows(table).FirstOrDefault();
        if (firstRow == null || !numeric && !ReferenceEquals(sourceCell.Parent, firstRow)) return;
        TableHeader? repeat = firstRow.GetFirstChild<TableRowProperties>()?.GetFirstChild<TableHeader>();
        if (!WordTable.HasFirstRowFormatting(table) && (repeat == null || !WordTableRow.ReadOnOff(repeat))) return;
        if (!WordTableCellSpanResolver.TryGetLogicalColumn(sourceCell, out int column)) return;
        TableCell? header = FindCell(firstRow, column);
        if (header == null || WordTableCellSpanResolver.GetColumnSpan(header) != 1 || header.Elements<Table>().Any()) return;
        Paragraph[] headerParagraphs = header.Elements<Paragraph>().ToArray();
        if (headerParagraphs.Length != 1) return;
        Paragraph paragraph = headerParagraphs[0];
        if (paragraph.ParagraphProperties?.Justification != null || paragraph.ParagraphProperties?.ParagraphStyleId != null ||
            paragraph.Elements<Run>().Count() != 1 || paragraph.ChildElements.Any(element => element is not Run && element is not ParagraphProperties)) return;
        bool sawNumber = false;
        foreach (TableRow row in WordTableContent.Rows(table).Skip(1)) {
            TableCell? value = FindCell(row, column);
            if (value == null || WordTableCellSpanResolver.GetColumnSpan(value) != 1 || value.Elements<Table>().Any() ||
                value.Elements<Paragraph>().Count() != 1) return;
            string text = value.InnerText;
            if (string.IsNullOrWhiteSpace(text)) continue;
            if (!WordNumericText.IsNumber(text)) return;
            sawNumber = true;
        }
        if (sawNumber) new WordParagraph(_document, paragraph).ParagraphAlignment = alignment;

        static TableCell? FindCell(TableRow row, int column) => WordTableContent.Cells(row)
            .FirstOrDefault(candidate => WordTableCellSpanResolver.TryGetLogicalColumn(candidate, out int start) && start == column);
    }
}
